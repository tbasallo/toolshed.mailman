using Amazon;
using Amazon.Runtime;
using Amazon.SimpleEmailV2;
using Amazon.SimpleEmailV2.Model;
using MimeKit;
using MimeKit.Text;
using MimeKit.Utils;

namespace Toolshed.Mailman.AwsSes;

/// <summary>
/// Builds and sends email messages via the Amazon SES (Simple Email Service) v2 API using the supplied
/// <see cref="AwsSesSettings"/>. Messages are built with MimeKit (supporting HTML, attachments and embedded
/// resources) and sent to SES as raw MIME content, so the full feature set of a MIME message is preserved.
/// The service holds message state (subject, recipients, attachments, categories) and a reusable SES client
/// that is created lazily on first send and reused across subsequent sends for efficiency.
/// <para>
/// This type is <b>not thread-safe</b>: it maintains mutable message state and a single shared SES client,
/// so a given instance must not be used concurrently from multiple threads. Register/resolve one instance per
/// logical unit of work and dispose it so the underlying client is released.
/// </para>
/// </summary>
/// <param name="settings">The configuration settings for the SES mail service.</param>
public class AwsSesMailmanService(AwsSesSettings settings) : IDisposable
{
    #region Fields

    private readonly AwsSesSettings _Settings = settings;
    private IAmazonSimpleEmailServiceV2? _Client;
    private bool _Disposed;

    #endregion

    #region Message Properties

    /// <summary>
    /// The subject line of the email message.
    /// </summary>
    public string? Subject { get; set; }

    /// <summary>
    /// The importance level of the email message. Defaults to Normal.
    /// </summary>
    public MessageImportance Importance { get; set; } = MessageImportance.Normal;

    /// <summary>
    /// The priority level of the email message. Defaults to Normal.
    /// </summary>
    public MessagePriority Priority { get; set; } = MessagePriority.Normal;

    /// <summary>
    /// When true, the service automatically calls <see cref="Reset(bool)"/> after a message is sent,
    /// clearing the subject, recipients, attachments, and categories.
    /// </summary>
    public bool IsResetedAfterMessageSent { get; set; }

    /// <summary>
    /// A comma delimited list of SES message tags to apply to the message. This overrides any categories
    /// provided by the settings. To keep the settings categories, leave this null.
    /// </summary>
    public string? Categories { get; set; }

    #endregion

    #region Recipients

    MailboxAddress? _From;
    /// <summary>
    /// The primary sender address. If not set, defaults to the FromAddress from settings.
    /// Returns null when neither an explicit From nor a settings FromAddress is available.
    /// </summary>
    public MailboxAddress? From
    {
        get
        {
            if (_From == null && !string.IsNullOrWhiteSpace(_Settings.FromAddress))
            {
                if (!string.IsNullOrWhiteSpace(_Settings.FromDisplayName))
                {
                    _From = new MailboxAddress(_Settings.FromDisplayName, _Settings.FromAddress);
                }
                else
                {
                    _From = MailboxAddress.Parse(_Settings.FromAddress);
                }
            }

            return _From;
        }
        set { _From = value; }
    }

    List<MailboxAddress>? _To;
    /// <summary>
    /// The list of primary recipients (To addresses).
    /// </summary>
    public List<MailboxAddress> To
    {
        get
        {
            _To ??= [];
            return _To;
        }
        set { _To = value; }
    }

    List<MailboxAddress>? _CC;
    /// <summary>
    /// The list of CC (carbon copy) recipients.
    /// </summary>
    public List<MailboxAddress> CC
    {
        get
        {
            _CC ??= [];
            return _CC;
        }
        set { _CC = value; }
    }

    List<MailboxAddress>? _Bcc;
    /// <summary>
    /// The list of BCC (blind carbon copy) recipients.
    /// </summary>
    public List<MailboxAddress> Bcc
    {
        get
        {
            _Bcc ??= [];
            return _Bcc;
        }
        set { _Bcc = value; }
    }

    /// <summary>
    /// Adds a recipient email address to the To list.
    /// </summary>
    /// <param name="email">The email address to add.</param>
    public void AddTo(string email) => To.Add(MailboxAddress.Parse(email));

    /// <summary>
    /// Adds a recipient with display name and email address to the To list.
    /// </summary>
    /// <param name="name">The display name of the recipient.</param>
    /// <param name="email">The email address of the recipient.</param>
    public void AddTo(string name, string email) => To.Add(new MailboxAddress(name, email));

    /// <summary>
    /// Adds a CC recipient email address.
    /// </summary>
    /// <param name="email">The email address to add.</param>
    public void AddCc(string email) => CC.Add(MailboxAddress.Parse(email));

    /// <summary>
    /// Adds a CC recipient with display name and email address.
    /// </summary>
    /// <param name="name">The display name of the recipient.</param>
    /// <param name="email">The email address of the recipient.</param>
    public void AddCc(string name, string email) => CC.Add(new MailboxAddress(name, email));

    /// <summary>
    /// Adds a BCC recipient email address.
    /// </summary>
    /// <param name="email">The email address to add.</param>
    public void AddBcc(string email) => Bcc.Add(MailboxAddress.Parse(email));

    /// <summary>
    /// Adds a BCC recipient with display name and email address.
    /// </summary>
    /// <param name="name">The display name of the recipient.</param>
    /// <param name="email">The email address of the recipient.</param>
    public void AddBcc(string name, string email) => Bcc.Add(new MailboxAddress(name, email));

    /// <summary>
    /// Adds multiple recipient email addresses to the To list.
    /// </summary>
    /// <param name="emails">The collection of email addresses to add.</param>
    public void AddTo(IEnumerable<string> emails) => To.AddRange(emails.Select(MailboxAddress.Parse));

    /// <summary>
    /// Adds multiple CC recipient email addresses.
    /// </summary>
    /// <param name="emails">The collection of email addresses to add.</param>
    public void AddCc(IEnumerable<string> emails) => CC.AddRange(emails.Select(MailboxAddress.Parse));

    /// <summary>
    /// Adds multiple BCC recipient email addresses.
    /// </summary>
    /// <param name="emails">The collection of email addresses to add.</param>
    public void AddBcc(IEnumerable<string> emails) => Bcc.AddRange(emails.Select(MailboxAddress.Parse));

    /// <summary>
    /// Clears the To, CC, and BCC recipient lists.
    /// </summary>
    public void ClearRecipients()
    {
        _To?.Clear();
        _CC?.Clear();
        _Bcc?.Clear();
    }

    #endregion

    #region State Management

    /// <summary>
    /// Clears the SUBJECT, TO, CC, BCC, categories, attachments, and linked resources. Can optionally
    /// reset the FROM. Settings are not affected.
    /// </summary>
    /// <param name="resetFrom">Specifies whether the FROM property should also be cleared.</param>
    public void Reset(bool resetFrom = false)
    {
        Subject = null;
        ClearRecipients();
        DisposeAttachmentsAndResources();

        Categories = null;

        if (resetFrom)
        {
            _From = null;
        }
    }

    /// <summary>
    /// A quick sanity check that verifies a sender, at least one recipient, and either a configured
    /// region or a service URL override are present.
    /// </summary>
    public bool IsValidToSend()
    {
        if (_From is null && string.IsNullOrWhiteSpace(_Settings.FromAddress))
        {
            return false;
        }

        var hasRecipient = (_To?.Count > 0) || (_CC?.Count > 0) || (_Bcc?.Count > 0);
        if (!hasRecipient)
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(_Settings.Region) && string.IsNullOrWhiteSpace(_Settings.ServiceUrl))
        {
            return false;
        }

        return true;
    }

    #endregion

    #region Attachments and Linked Resources

    /// <summary>
    /// The collection of file attachments to include with the email.
    /// </summary>
    public AttachmentCollection Attachments { get; set; } = new AttachmentCollection(false);

    /// <summary>
    /// A collection of linked resources (embedded images) that can be referenced in HTML body using cid: URLs.
    /// </summary>
    public List<MimeEntity> LinkedResources { get; set; } = [];

    /// <summary>
    /// Adds a linked resource (embedded image) from a file path. Returns the Content-Id to use in HTML (e.g., src="cid:{contentId}").
    /// </summary>
    /// <param name="filePath">The path to the image file.</param>
    /// <param name="contentType">Optional content type (e.g., "image/png"). If not provided, it will be inferred from the file extension.</param>
    /// <returns>The Content-Id to reference in HTML.</returns>
    public string AddLinkedResource(string filePath, string? contentType = null)
    {
        var mimeType = contentType ?? MimeTypes.GetMimeType(filePath);

        var resource = new MimePart(mimeType)
        {
            Content = new MimeContent(File.OpenRead(filePath)),
            ContentDisposition = new ContentDisposition(ContentDisposition.Inline),
            ContentTransferEncoding = ContentEncoding.Base64,
            FileName = Path.GetFileName(filePath),
            ContentId = MimeUtils.GenerateMessageId()
        };
        LinkedResources.Add(resource);

        return resource.ContentId;
    }

    /// <summary>
    /// Adds a linked resource (embedded image) from a stream. Returns the Content-Id to use in HTML (e.g., src="cid:{contentId}").
    /// </summary>
    /// <param name="stream">The stream containing the image data.</param>
    /// <param name="fileName">The file name for the resource.</param>
    /// <param name="contentType">The content type (e.g., "image/png").</param>
    /// <returns>The Content-Id to reference in HTML.</returns>
    public string AddLinkedResource(Stream stream, string fileName, string contentType)
    {
        var resource = new MimePart(contentType)
        {
            Content = new MimeContent(stream),
            ContentDisposition = new ContentDisposition(ContentDisposition.Inline),
            ContentTransferEncoding = ContentEncoding.Base64,
            FileName = fileName,
            ContentId = MimeUtils.GenerateMessageId()
        };
        LinkedResources.Add(resource);

        return resource.ContentId;
    }

    /// <summary>
    /// Adds a linked resource (embedded image) from a byte array. Returns the Content-Id to use in HTML (e.g., src="cid:{contentId}").
    /// </summary>
    /// <param name="data">The byte array containing the image data.</param>
    /// <param name="fileName">The file name for the resource.</param>
    /// <param name="contentType">The content type (e.g., "image/png").</param>
    /// <returns>The Content-Id to reference in HTML.</returns>
    public string AddLinkedResource(byte[] data, string fileName, string contentType)
    {
        var stream = new MemoryStream(data);
        return AddLinkedResource(stream, fileName, contentType);
    }

    private void DisposeAttachmentsAndResources()
    {
        Attachments.Clear(true);
        LinkedResources.Clear();
    }

    #endregion

    #region Message Building

    /// <summary>
    /// Creates a MimeMessage from the current service state with the specified body content.
    /// </summary>
    /// <param name="body">The body content of the email.</param>
    /// <param name="isHtml">Whether the body content is HTML formatted. Defaults to true.</param>
    /// <returns>A configured MimeMessage ready for sending.</returns>
    public MimeMessage GetMessage(string body, bool isHtml = true)
    {
        var message = new MimeMessage
        {
            Subject = Subject,
            Priority = Priority,
            Importance = Importance
        };

        if (Attachments.Count > 0 || LinkedResources.Count > 0)
        {
            var builder = new BodyBuilder();
            if (isHtml)
            {
                builder.HtmlBody = body;
            }
            else
            {
                builder.TextBody = body;
            }

            foreach (var item in Attachments)
            {
                builder.Attachments.Add(item);
            }

            foreach (var linkedResource in LinkedResources)
            {
                builder.LinkedResources.Add(linkedResource);
            }

            message.Body = builder.ToMessageBody();
        }
        else
        {
            message.Body = new TextPart(isHtml ? TextFormat.Html : TextFormat.Text) { Text = body };
        }

        var from = From
            ?? throw new InvalidOperationException($"A sender is required. Set {nameof(From)} or {nameof(AwsSesSettings)}.{nameof(AwsSesSettings.FromAddress)}.");
        message.From.Add(from);

        if (_To != null && _To.Count > 0)
        {
            message.To.AddRange(To);
        }
        if (_CC != null && _CC.Count > 0)
        {
            message.Cc.AddRange(CC);
        }
        if (_Bcc != null && _Bcc.Count > 0)
        {
            message.Bcc.AddRange(Bcc);
        }

        return message;
    }

    #endregion

    #region Sending

    /// <summary>
    /// Builds a message from the current service state and sends it through the Amazon SES API. The body
    /// can be either plain text or HTML.
    /// </summary>
    /// <param name="body">The body content of the email.</param>
    /// <param name="isHtml">Whether the body content is HTML formatted. Defaults to true.</param>
    /// <param name="cancellationToken">A token to cancel the send operation.</param>
    /// <returns>The SES <see cref="SendEmailResponse"/> containing the accepted message id.</returns>
    public Task<SendEmailResponse> SendMessageAsync(string body, bool isHtml = true, CancellationToken cancellationToken = default)
    {
        var message = GetMessage(body, isHtml);
        return SendCoreAsync(message, disposeMessage: true, cancellationToken);
    }

    /// <summary>
    /// Sends a pre-configured <see cref="MimeMessage"/> through the Amazon SES API as raw content.
    /// </summary>
    /// <param name="mailMessage">The MimeMessage to send.</param>
    /// <param name="disposeMessage">Disposes the MimeMessage after sending. If you're not using the message after this send, this should be true.</param>
    /// <param name="cancellationToken">A token to cancel the send operation.</param>
    /// <returns>The SES <see cref="SendEmailResponse"/> containing the accepted message id.</returns>
    public Task<SendEmailResponse> SendMessageAsync(MimeMessage mailMessage, bool disposeMessage = false, CancellationToken cancellationToken = default)
    {
        return SendCoreAsync(mailMessage, disposeMessage, cancellationToken);
    }

    /// <summary>
    /// Sends multiple pre-configured messages through the Amazon SES API over the reusable client. The
    /// client is created once and reused across all messages in the batch.
    /// </summary>
    /// <param name="mailMessages">The messages to send.</param>
    /// <param name="disposeMessages">Disposes each MimeMessage after it is sent. If you're not using the messages after this send, this should be true.</param>
    /// <param name="cancellationToken">A token to cancel the send operation.</param>
    /// <returns>The SES responses for each message, in order.</returns>
    public async Task<IReadOnlyList<SendEmailResponse>> SendMessagesAsync(IEnumerable<MimeMessage> mailMessages, bool disposeMessages = false, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(mailMessages);

        var responses = new List<SendEmailResponse>();
        foreach (var message in mailMessages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            responses.Add(await SendCoreAsync(message, disposeMessages, cancellationToken).ConfigureAwait(false));
        }

        return responses;
    }

    /// <summary>
    /// Serializes the message to raw MIME content and sends it through the SES API, applying the
    /// configuration set and message tags, and optionally resetting the service state.
    /// </summary>
    private async Task<SendEmailResponse> SendCoreAsync(MimeMessage mailMessage, bool disposeMessage, CancellationToken cancellationToken)
    {
        try
        {
            var client = GetClient();

            using var stream = new MemoryStream();
            await mailMessage.WriteToAsync(stream, cancellationToken).ConfigureAwait(false);
            stream.Position = 0;

            var request = new SendEmailRequest
            {
                Content = new EmailContent
                {
                    Raw = new RawMessage { Data = stream }
                }
            };

            if (!string.IsNullOrWhiteSpace(_Settings.ConfigurationSetName))
            {
                request.ConfigurationSetName = _Settings.ConfigurationSetName;
            }

            if (!string.IsNullOrWhiteSpace(_Settings.FeedbackForwardingEmailAddress))
            {
                request.FeedbackForwardingEmailAddress = _Settings.FeedbackForwardingEmailAddress;
            }

            var tags = BuildTags();
            if (tags.Count > 0)
            {
                request.EmailTags = tags;
            }

            return await client.SendEmailAsync(request, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            if (IsResetedAfterMessageSent)
            {
                Reset();
            }

            if (disposeMessage)
            {
                mailMessage.Dispose();
            }
        }
    }

    #endregion

    #region Client Management

    /// <summary>
    /// Returns the reusable SES client, creating it lazily on first use from <see cref="AwsSesSettings"/>.
    /// The client is owned by this service and disposed when the service is disposed.
    /// </summary>
    private IAmazonSimpleEmailServiceV2 GetClient()
    {
        ObjectDisposedException.ThrowIf(_Disposed, this);

        if (_Client != null)
        {
            return _Client;
        }

        var config = new AmazonSimpleEmailServiceV2Config();

        if (!string.IsNullOrWhiteSpace(_Settings.ServiceUrl))
        {
            config.ServiceURL = _Settings.ServiceUrl;
        }
        else
        {
            var region = _Settings.GetRegionEndpoint()
                ?? throw new InvalidOperationException($"Either {nameof(AwsSesSettings.Region)} or {nameof(AwsSesSettings.ServiceUrl)} must be configured.");
            config.RegionEndpoint = region;
        }

        if (_Settings.Timeout.HasValue)
        {
            config.Timeout = TimeSpan.FromMilliseconds(_Settings.Timeout.Value);
        }

        if (!string.IsNullOrWhiteSpace(_Settings.AccessKeyId) && !string.IsNullOrWhiteSpace(_Settings.SecretAccessKey))
        {
            AWSCredentials credentials = string.IsNullOrWhiteSpace(_Settings.SessionToken)
                ? new BasicAWSCredentials(_Settings.AccessKeyId, _Settings.SecretAccessKey)
                : new SessionAWSCredentials(_Settings.AccessKeyId, _Settings.SecretAccessKey, _Settings.SessionToken);

            _Client = new AmazonSimpleEmailServiceV2Client(credentials, config);
        }
        else
        {
            // Fall back to the AWS SDK default credential resolution chain.
            _Client = new AmazonSimpleEmailServiceV2Client(config);
        }

        return _Client;
    }

    #endregion

    #region Helpers

    /// <summary>
    /// Builds the list of SES message tags from the instance <see cref="Categories"/> (which takes
    /// precedence) or the settings categories. Each comma delimited entry becomes a tag with value "true".
    /// </summary>
    private List<MessageTag> BuildTags()
    {
        var source = !string.IsNullOrWhiteSpace(Categories) ? Categories : _Settings.Categories;
        if (string.IsNullOrWhiteSpace(source))
        {
            return [];
        }

        var items = source.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        return [.. items.Select(name => new MessageTag { Name = name, Value = "true" })];
    }

    #endregion

    #region IDisposable

    /// <summary>
    /// Disposes the reusable SES client held by this service.
    /// </summary>
    public void Dispose()
    {
        if (_Disposed)
        {
            return;
        }

        _Disposed = true;
        _Client?.Dispose();
        _Client = null;
        DisposeAttachmentsAndResources();

        GC.SuppressFinalize(this);
    }

    #endregion
}
