using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using MimeKit.IO;
using MimeKit.Text;
using MimeKit.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Toolshed.Mailman;

/// <summary>
/// Builds and sends email messages via MailKit/MimeKit using the supplied <see cref="MailmanSettings"/>.
/// The service holds message state (subject, recipients, attachments, categories) and a reusable SMTP
/// connection that is created lazily on first send and reused across subsequent sends for efficiency.
/// <para>
/// This type is <b>not thread-safe</b>: it maintains mutable message state and a single shared SMTP
/// client, so a given instance must not be used concurrently from multiple threads. Register/resolve
/// one instance per logical unit of work and dispose it (preferably via <c>await using</c>) so the
/// underlying SMTP connection is released.
/// </para>
/// </summary>
/// <param name="settings">The configuration settings for the mail service</param>
public class MailmanService(MailmanSettings settings) : IDisposable, IAsyncDisposable
{
    private readonly MailmanSettings _settings = settings;
    private SmtpClient? _smtpClient;
    private bool _disposed;

    /// <summary>
    /// The subject line of the email message
    /// </summary>
    public string? Subject { get; set; }

    /// <summary>
    /// An optional view name for callers that render the email body themselves (for example a Razor
    /// view/template engine) before passing the result to <see cref="SendMessageAsync(string, bool, CancellationToken)"/>.
    /// This value is not used internally by the service and has no effect on sending.
    /// </summary>
    public string? ViewName { get; set; }

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
    /// clearing the subject, recipients, attachments, and categories. The reusable SMTP connection is
    /// left intact so it can be reused by the next send.
    /// </summary>
    public bool IsResetedAfterMessageSent { get; set; }

    /// <summary>
    /// A list of categories to add to X-SMTPAPI header. This will overwrite any categories provided by the settings. To keep the current categories, use the AddCategory() method
    /// </summary>
    public string? Categories { get; set; }

    private string? InternalCategories { get; set; } = settings.Categories;


    /// <summary>
    /// Adds a recipient email address to the To list
    /// </summary>
    /// <param name="email">The email address to add</param>
    public void AddTo(string email)
    {
        To.Add(MailboxAddress.Parse(email));
    }

    /// <summary>
    /// Adds a recipient with display name and email address to the To list
    /// </summary>
    /// <param name="name">The display name of the recipient</param>
    /// <param name="email">The email address of the recipient</param>
    public void AddTo(string name, string email)
    {
        To.Add(new MailboxAddress(name, email));
    }

    /// <summary>
    /// Adds a CC recipient email address
    /// </summary>
    /// <param name="email">The email address to add</param>
    public void AddCc(string email)
    {
        CC.Add(MailboxAddress.Parse(email));
    }

    /// <summary>
    /// Adds a CC recipient with display name and email address
    /// </summary>
    /// <param name="name">The display name of the recipient</param>
    /// <param name="email">The email address of the recipient</param>
    public void AddCc(string name, string email)
    {
        CC.Add(new MailboxAddress(name, email));
    }

    /// <summary>
    /// Adds a BCC recipient email address
    /// </summary>
    /// <param name="email">The email address to add</param>
    public void AddBcc(string email)
    {
        Bcc.Add(MailboxAddress.Parse(email));
    }

    /// <summary>
    /// Adds a BCC recipient with display name and email address
    /// </summary>
    /// <param name="name">The display name of the recipient</param>
    /// <param name="email">The email address of the recipient</param>
    public void AddBcc(string name, string email)
    {
        Bcc.Add(new MailboxAddress(name, email));
    }

    /// <summary>
    /// Adds multiple recipient email addresses to the To list
    /// </summary>
    /// <param name="emails">The collection of email addresses to add</param>
    public void AddTo(IEnumerable<string> emails)
    {
        To.AddRange(emails.Select(MailboxAddress.Parse));
    }

    /// <summary>
    /// Adds multiple CC recipient email addresses
    /// </summary>
    /// <param name="emails">The collection of email addresses to add</param>
    public void AddCc(IEnumerable<string> emails)
    {
        CC.AddRange(emails.Select(MailboxAddress.Parse));
    }

    /// <summary>
    /// Adds multiple BCC recipient email addresses
    /// </summary>
    /// <param name="emails">The collection of email addresses to add</param>
    public void AddBcc(IEnumerable<string> emails)
    {
        Bcc.AddRange(emails.Select(MailboxAddress.Parse));
    }

    MailboxAddress? _From;
    /// <summary>
    /// The primary sender address. If not set, defaults to the FromAddress from settings.
    /// Returns null when neither an explicit From nor a settings FromAddress is available.
    /// </summary>
    public MailboxAddress? From
    {
        get
        {
            if (_From == null && !string.IsNullOrWhiteSpace(_settings.FromAddress))
            {
                if (!string.IsNullOrWhiteSpace(_settings.FromDisplayName))
                {
                    _From = new MailboxAddress(_settings.FromDisplayName, _settings.FromAddress);
                }
                else
                {
                    _From = MailboxAddress.Parse(_settings.FromAddress);
                }
            }

            return _From;
        }
        set { _From = value; }
    }

    List<MailboxAddress>? _To;
    /// <summary>
    /// The list of primary recipients (To addresses)
    /// </summary>
    public List<MailboxAddress> To
    {
        get
        {
            _To ??= [];
            return _To;
        }
        set
        {
            _To = value;
        }
    }

    List<MailboxAddress>? _CC;
    /// <summary>
    /// The list of CC (carbon copy) recipients
    /// </summary>
    public List<MailboxAddress> CC
    {
        get
        {
            _CC ??= [];
            return _CC;
        }
        set
        {
            _CC = value;
        }
    }

    List<MailboxAddress>? _Bcc;
    /// <summary>
    /// The list of BCC (blind carbon copy) recipients
    /// </summary>
    public List<MailboxAddress> Bcc
    {
        get
        {
            _Bcc ??= [];
            return _Bcc;
        }
        set
        {
            _Bcc = value;
        }
    }


    /// <summary>
    /// Clears all recipients (To, CC, and BCC) from the message
    /// </summary>
    public void ClearRecipients()
    {
        _To?.Clear();
        _CC?.Clear();
        _Bcc?.Clear();
    }

    /// <summary>
    /// Clears the SUBJECT, TO, CC, and BCC properties. Can optionally reset the FROM. Settings are not affected.
    /// </summary>
    /// <param name="resetFrom">Specifies whether the FROM property should also be cleared</param>
    public void Reset(bool resetFrom = false)
    {
        Subject = null;
        ClearRecipients();
        DisposeAttachmentsAndResources();

        Categories = null;
        InternalCategories = _settings.Categories;

        if (resetFrom)
        {
            _From = null;
        }
    }

    /// <summary>
    /// There are no checks or exception catching, .NET will that for you. However, if you want a quick sanity check to make sure
    /// typical issues are looked at, we can do that. We check recipients, from  and if Delivery is network, we check for a host
    /// </summary>
    public bool IsValidToSend()
    {
        if (_From is null && string.IsNullOrWhiteSpace(_settings.FromAddress))
        {
            return false;
        }

        var hasRecipient = (_To?.Count > 0) || (_CC?.Count > 0) || (_Bcc?.Count > 0);
        if (!hasRecipient)
        {
            return false;
        }
        if (_settings.DeliveryMethod == System.Net.Mail.SmtpDeliveryMethod.Network && string.IsNullOrWhiteSpace(_settings.Host))
        {
            return false;
        }
        if (_settings.DeliveryMethod == System.Net.Mail.SmtpDeliveryMethod.SpecifiedPickupDirectory && string.IsNullOrWhiteSpace(_settings.PickupDirectoryLocation))
        {
            return false;
        }

        return true;
    }

    /// <summary>
    /// The collection of file attachments to include with the email
    /// </summary>
    public AttachmentCollection Attachments { get; set; } = new AttachmentCollection(false);

    /// <summary>
    /// A collection of linked resources (embedded images) that can be referenced in HTML body using cid: URLs
    /// </summary>
    public List<MimeEntity> LinkedResources { get; set; } = [];

    /// <summary>
    /// Adds a linked resource (embedded image) from a file path. Returns the Content-Id to use in HTML (e.g., src="cid:{contentId}")
    /// </summary>
    /// <param name="filePath">The path to the image file</param>
    /// <param name="contentType">Optional content type (e.g., "image/png"). If not provided, it will be inferred from the file extension.</param>
    /// <returns>The Content-Id to reference in HTML</returns>
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
    /// Adds a linked resource (embedded image) from a stream. Returns the Content-Id to use in HTML (e.g., src="cid:{contentId}")
    /// </summary>
    /// <param name="stream">The stream containing the image data</param>
    /// <param name="fileName">The file name for the resource</param>
    /// <param name="contentType">The content type (e.g., "image/png")</param>
    /// <returns>The Content-Id to reference in HTML</returns>
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
    /// Adds a linked resource (embedded image) from a byte array. Returns the Content-Id to use in HTML (e.g., src="cid:{contentId}")
    /// </summary>
    /// <param name="data">The byte array containing the image data</param>
    /// <param name="fileName">The file name for the resource</param>
    /// <param name="contentType">The content type (e.g., "image/png")</param>
    /// <returns>The Content-Id to reference in HTML</returns>
    public string AddLinkedResource(byte[] data, string fileName, string contentType)
    {
        var stream = new MemoryStream(data);
        return AddLinkedResource(stream, fileName, contentType);
    }

    /// <summary>
    /// Adds a category to the email for tracking purposes. Categories are appended to existing categories.
    /// </summary>
    /// <param name="category">The category name to add</param>
    public void AddCategory(string category)
    {
        if (string.IsNullOrWhiteSpace(InternalCategories))
        {
            InternalCategories = category;
        }
        else
        {
            InternalCategories += "," + category;
        }

    }



    /// <summary>
    /// Creates a MimeMessage from the current service state with the specified body content
    /// </summary>
    /// <param name="body">The body content of the email</param>
    /// <param name="isHtml">Whether the body content is HTML formatted. Defaults to true.</param>
    /// <returns>A configured MimeMessage ready for sending</returns>
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
            if (isHtml)
            {
                message.Body = new TextPart(TextFormat.Html) { Text = body };
            }
            else
            {
                message.Body = new TextPart(TextFormat.Text) { Text = body };
            }
        }


        var from = From;
        if (from != null)
        {
            message.From.Add(from);
        }

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

    /// <summary>
    /// Sends a message using the SMTP client. The body can be either plain text or HTML.
    /// </summary>
    /// <param name="body">The body content of the email</param>
    /// <param name="isHtml">Whether the body content is HTML formatted. Defaults to true.</param>
    /// <param name="cancellationToken">A token to cancel the send operation</param>
    /// <returns>A task representing the asynchronous send operation</returns>
    public Task SendMessageAsync(string body, bool isHtml = true, CancellationToken cancellationToken = default)
    {
        var mailMessage = GetMessage(body, isHtml);
        return SendCoreAsync(mailMessage, disposeMessage: true, cancellationToken);
    }


    /// <summary>
    /// Sends a pre-configured MimeMessage over the service's reusable SMTP connection. The connection is
    /// established on first use and reused by later sends; if it has gone stale it is transparently
    /// rebuilt and the send is retried once.
    /// </summary>
    /// <param name="mailMessage">The MimeMessage to send</param>
    /// <param name="disposeMessage">Disposes the MimeMessage after sending. If you're not using the message after this send, this should be true.</param>
    /// <param name="cancellationToken">A token to cancel the send operation</param>
    /// <returns>A task representing the asynchronous send operation</returns>
    public Task SendMessageAsync(MimeMessage mailMessage, bool disposeMessage = false, CancellationToken cancellationToken = default)
    {
        return SendCoreAsync(mailMessage, disposeMessage: disposeMessage, cancellationToken);
    }

    /// <summary>
    /// Sends multiple pre-configured messages over the service's reusable SMTP connection. This is more
    /// efficient than a naive loop because the connection and authentication handshake are performed only
    /// once and the connection is kept open across all messages in the batch. If the connection drops
    /// mid-batch it is transparently rebuilt and the failing message is retried once. When the delivery
    /// method is a pickup directory, each message is written to disk instead.
    /// </summary>
    /// <param name="mailMessages">The messages to send</param>
    /// <param name="disposeMessages">Disposes each MimeMessage after it is sent. If you're not using the messages after this send, this should be true.</param>
    /// <param name="cancellationToken">A token to cancel the send operation</param>
    /// <returns>A task representing the asynchronous send operation</returns>
    public async Task SendMessagesAsync(IEnumerable<MimeMessage> mailMessages, bool disposeMessages = false, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(mailMessages);

        if (_settings.DeliveryMethod == System.Net.Mail.SmtpDeliveryMethod.SpecifiedPickupDirectory)
        {
            if (string.IsNullOrWhiteSpace(_settings.PickupDirectoryLocation))
            {
                throw new InvalidOperationException($"{nameof(MailmanSettings.PickupDirectoryLocation)} must be set when {nameof(MailmanSettings.DeliveryMethod)} is {nameof(System.Net.Mail.SmtpDeliveryMethod.SpecifiedPickupDirectory)}.");
            }

            foreach (var message in mailMessages)
            {
                cancellationToken.ThrowIfCancellationRequested();
                ApplyCategories(message);
                SaveToPickupDirectory(message, _settings.PickupDirectoryLocation);
                if (disposeMessages)
                {
                    message.Dispose();
                }
            }
            return;
        }

        // Establish (or reuse) the connection up front so a bad configuration fails fast.
        await GetConnectedClientAsync(cancellationToken).ConfigureAwait(false);
        foreach (var message in mailMessages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ApplyCategories(message);
            await SendWithReconnectAsync(message, cancellationToken).ConfigureAwait(false);
            if (disposeMessages)
            {
                message.Dispose();
            }
        }
    }

    /// <summary>
    /// Applies the X-SMTPAPI category header (when categories are configured), then either writes the
    /// message to the pickup directory or sends it over the reusable SMTP connection (with a single
    /// reconnect-and-retry on a stale connection), and optionally resets the service state.
    /// </summary>
    private async Task SendCoreAsync(MimeMessage mailMessage, bool disposeMessage, CancellationToken cancellationToken)
    {
        try
        {
            ApplyCategories(mailMessage);

            if (_settings.DeliveryMethod == System.Net.Mail.SmtpDeliveryMethod.SpecifiedPickupDirectory)
            {
                if (string.IsNullOrWhiteSpace(_settings.PickupDirectoryLocation))
                {
                    throw new InvalidOperationException($"{nameof(MailmanSettings.PickupDirectoryLocation)} must be set when {nameof(MailmanSettings.DeliveryMethod)} is {nameof(System.Net.Mail.SmtpDeliveryMethod.SpecifiedPickupDirectory)}.");
                }

                SaveToPickupDirectory(mailMessage, _settings.PickupDirectoryLocation);
                return;
            }

            await SendWithReconnectAsync(mailMessage, cancellationToken).ConfigureAwait(false);
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

    /// <summary>
    /// Returns a connected (and, when credentials are configured, authenticated) <see cref="SmtpClient"/>.
    /// The client is created lazily on first use and reused across subsequent sends so that the connection
    /// and authentication handshake are not repeated for every message. If a previously created client has
    /// been disconnected (for example by the server), it is transparently reconnected. The client is owned
    /// by this service and is disposed when the service is disposed.
    /// </summary>
    private async Task<SmtpClient> GetConnectedClientAsync(CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        _smtpClient ??= new SmtpClient();

        if (_settings.Timeout.HasValue)
        {
            _smtpClient.Timeout = _settings.Timeout.Value;
        }

        _smtpClient.CheckCertificateRevocation = _settings.CheckCertificateRevocation;

        if (!_smtpClient.IsConnected)
        {
            await _smtpClient.ConnectAsync(_settings.Host, _settings.Port, _settings.SecureSocketOptions, cancellationToken).ConfigureAwait(false);
        }

        if (!_smtpClient.IsAuthenticated && !string.IsNullOrWhiteSpace(_settings.UserName) && !string.IsNullOrWhiteSpace(_settings.Password))
        {
            await _smtpClient.AuthenticateAsync(new System.Net.NetworkCredential(_settings.UserName, _settings.Password), cancellationToken).ConfigureAwait(false);
        }

        return _smtpClient;
    }

    /// <summary>
    /// Sends a single message over the persistent client. If the send fails because the pooled
    /// connection has gone stale (for example the server dropped an idle connection), the client is
    /// torn down, a fresh connection is established, and the send is retried exactly once.
    /// </summary>
    private async Task SendWithReconnectAsync(MimeMessage mailMessage, CancellationToken cancellationToken)
    {
        var sm = await GetConnectedClientAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            await sm.SendAsync(mailMessage, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex) when (!cancellationToken.IsCancellationRequested && IsTransientConnectionException(ex))
        {
            // The pooled connection was likely stale. Rebuild it and retry the send once.
            DisposeClient();
            var retryClient = await GetConnectedClientAsync(cancellationToken).ConfigureAwait(false);
            await retryClient.SendAsync(mailMessage, cancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Determines whether an exception indicates the persistent connection is no longer usable and a
    /// reconnect-and-retry is warranted.
    /// </summary>
    private static bool IsTransientConnectionException(Exception ex)
    {
        return ex is MailKit.ServiceNotConnectedException
            or MailKit.Net.Smtp.SmtpProtocolException
            or System.IO.IOException
            or System.Net.Sockets.SocketException;
    }

    /// <summary>
    /// Disposes and clears the persistent SMTP client so the next send establishes a fresh connection.
    /// </summary>
    private void DisposeClient()
    {
        if (_smtpClient != null)
        {
            _smtpClient.Dispose();
            _smtpClient = null;
        }
    }

    /// <summary>
    /// Resets the reusable SMTP connection by disposing the current client. The next send will lazily
    /// establish (and re-authenticate) a brand new connection. Use this to force a clean reconnect, for
    /// example after changing network conditions or when a long-lived connection may have gone stale.
    /// Prefer <see cref="ResetConnectionAsync(CancellationToken)"/> so the connection can be closed gracefully.
    /// </summary>
    public void ResetConnection()
    {
        DisposeClient();
    }

    /// <summary>
    /// Asynchronously resets the reusable SMTP connection. When connected, the client is gracefully
    /// disconnected (SMTP QUIT) and disposed; the next send will lazily establish (and re-authenticate)
    /// a brand new connection.
    /// </summary>
    /// <param name="cancellationToken">A token to cancel the graceful disconnect.</param>
    /// <returns>A task representing the asynchronous reset operation.</returns>
    public async Task ResetConnectionAsync(CancellationToken cancellationToken = default)
    {
        if (_smtpClient == null)
        {
            return;
        }

        if (_smtpClient.IsConnected)
        {
            try
            {
                await _smtpClient.DisconnectAsync(true, cancellationToken).ConfigureAwait(false);
            }
            catch
            {
                // Best-effort graceful disconnect; the client is disposed regardless below.
            }
        }

        _smtpClient.Dispose();
        _smtpClient = null;
    }

    /// <summary>
    /// Adds the X-SMTPAPI header with a properly formatted JSON array of categories when any are configured.
    /// The instance <see cref="Categories"/> takes precedence over categories provided via settings.
    /// </summary>
    private void ApplyCategories(MimeMessage mailMessage)
    {
        var categories = !string.IsNullOrWhiteSpace(Categories) ? Categories : InternalCategories;

        // Remove any previously applied header so re-sending the same message does not duplicate it.
        mailMessage.Headers.RemoveAll("X-SMTPAPI");

        if (string.IsNullOrWhiteSpace(categories))
        {
            return;
        }

        var items = categories
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        if (items.Length == 0)
        {
            return;
        }

        var payload = System.Text.Json.JsonSerializer.Serialize(new { category = items });
        mailMessage.Headers.Add("X-SMTPAPI", payload);
    }

    private static void SaveToPickupDirectory(MimeMessage message, string pickupDirectory)
    {
        do
        {
            // Generate a random file name to save the message to.
            var path = Path.Combine(pickupDirectory, Guid.NewGuid().ToString() + ".eml");
            Stream stream;

            try
            {
                // Attempt to create the new file.
                stream = File.Open(path, FileMode.CreateNew);
            }
            catch (IOException)
            {
                // If the file already exists, try again with a new Guid.
                if (File.Exists(path))
                    continue;

                // Otherwise, fail immediately since it probably means that there is
                // no graceful way to recover from this error.
                throw;
            }

            try
            {
                using (stream)
                {
                    // IIS pickup directories expect the message to be "byte-stuffed"
                    // which means that lines beginning with "." need to be escaped
                    // by adding an extra "." to the beginning of the line.
                    //
                    // Use an SmtpDataFilter "byte-stuff" the message as it is written
                    // to the file stream. This is the same process that an SmtpClient
                    // would use when sending the message in a `DATA` command.
                    using var filtered = new FilteredStream(stream);
                    filtered.Add(new SmtpDataFilter());

                    // Make sure to write the message in DOS (<CR><LF>) format.
                    var options = FormatOptions.Default.Clone();
                    options.NewLineFormat = NewLineFormat.Dos;

                    message.WriteTo(options, filtered);
                    filtered.Flush();
                    return;
                }
            }
            catch
            {
                // An exception here probably means that the disk is full.
                //
                // Delete the file that was created above so that incomplete files are not
                // left behind for IIS to send accidentally.
                File.Delete(path);
                throw;
            }
        } while (true);
    }

    /// <summary>
    /// Disposes and clears any attachments and linked resources. These may hold open file or
    /// stream handles (e.g., from <see cref="AddLinkedResource(string, string?)"/>), so they
    /// must be released when they are no longer needed.
    /// </summary>
    private void DisposeAttachmentsAndResources()
    {
        if (LinkedResources != null)
        {
            foreach (var resource in LinkedResources)
            {
                resource?.Dispose();
            }
            LinkedResources.Clear();
        }

        if (Attachments != null)
        {
            foreach (var attachment in Attachments)
            {
                attachment?.Dispose();
            }
            Attachments.Clear();
        }
    }

    /// <summary>
    /// Releases the pending attachments, linked resources, and the reusable SMTP client held by this service.
    /// </summary>
    /// <param name="disposing">True when called from <see cref="Dispose()"/>; false when called from a finalizer.</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposed)
        {
            return;
        }

        if (disposing)
        {
            DisposeAttachmentsAndResources();
            DisposeClient();
        }

        _disposed = true;
    }

    /// <summary>
    /// Releases all resources used by the <see cref="MailmanService"/>, including the reusable SMTP
    /// connection and any open file or stream handles held by pending attachments and linked resources
    /// that were not sent. Prefer <see cref="DisposeAsync"/> so the SMTP connection can be closed gracefully.
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Asynchronously releases resources: disposes any pending attachments and linked resources and,
    /// when connected, gracefully disconnects (SMTP QUIT) and disposes the reusable SMTP client.
    /// Override to release additional resources asynchronously in derived types.
    /// </summary>
    protected virtual async ValueTask DisposeAsyncCore()
    {
        if (!_disposed)
        {
            DisposeAttachmentsAndResources();

            if (_smtpClient != null)
            {
                if (_smtpClient.IsConnected)
                {
                    try
                    {
                        await _smtpClient.DisconnectAsync(true).ConfigureAwait(false);
                    }
                    catch
                    {
                        // Best-effort graceful disconnect; ignore failures during teardown.
                    }
                }

                _smtpClient.Dispose();
                _smtpClient = null;
            }

            _disposed = true;
        }
    }

    /// <summary>
    /// Asynchronously releases all resources used by the <see cref="MailmanService"/>, gracefully closing
    /// the reusable SMTP connection and releasing any open file or stream handles held by pending
    /// attachments and linked resources that were not sent.
    /// </summary>
    public async ValueTask DisposeAsync()
    {
        await DisposeAsyncCore().ConfigureAwait(false);
        GC.SuppressFinalize(this);
    }
}