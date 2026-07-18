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
/// Initializes a new instance of the MailmanService with the specified settings
/// </summary>
/// <param name="settings">The configuration settings for the mail service</param>
public class MailmanService(MailmanSettings settings) : IDisposable, IAsyncDisposable
{
    private readonly MailmanSettings _settings = settings;
    private bool _disposed;

    /// <summary>
    /// The subject line of the email message
    /// </summary>
    public string? Subject { get; set; }

    /// <summary>
    /// The name of the view to use for rendering the email body
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
    /// When true, the service will automatically reset after sending a message
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
    /// Sends a pre-configured MimeMessage using the SMTP client
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
    /// Applies the X-SMTPAPI category header (when categories are configured), saves to the pickup
    /// directory or connects and sends the message via SMTP, and optionally resets the service state.
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

            using var sm = new SmtpClient();
            if (_settings.Timeout.HasValue)
            {
                sm.Timeout = _settings.Timeout.Value;
            }

            sm.CheckCertificateRevocation = _settings.CheckCertificateRevocation;

            await sm.ConnectAsync(_settings.Host, _settings.Port, _settings.SecureSocketOptions, cancellationToken).ConfigureAwait(false);

            if (!string.IsNullOrWhiteSpace(_settings.UserName) && !string.IsNullOrWhiteSpace(_settings.Password))
            {
                await sm.AuthenticateAsync(new System.Net.NetworkCredential(_settings.UserName, _settings.Password), cancellationToken).ConfigureAwait(false);
            }

            await sm.SendAsync(mailMessage, cancellationToken).ConfigureAwait(false);
            await sm.DisconnectAsync(true, cancellationToken).ConfigureAwait(false);

            if (IsResetedAfterMessageSent)
            {
                Reset();
            }
        }
        finally
        {
            if (disposeMessage)
            {
                mailMessage.Dispose();
            }
        }
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
    /// Releases the file and stream handles held by any pending attachments and linked resources.
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
        }

        _disposed = true;
    }

    /// <summary>
    /// Releases all resources used by the <see cref="MailmanService"/>, including any open file
    /// or stream handles held by pending attachments and linked resources that were not sent.
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Asynchronously releases the file and stream handles held by any pending attachments and
    /// linked resources. Override to release additional resources asynchronously in derived types.
    /// </summary>
    protected virtual ValueTask DisposeAsyncCore()
    {
        if (!_disposed)
        {
            DisposeAttachmentsAndResources();
            _disposed = true;
        }

        return ValueTask.CompletedTask;
    }

    /// <summary>
    /// Asynchronously releases all resources used by the <see cref="MailmanService"/>, including any
    /// open file or stream handles held by pending attachments and linked resources that were not sent.
    /// </summary>
    public async ValueTask DisposeAsync()
    {
        await DisposeAsyncCore().ConfigureAwait(false);
        GC.SuppressFinalize(this);
    }
}