using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using MimeKit.IO;
using MimeKit.Text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Toolshed.Mailman;

public class MailmanService
{
    private MailmanSettings _settings;

    public string Subject { get; set; }
    public string ViewName { get; set; }
    public MessageImportance Importance { get; set; } = MessageImportance.Normal;
    public MessagePriority Priority { get; set; } = MessagePriority.Normal;
    public bool IsBodyHtml { get; set; } = true;
    public bool IsResetedAfterMessageSent { get; set; }

    /// <summary>
    /// A list of categories to add to X-SMTPAPI header. This will overwrite any categories provided by the settings. To keep the current categories, use the AddCategory() method
    /// </summary>
    public string Categories { get; set; }

    private string InternalCategories { get; set; }


    public void AddTo(string email)
    {
        To.Add(new MailboxAddress(email, email));
    }
    public void AddTo(string name, string email)
    {
        To.Add(new MailboxAddress(name, email));
    }
    public void AddCc(string email)
    {
        CC.Add(new MailboxAddress(email, email));
    }
    public void AddCc(string name, string email)
    {
        CC.Add(new MailboxAddress(name, email));
    }
    public void AddBcc(string email)
    {
        Bcc.Add(new MailboxAddress(email, email));
    }
    public void AddBcc(string name, string email)
    {
        Bcc.Add(new MailboxAddress(name, email));
    }

    public void AddTo(IEnumerable<string> emails)
    {
        To.AddRange(emails.Select(email => new MailboxAddress(email, email)).ToList());
    }
    public void AddCc(IEnumerable<string> emails)
    {
        CC.AddRange(emails.Select(email => new MailboxAddress(email, email)).ToList());
    }
    public void AddBcc(IEnumerable<string> emails)
    {
        Bcc.AddRange(emails.Select(email => new MailboxAddress(email, email)).ToList());
    }

    public void AddFrom(string email)
    {
        Froms.Add(new MailboxAddress(email, email));
    }



    MailboxAddress _From;
    public MailboxAddress From
    {
        get
        {
            if (_From == null)
            {
                if (!string.IsNullOrWhiteSpace(_settings.FromDisplayName))
                {
                    _From = new MailboxAddress(_settings.FromDisplayName, _settings.FromAddress);
                }
                else
                {
                    _From = new MailboxAddress(_settings.FromAddress, _settings.FromAddress);
                }
            }

            return _From;
        }
        set { _From = value; }
    }


    List<MailboxAddress> _Froms;
    /// <summary>
    ///
    /// </summary>
    public List<MailboxAddress> Froms
    {
        get
        {
            if (_Froms == null)
            {
                _Froms = new List<MailboxAddress>();
            }

            return _Froms;
        }
        set
        {
            _Froms = value;
        }
    }

    List<MailboxAddress> _To;
    public List<MailboxAddress> To
    {
        get
        {
            if (_To == null)
            {
                _To = new List<MailboxAddress>();
            }

            return _To;
        }
        set
        {
            _To = value;
        }
    }

    List<MailboxAddress> _CC;
    public List<MailboxAddress> CC
    {
        get
        {
            if (_CC == null)
            {
                _CC = new List<MailboxAddress>();
            }

            return _CC;
        }
        set
        {
            _CC = value;
        }
    }

    List<MailboxAddress> _Bcc;
    public List<MailboxAddress> Bcc
    {
        get
        {
            if (_Bcc == null)
            {
                _Bcc = new List<MailboxAddress>();
            }

            return _Bcc;
        }
        set
        {
            _Bcc = value;
        }
    }


    public MailmanService(MailmanSettings settings)
    {
        _settings = settings;

        InternalCategories = settings.Categories;
    }

    /// <summary>
    /// Clears the SUBJECT, TO, CC, and BCC properties. Can optionally reset the FROM. Settings are not affected.
    /// </summary>
    /// <param name="resetFrom">Specifies whether the FROM property should also be cleared</param>
    public void Reset(bool resetFrom = false)
    {
        Subject = null;
        _To = null;
        _CC = null;
        _Bcc = null;

        if (resetFrom)
        {
            From = null;
        }
    }

    /// <summary>
    /// There are no checks or exception catching, .NET will that for you. However, if you want a quick sanity check to make sure
    /// typical issues are looked at, we can do that. We check recipients, from  and if Delivery is network, we check for a host
    /// </summary>
    public bool IsValidToSend()
    {
        if (From is null && string.IsNullOrWhiteSpace(_settings.FromAddress))
        {
            return false;
        }
        if (_To is null && _CC is null && _Bcc is null)
        {
            return false;
        }
        var recipientValidatedAndItisGood = false;
        if (_To != null && _To.Count > 0)
        {
            recipientValidatedAndItisGood = true;
        }
        if (_CC != null && _CC.Count > 0)
        {
            recipientValidatedAndItisGood = true;
        }
        if (_Bcc != null && _Bcc.Count > 0)
        {
            recipientValidatedAndItisGood = true;
        }
        if (!recipientValidatedAndItisGood)
        {
            return false;
        }
        if (_settings.DeliveryMethod == System.Net.Mail.SmtpDeliveryMethod.Network && string.IsNullOrWhiteSpace(_settings.Host))
        {
            return false;
        }

        return true;
    }

    public AttachmentCollection Attachments { get; set; } = new AttachmentCollection(false);

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


    
    
    public MimeMessage GetMessage(string body, bool isHtml = true)
    {
        IsBodyHtml = isHtml;
        var message = new MimeMessage
        {
            Subject = Subject,
            //SubjectEncoding = Encoding,
            //BodyEncoding = Encoding,
            Priority = Priority,
            Importance = Importance
        };

        if (Attachments.Count > 0)
        {
            var builder = new BodyBuilder();
            if (IsBodyHtml)
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

            message.Body = builder.ToMessageBody();
        }
        else
        {
            if (IsBodyHtml)
            {
                message.Body = new TextPart(TextFormat.Html) { Text = body };
            }
            else
            {
                message.Body = new TextPart(TextFormat.Text) { Text = body };
            }
        }

        if (_Froms != null && _Froms.Count > 0)
        {
            message.From.AddRange(Froms);

            if (message.From.Count > 1)
            {
                message.Sender = Froms[0];
            }
        }
        else if (From != null)
        {
            message.From.Add(From);
        }

        if (message.From.Count == 0 && !string.IsNullOrWhiteSpace(_settings.FromAddress))
        {
            message.From.Add(new MailboxAddress(_settings.FromDisplayName ?? _settings.FromAddress, _settings.FromAddress));
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
    /// Send a message using the SMTP client. The message can either be a string or HTML. 
    /// There is no longer a method to use a MVC view to create a message. The RazorSlice extensions my be added, but fo rnow they are gone.
    /// </summary>
    /// <param name="body"></param>
    /// <param name="isHtml"></param>
    /// <returns></returns>
    public async Task SendMessageAsync(string body, bool isHtml = true)
    {
        var mailMessage = GetMessage(body, isHtml);

        if (!string.IsNullOrWhiteSpace(Categories))
        {
            mailMessage.Headers.Add("X-SMTPAPI", "{\"category\":[\"" + Categories + "\"]}");
        }
        else if (!string.IsNullOrWhiteSpace(InternalCategories))
        {
            mailMessage.Headers.Add("X-SMTPAPI", "{\"category\":[\"" + InternalCategories + "\"]}");
        }

        if (_settings.DeliveryMethod == System.Net.Mail.SmtpDeliveryMethod.SpecifiedPickupDirectory)
        {
            SaveToPickupDirectory(mailMessage, _settings.PickupDirectoryLocation);
            return;
        }

        using (var sm = new SmtpClient { })
        {
            if (_settings.Timeout.HasValue)
            {
                sm.Timeout = _settings.Timeout.Value;
            }

            sm.CheckCertificateRevocation = _settings.CheckCertificateRevocation;

            await sm.ConnectAsync(_settings.Host, _settings.Port, SecureSocketOptions.Auto);

            if (!string.IsNullOrWhiteSpace(_settings.UserName) || !string.IsNullOrWhiteSpace(_settings.Password))
            {
                sm.Authenticate(new System.Net.NetworkCredential(_settings.UserName, _settings.Password));
            }

            await sm.SendAsync(mailMessage);
            await sm.DisconnectAsync(true);
        }

        if (IsResetedAfterMessageSent)
        {
            Reset();
        }
    }
    
    
    //this is the final send using these shortcuts
    public async Task SendMessageAsync(MimeMessage mailMessage)
    {
        if (!string.IsNullOrWhiteSpace(Categories))
        {
            mailMessage.Headers.Add("X-SMTPAPI", "{\"category\":[\"" + Categories + "\"]}");
        }
        else if (!string.IsNullOrWhiteSpace(InternalCategories))
        {
            mailMessage.Headers.Add("X-SMTPAPI", "{\"category\":[\"" + InternalCategories + "\"]}");
        }

        if (Attachments.Count > 0)
        {

        }

        if (_settings.DeliveryMethod == System.Net.Mail.SmtpDeliveryMethod.SpecifiedPickupDirectory)
        {
            SaveToPickupDirectory(mailMessage, _settings.PickupDirectoryLocation);
            return;
        }

        using (var sm = new SmtpClient { })
        {
            if (_settings.Timeout.HasValue)
            {
                sm.Timeout = _settings.Timeout.Value;
            }

            sm.CheckCertificateRevocation = _settings.CheckCertificateRevocation;

            await sm.ConnectAsync(_settings.Host, _settings.Port, SecureSocketOptions.Auto);

            if (!string.IsNullOrWhiteSpace(_settings.UserName) || !string.IsNullOrWhiteSpace(_settings.Password))
            {
                sm.Authenticate(new System.Net.NetworkCredential(_settings.UserName, _settings.Password));
            }

            await sm.SendAsync(mailMessage);
            await sm.DisconnectAsync(true);
        }

        if (IsResetedAfterMessageSent)
        {
            Reset();
        }
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
                    using (var filtered = new FilteredStream(stream))
                    {
                        filtered.Add(new SmtpDataFilter());

                        // Make sure to write the message in DOS (<CR><LF>) format.
                        var options = FormatOptions.Default.Clone();
                        options.NewLineFormat = NewLineFormat.Dos;

                        message.WriteTo(options, filtered);
                        filtered.Flush();
                        return;
                    }
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
}