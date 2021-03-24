using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using MimeKit.IO;
using MimeKit.Text;

namespace Toolshed.Mailman
{
    public class MailmanService
    {
        private ViewRenderService _viewRenderService;
        private MailmanSettings _settings;

        public string Subject { get; set; }
        public string ViewName { get; set; }
        public bool IsAlternateViewsUsed { get; set; }
        public System.Text.Encoding Encoding { get; set; } = System.Text.Encoding.UTF8;
        public MessageImportance Importance { get; set; } = MessageImportance.Normal;
        public MessagePriority Priority { get; set; } = MessagePriority.Normal;
        public bool IsBodyHtml { get; set; } = true;
        public bool IsResetedAfterMessageSent { get; set; }

        /// <summary>
        /// A list of categories to add to X-SMTPAPI header. This will overwrite any categories provided by the settings. To keep the current categories, use the AddCategory() method
        /// </summary>
        public string Categories { get; set; }

        private string InternalCategories { get; set; }

        /// <summary>
        /// Indicates what password to use from the settings and overrides the value in the settings if provided
        /// </summary>
        public int? UsePassword { get; set; }


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


        public MailmanService(MailmanSettings settings, ViewRenderService viewRenderService)
        {
            _viewRenderService = viewRenderService;
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

        public Task<MimeMessage> GetMessageAsync<T>(T model)
        {
            return GetMessageAsync(ViewName, model);
        }
        public async Task<MimeMessage> GetMessageAsync<T>(string viewName, T model)
        {
            var message = new MimeMessage
            {
                Subject = Subject,
                //SubjectEncoding = Encoding,
                //BodyEncoding = Encoding,
                Priority = Priority,
                Importance = Importance
            };

            if (!string.IsNullOrWhiteSpace(viewName))
            {
                var html = await _viewRenderService.RenderAsString(viewName, model);
                if (IsBodyHtml)
                {
                    message.Body = new TextPart(TextFormat.Html) { Text = html };
                }
                else
                {
                    message.Body = new TextPart(TextFormat.Text) { Text = html };
                }
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

        public async Task SendMessageAsync<T>(string viewName, T model)
        {
            var message = await GetMessageAsync(viewName, model);
            await SendMessageAsync(message);
        }
        public async Task SendMessageAsync<T>(T model)
        {
            if (string.IsNullOrWhiteSpace(ViewName))
            {
                throw new ArgumentNullException("ViewName", "The view name must be provided either by the property ViewName or using the method that has it as a parameter");
            }

            var message = await GetMessageAsync(ViewName, model);
            await SendMessageAsync(message);
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

                if (!string.IsNullOrWhiteSpace(_settings.UserName) || !string.IsNullOrWhiteSpace(_settings.Password))
                {
                    sm.Authenticate(new System.Net.NetworkCredential(_settings.UserName, _settings.Password));
                }

                await sm.ConnectAsync(_settings.Host, _settings.Port, SecureSocketOptions.Auto);
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
}

