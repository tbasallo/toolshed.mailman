using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Threading.Tasks;

namespace Toolshed.Mailman
{
    public class MailmanService
    {
        private ViewRenderService _viewRenderService;
        private MailmanSettings _settings;

        public string Subject { get; set; }
        public MailAddress From { get; set; }
        public string ViewName { get; set; }

        List<string> _To;
        public List<string> To
        {
            get
            {
                if (_To == null)
                {
                    _To = new List<string>();
                }

                return _To;
            }
            set
            {
                _To = value;
            }
        }

        List<string> _CC;
        public List<string> CC
        {
            get
            {
                if (_CC == null)
                {
                    _CC = new List<string>();
                }

                return _CC;
            }
            set
            {
                _CC = value;
            }
        }

        List<string> _Bcc;
        public List<string> Bcc
        {
            get
            {
                if (_Bcc == null)
                {
                    _Bcc = new List<string>();
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
            if (_settings.DeliveryMethod == SmtpDeliveryMethod.Network && string.IsNullOrWhiteSpace(_settings.Host))
            {
                return false;
            }

            return true;
        }

        public Task<MailMessage> GetMessage<T>(T model)
        {
            return GetMessage(ViewName, model);
        }
        public async Task<MailMessage> GetMessage<T>(string viewName, T model)
        {
            var message = new MailMessage { IsBodyHtml = true, Subject = Subject };

            if (From != null)
            {
                message.From = From;
            }
            else if (!string.IsNullOrWhiteSpace(_settings.FromAddress))
            {
                if (string.IsNullOrWhiteSpace(_settings.FromDisplayName))
                {
                    message.From = new MailAddress(_settings.FromAddress, _settings.FromDisplayName);
                }
                else
                {
                    message.From = new MailAddress(_settings.FromAddress);
                }
            }

            var html = await _viewRenderService.RenderAsString(viewName, model);
            message.Body = html;
            message.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(html, System.Text.Encoding.UTF8, "text/html"));
            message.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(message.Subject, System.Text.Encoding.UTF8, "plain/text"));

            if (_To != null && _To.Count > 0)
            {
                if (_To.Count == 1)
                {
                    message.To.Add(_To[0]);
                }
                else
                {
                    message.To.Add(string.Join(',', _To));
                }
            }
            if (_CC != null && _CC.Count > 0)
            {
                if (_CC.Count == 1)
                {
                    message.To.Add(_CC[0]);
                }
                else
                {
                    message.To.Add(string.Join(',', _CC));
                }
            }
            if (_Bcc != null && _Bcc.Count > 0)
            {
                if (_Bcc.Count == 1)
                {
                    message.To.Add(_Bcc[0]);
                }
                else
                {
                    message.To.Add(string.Join(',', _Bcc));
                }
            }

            return message;
        }

        public async Task SendMessage<T>(string viewName, T model)
        {
            var message = await GetMessage(viewName, model);
            await SendMessage(message);
        }
        public async Task SendMessage<T>(T model)
        {
            if (string.IsNullOrWhiteSpace(ViewName))
            {
                throw new ArgumentNullException("ViewName", "The view name must be provided either by the property ViewName or using the method that has it as a parameter");
            }

            var message = await GetMessage(ViewName, model);
            await SendMessage(message);
        }

        public async Task SendMessage(MailMessage mailMessage)
        {
            using (var smtp = new SmtpClient())
            {
                smtp.DeliveryMethod = _settings.DeliveryMethod;

                if (smtp.DeliveryMethod == SmtpDeliveryMethod.SpecifiedPickupDirectory)
                {
                    smtp.PickupDirectoryLocation = _settings.PickupDirectoryLocation;
                }
                else
                {
                    if (!string.IsNullOrWhiteSpace(_settings.UserName) || !string.IsNullOrWhiteSpace(_settings.Password))
                    {
                        smtp.Credentials = new System.Net.NetworkCredential(_settings.UserName, _settings.Password);
                    }

                    if (!string.IsNullOrWhiteSpace(_settings.Host))
                    {
                        smtp.Host = _settings.Host;
                    }

                    if (_settings.Port.HasValue)
                    {
                        smtp.Port = _settings.Port.Value;
                    }

                    if (_settings.EnableSsl.HasValue)
                    {
                        smtp.EnableSsl = _settings.EnableSsl.Value;
                    }

                    if (_settings.Timeout.HasValue)
                    {
                        smtp.Timeout = _settings.Timeout.Value;
                    }
                }

                await smtp.SendMailAsync(mailMessage);
            }
        }
    }
}
