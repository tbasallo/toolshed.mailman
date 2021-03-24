using System.Net.Mail;

namespace Toolshed.Mailman
{
    public class MailmanSettings
    {
        public string PickupDirectoryLocation { get; set; }
        /// <summary>
        /// This uses .NET delivery method because MimeKit doesn't have a version of this, so we'll keep using instead of creating or own
        /// </summary>
        public SmtpDeliveryMethod DeliveryMethod { get; set; }
        public string UserName { get; set; }

        public string Password { get; set; }        
        

        public string Host { get; set; }
        public int Port { get; set; } = 25;        
        public int? Timeout { get; set; }
        public string FromAddress { get; set; }
        public string FromDisplayName { get; set; }

        /// <summary>
        /// A comma delimited list of values to add to the X-SMTPAPI header when the email is sent
        /// </summary>
        public string Categories { get; set; }
    }
}
