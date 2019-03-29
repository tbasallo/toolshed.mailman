using System.Net.Mail;

namespace Toolshed.Mailman
{
    public class MailmanSettings
    {
        public string PickupDirectoryLocation { get; set; }
        public SmtpDeliveryMethod DeliveryMethod { get; set; }
        public string UserName { get; set; }

        public string Password { get; set; }
        public string Password1 { get; set; }
        public string Password2 { get; set; }
        /// <summary>
        /// Indicates which password to use. 0 indicates "Password" and is the default. 1 and 2 indicate Passwordx where x is the number after the word "Password"
        /// </summary>
        public int UsePassword { get; set; }

        public string Host { get; set; }
        public int? Port { get; set; }
        public bool? EnableSsl { get; set; }
        public int? Timeout { get; set; }
        public string FromAddress { get; set; }
        public string FromDisplayName { get; set; }
    }
}
