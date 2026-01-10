using System.Net.Mail;

namespace Toolshed.Mailman
{
    /// <summary>
    /// Configuration settings for the Mailman email service
    /// </summary>
    public class MailmanSettings
    {
        /// <summary>
        /// The directory path where email messages are saved when using SpecifiedPickupDirectory delivery method
        /// </summary>
        public string PickupDirectoryLocation { get; set; }

        /// <summary>
        /// This uses .NET delivery method because MimeKit doesn't have a version of this, so we'll keep using instead of creating our own
        /// </summary>
        public SmtpDeliveryMethod DeliveryMethod { get; set; }

        /// <summary>
        /// The username authentication
        /// </summary>
        public string UserName { get; set; }

        /// <summary>
        /// The password authentication
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// The server hostname or IP address
        /// </summary>
        public string Host { get; set; }

        /// <summary>
        /// The server port. Defaults to 25.
        /// </summary>
        public int Port { get; set; } = 25;

        /// <summary>
        /// The timeout in milliseconds for operations. If null, the default timeout is used.
        /// </summary>
        public int? Timeout { get; set; }

        /// <summary>
        /// The default sender email address used when From is not explicitly set
        /// </summary>
        public string FromAddress { get; set; }

        /// <summary>
        /// The default sender display name used when From is not explicitly set
        /// </summary>
        public string FromDisplayName { get; set; }

        /// <summary>
        /// A comma delimited list of values to add to the X-SMTPAPI header when the email is sent
        /// </summary>
        public string Categories { get; set; }

        /// <summary>
        /// Gets or sets whether connecting via SSL/TLS should check certificate revocation.
        /// Defaults to true for security reasons. Set to false if Certificate Authority
        /// CRL/OCSP servers are unreachable, which can cause SslHandshakeException.
        /// </summary>
        public bool CheckCertificateRevocation { get; set; } = true;
    }
}
