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

        //FROM MAILKIT
        //
        // Summary:
        //     Get or set whether connecting via SSL/TLS should check certificate revocation.
        //
        // Value:
        //     true if certificate revocation should be checked; otherwise, false.
        //
        // Remarks:
        //     Gets or sets whether connecting via SSL/TLS should check certificate revocation.
        //     Normally, the value of this property should be set to true (the default) for
        //     security reasons, but there are times when it may be necessary to set it to false.
        //     For example, most Certificate Authorities are probably pretty good at keeping
        //     their CRL and/or OCSP servers up 24/7, but occasionally they do go down or are
        //     otherwise unreachable due to other network problems between the client and the
        //     Certificate Authority. When this happens, it becomes impossible to check the
        //     revocation status of one or more of the certificates in the chain resulting in
        //     an MailKit.Security.SslHandshakeException being thrown in the Connect method.
        //     If this becomes a problem, it may become desirable to set MailKit.MailService.CheckCertificateRevocation
        //     to false.
        public bool CheckCertificateRevocation { get; set; } = true;
    }
}
