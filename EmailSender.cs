// Copyright © 2016-2023  ASM-SW
//asmeyers@outlook.com  https://github.com/asm-sw

using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.Identity.Client;
using MimeKit;
using System.IO;

namespace EmailWithAttachedFile
{
    /// <summary>
    /// This class sends email with an attached file.  The email is based on a text template file.  The attached file can be any file type.
    /// The token:  "<NAME>" in the template file is replaced with the a string passed into the sender.
    /// </summary>
    class EmailSender
    {
        bool m_init = false;
        string m_emailTemplate = string.Empty;
        string m_accessToken = string.Empty;
        readonly string NAMETOKEN = "<NAME>";

        /// <summary>
        /// Checks to see if the template file exists and reads it.
        /// Authenticates with Microsoft Graph/Office 365 using OAuth2.
        /// </summary>
        public async Task<(bool success, string errMsg)> InitAsync()
        {
            string errMsg = string.Empty;

            // 1. Authenticate and get Access Token
            try
            {
                string[] scopes = { "https://outlook.office.com/SMTP.Send" };
                IPublicClientApplication pca = PublicClientApplicationBuilder.Create(EmailSenderTest.EmailSettings.ClientId)
                    .WithAuthority(AzureCloudInstance.AzurePublic, EmailSenderTest.EmailSettings.TenantId)
                    .WithRedirectUri("http://localhost")
                    .Build();

                // This will trigger the browser popup for interactive login
                AuthenticationResult authResult = await pca.AcquireTokenInteractive(scopes).ExecuteAsync();
                m_accessToken = authResult.AccessToken;
            }
            catch (MsalException msalEx)
            {
                errMsg = $"Authentication Error: {msalEx.Message}\n";
                return (false, errMsg);
            }
            catch (Exception ex)
            {
                errMsg = $"General Error during Auth: {ex.Message}\n";
                return (false, errMsg);
            }

            // 2. Read Template File
            errMsg = string.Empty;

            if (!File.Exists(EmailSenderTest.EmailSettings.TemplateFileName))
            {
                errMsg = $"ERROR: Template file does not exist: {EmailSenderTest.EmailSettings.TemplateFileName}\n";
                return (false, errMsg);
            }

            try
            {
                using StreamReader streamReader = new StreamReader(EmailSenderTest.EmailSettings.TemplateFileName);
                m_emailTemplate = streamReader.ReadToEnd();
            }
            catch (Exception ex)
            {
                errMsg = $"ERROR: {ex.Message}\n";
                return (false, errMsg);
            }

            if (m_emailTemplate.IndexOf(NAMETOKEN) < 0)
                errMsg = "WARNING:  template file does not contain the token \"<NAME>\"\n";

            m_init = true;
            return (true, errMsg);
        }

        /// <summary>
        /// Sends an email
        /// </summary>
        /// <param name="name">This is used to replace the name token in the template file </param>
        /// <param name="email">email address list to send the email to.  Separate emails with a semicolon</param>
        /// <param name="fileName">filename to attach to the file</param>
        public async Task<(bool success, string errMsg)> SendMailAsync(string name, string email, string fileName)
        {
            string errMsg = string.Empty;
            if (!m_init)
            {
                errMsg = "ERROR:  Email Sender has not been initialized";
                return (false, errMsg);
            }
            string body = m_emailTemplate.Replace(NAMETOKEN, name);
            try
            {
                BodyBuilder builder = new BodyBuilder { TextBody = body };
                builder.Attachments.Add(fileName);

                foreach (string attachment in EmailSenderTest.EmailSettings.Attachments)
                {
                    if (System.IO.File.Exists(attachment))
                        builder.Attachments.Add(attachment);
                }

                MimeKit.MimeMessage mail = new MimeMessage();
                mail.From.Add(new MailboxAddress(EmailSenderTest.EmailSettings.EmailAuthor, EmailSenderTest.EmailSettings.EmailAddress));

                string[] emailList = email.Split(';');
                foreach (string emailItem in emailList)
                {
                    string item = emailItem.Trim();
                    if (!string.IsNullOrWhiteSpace(item))
                        mail.To.Add(MailboxAddress.Parse(item));
                }
                mail.Subject = EmailSenderTest.EmailSettings.MailSubject;
                mail.Body = builder.ToMessageBody();

                using (SmtpClient client = new SmtpClient())
                {
                    await client.ConnectAsync("smtp.office365.com", 587, SecureSocketOptions.StartTls);

                    SaslMechanismOAuth2 oauth2 = new SaslMechanismOAuth2(EmailSenderTest.EmailSettings.EmailAddress, m_accessToken);
                    await client.AuthenticateAsync(oauth2);

                    await client.SendAsync(mail);
                    await client.DisconnectAsync(true);
                }
            }
            catch (Exception ex)
            {
                errMsg = $"ERROR sending mail to: {name}, {email}\n\t{ex.Message}";
                return (false, errMsg);
            }

            return (true, string.Empty);
        }
    }
}
