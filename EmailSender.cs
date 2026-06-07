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
            //Authenticate and get Access Token
            try
            {
                string[] scopes = ["https://outlook.office.com/SMTP.Send"];
                IPublicClientApplication pca = PublicClientApplicationBuilder.Create(EmailSettings.ClientId)
                    .WithAuthority(AzureCloudInstance.AzurePublic, EmailSettings.TenantId)
                    .WithRedirectUri("http://localhost")
                    .Build();

                AuthenticationResult authResult;
                var accounts = await pca.GetAccountsAsync();

                // If silent fails, then trigger the browser popup
                authResult = await pca.AcquireTokenInteractive(scopes).ExecuteAsync();
                m_accessToken = authResult.AccessToken;
            }
            catch (MsalException msalEx)
            {
                return (false, $"Authentication Error: {msalEx.Message}\n");
            }
            catch (Exception ex)
            {
                return (false, $"General Error during Auth: {ex.Message}\n");
            }

            if (!File.Exists(EmailSettings.TemplateFileName))
            {
                return (false, $"ERROR: Template file does not exist: {EmailSettings.TemplateFileName}\n");
            }
            try
            {
                using StreamReader streamReader = new(EmailSettings.TemplateFileName);
                m_emailTemplate = streamReader.ReadToEnd();
            }
            catch (Exception ex)
            {
                return (false, $"ERROR: {ex.Message}\n");
            }

            m_init = true;
            if (m_emailTemplate.IndexOf(NAMETOKEN) < 0)
                return (true, "WARNING:  template file does not contain the token \"<NAME>\"\n");
            return (true, string.Empty);
        }

        /// <summary>
        /// Sends an email
        /// </summary>
        /// <param name="name">This is used to replace the name token in the template file </param>
        /// <param name="email">email address list to send the email to.  Separate emails with a semicolon</param>
        /// <param name="fileName">filename to attach to the file</param>
        public async Task<(bool success, string errMsg)> SendMailAsync(string name, string email, string fileName)
        {
            if (!m_init)
                return (false, "ERROR:  Email Sender has not been initialized");

            try
            {
                string body = m_emailTemplate.Replace(NAMETOKEN, name);
                BodyBuilder builder = new() { TextBody = body };
                builder.Attachments.Add(fileName);

                foreach (string attachment in EmailSettings.Attachments)
                {
                    if (System.IO.File.Exists(attachment))
                        builder.Attachments.Add(attachment);
                }

                MimeKit.MimeMessage mail = new();
                mail.From.Add(new MailboxAddress(EmailSettings.EmailAuthor, EmailSettings.EmailAddress));

                string[] emailList = (email ?? string.Empty).Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                foreach (string emailItem in emailList)
                    mail.To.Add(MailboxAddress.Parse(emailItem));

                mail.Subject = EmailSettings.MailSubject;
                mail.Body = builder.ToMessageBody();

                using SmtpClient client = new();
                await client.ConnectAsync("smtp.office365.com", 587, SecureSocketOptions.StartTls);

                SaslMechanismOAuth2 oauth2 = new(EmailSettings.EmailAddress, m_accessToken);
                await client.AuthenticateAsync(oauth2);

                await client.SendAsync(mail);
                await client.DisconnectAsync(true);
            }
            catch (Exception ex)
            {
                return (false, $"ERROR sending mail to: {name}, {email}\n\t{ex.Message}");
            }

            return (true, string.Empty);
        }
    }
}
