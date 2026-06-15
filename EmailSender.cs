﻿// Copyright © 2016-2023  ASM-SW
//asmeyers@outlook.com  https://github.com/asm-sw

using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.Identity.Client;
using MimeKit;
using System.IO;

namespace EmailWithAttachedFile
{
    public class EmailJob
    {
        public string Name { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
    }


    /// <summary>
    /// This class sends email with an attached file.  The email is based on a text template file.  The attached file can be any file type.
    /// The token:  "<NAME>" in the template file is replaced with the a string passed into the sender.
    /// </summary>
    class EmailSender
    {
        bool m_init = false;
        string m_emailTemplate = string.Empty;
        string m_accessToken = string.Empty;
        SmtpClient? m_smtpClient;
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
        /// Establishes a connection to the SMTP server and authenticates.
        /// </summary>
        public async Task<(bool success, string errMsg)> ConnectAsync(CancellationToken token = default)
        {
            try
            {
                m_smtpClient = new SmtpClient();
                await m_smtpClient.ConnectAsync("smtp.office365.com", 587, SecureSocketOptions.StartTls, token);

                SaslMechanismOAuth2 oauth2 = new(EmailSettings.EmailAddress, m_accessToken);
                await m_smtpClient.AuthenticateAsync(oauth2, token);
                return (true, string.Empty);
            }
            catch (Exception ex)
            {
                return (false, $"Failed to connect to SMTP server: {ex.Message}");
            }
        }

        /// <summary>
        /// Disconnects and cleans up the SMTP client.
        /// </summary>
        public async Task DisconnectAsync()
        {
            if (m_smtpClient != null)
            {
                if (m_smtpClient.IsConnected)
                    await m_smtpClient.DisconnectAsync(true);
                m_smtpClient.Dispose();
                m_smtpClient = null;
            }
        }

        /// <summary>
        /// Sends an email
        /// </summary>
        /// <param name="job">The email job containing recipient details and attachment</param>
        public async Task<(bool success, string errMsg)> SendMailAsync(EmailJob job, CancellationToken token = default)
        {
            if (!m_init || m_smtpClient == null || !m_smtpClient.IsConnected)
                return (false, "ERROR: Email Sender is not connected or initialized");

            try
            {
                string body = m_emailTemplate.Replace(NAMETOKEN, job.Name);
                BodyBuilder builder = new() { TextBody = body };
                builder.Attachments.Add(job.FileName);

                foreach (string attachment in EmailSettings.Attachments)
                {
                    if (System.IO.File.Exists(attachment))
                        builder.Attachments.Add(attachment);
                }

                MimeKit.MimeMessage mail = new();
                mail.From.Add(new MailboxAddress(EmailSettings.EmailAuthor, EmailSettings.EmailAddress));

                char[] delimiters = [',', ';', ' '];
                string[] emailList = (job.Email ?? string.Empty).Split(delimiters, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                foreach (string emailAddress in emailList)
                    mail.To.Add(MailboxAddress.Parse(emailAddress));

                mail.Subject = EmailSettings.MailSubject;
                mail.Body = builder.ToMessageBody();

                await m_smtpClient.SendAsync(mail, token);
            }
            catch (Exception ex)
            {
                return (false, $"ERROR sending mail to: {job.Name}, {job.Email}\n\t{ex.Message}");
            }

            return (true, string.Empty);
        }
    }
}
