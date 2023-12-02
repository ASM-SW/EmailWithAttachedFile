// Copyright © 2016-2023  ASM-SW
//asmeyers@outlook.com  https://github.com/asm-sw

using System;
using System.IO;
using MimeKit;
using System.Text;
using System.Net;

namespace EmailWithAttachedFile
{
    /// <summary>
    /// This class sends email with an attached file.  The email is based on a text template file.  The attached file can be any file type.
    /// The token:  "<NAME>" in the template file is replaced with the a string passed into the sender.
    /// </summary>
    class EmailSender
    {
        public EmailSender() { }

        ConfigurationEmailWAF m_config;
        bool m_init = false;
        string m_emailTemplate = string.Empty;
        readonly string NAMETOKEN = "<NAME>";
        NetworkCredential m_credential;

        /// <summary>
        /// Checks to see if the template file exists and reads it.
        /// Checks the the token is in the template file, and displays a warning if it is missing.
        /// </summary>
        /// <param name="config"></param>
        /// <param name="errMsg"></param>
        /// <returns></returns>
        public bool Init(ref ConfigurationEmailWAF config, out StringBuilder errMsg)
        {
            errMsg = new StringBuilder();
            m_config = config;
            m_init = true;
            m_credential = new NetworkCredential(m_config.FromEmail, m_config.Password);

            if (!File.Exists(m_config.TemplateFileName))
            {
                errMsg.AppendFormat("ERROR: Template file does not exist: {0}\n", m_config.TemplateFileName);
                m_init = false;
            }

            try
            {
                using (StreamReader streamReader = new StreamReader(m_config.TemplateFileName))
                {
                    m_emailTemplate = streamReader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                errMsg.AppendFormat("ERROR: {0}\n", ex.Message);
                m_init = false;
                return m_init;
            }

            if (m_emailTemplate.IndexOf(NAMETOKEN) < 0)
                errMsg.AppendLine("WARNING:  template file does not contain the token \"<NAME>\"");

            return m_init;
        }

        /// <summary>
        /// Sends an email
        /// </summary>
        /// <param name="name">This is used to replace the name token in the template file </param>
        /// <param name="email">email address list to send the email to.  Separate emails with a semicolon</param>
        /// <param name="fileName">filename to attach to the file</param>
        /// <param name="errMsg">If the method returns false, contains a error message</param>
        /// <returns>true if no errors occurred</returns>
        public bool SendMail(string name, string email, string fileName, out string errMsg)
        {
            errMsg = string.Empty;
            if (!m_init)
            {
                errMsg = "ERROR:  Email Sender has not been initialized";
                return false;
            }
            string body = m_emailTemplate.Replace(NAMETOKEN, name);
            try
            {
                BodyBuilder builder = new BodyBuilder { TextBody = body };
                builder.Attachments.Add(fileName);

                MimeKit.MimeMessage mail = new MimeMessage();
                mail.From.Add(MailboxAddress.Parse(m_config.FromEmail));

                string[] emailList = email.Split(';');
                foreach (string emailItem in emailList)
                {
                    string item = emailItem.Trim();
                    if (!string.IsNullOrWhiteSpace(item))
                        mail.To.Add(MailboxAddress.Parse(item));
                }
                mail.Subject = m_config.MailSubject;
                mail.Body = builder.ToMessageBody();

                //todo reuse client, would have to process list of emails here, rather than one at a time.
                // this might allow you to use one connection for multliple emails
                using (MailKit.Net.Smtp.SmtpClient client = new MailKit.Net.Smtp.SmtpClient())
                {
                    MailKit.Security.SecureSocketOptions secureOption = MailKit.Security.SecureSocketOptions.None;
                    if (m_config.SmtpEnabledSSL)
                        secureOption = MailKit.Security.SecureSocketOptions.Auto;

                    client.Connect(m_config.SmtpServer, m_config.SmtpPort, secureOption);
                    client.Authenticate(credentials: m_credential);
                    client.Send(mail);
                    client.Disconnect(true);
                }
            }
            catch (Exception ex)
            {
                errMsg = $"ERROR sending mail to: {name}, {email}\n\t{ex.Message}";
                return false;
            }

            return true;
        }
    }
}
