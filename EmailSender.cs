using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;

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
        /// <param name="email">email address to send the email to</param>
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

            using (MailMessage mail = new MailMessage())
            {

                using (SmtpClient smtpClient = new SmtpClient(m_config.SmtpServer))
                {
                    //todo reuse smtp client  https://msdn.microsoft.com/en-us/library/system.net.mail.smtpclient(v=vs.110).aspx

                    try
                    {

                        mail.From = new MailAddress(m_config.FromEmail);
                        mail.To.Add(email);
                        mail.Subject = m_config.MailSubject;
                        mail.Body = body;

                        Attachment attachment;
                        attachment = new Attachment(fileName);
                        FileInfo fInfo = new System.IO.FileInfo(fileName);
                        mail.Attachments.Add(attachment);

                        ContentDisposition disposition = attachment.ContentDisposition;
                        disposition.CreationDate = fInfo.CreationTime;
                        disposition.ModificationDate = fInfo.CreationTime;
                        disposition.ReadDate = fInfo.CreationTime;
                        disposition.FileName = Path.GetFileName(fileName);
                        disposition.Size = fInfo.Length;
                        disposition.DispositionType = DispositionTypeNames.Attachment;

                        smtpClient.Port = m_config.SmtpPort;
                        smtpClient.Credentials = new NetworkCredential(m_config.FromEmail, m_config.Password);
                        smtpClient.EnableSsl = m_config.SmtpEnabledSSL;

                        smtpClient.Send(mail);
                    }
                    catch (Exception ex)
                    {
                        errMsg = "ERROR sending mail to: " + name + "\n\t" + ex.Message;
                        return false;
                    }

                    return true;
                }
            }
        }
    }
}
