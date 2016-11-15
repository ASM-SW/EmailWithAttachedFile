using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;

namespace EmailWithAttachedFile
{
    class EmailSender
    {
        public EmailSender() { }

        ConfigurationEmailWAF m_config;
        bool m_init = false;
        string m_emailTemplate = string.Empty;
        readonly string NAMETOKEN = "<NAME>";


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
