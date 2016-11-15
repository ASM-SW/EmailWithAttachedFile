using System;
using System.Runtime.Serialization;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml.Serialization;
using System.Windows;

namespace EmailWithAttachedFile
{
    [Serializable]
    public class ConfigurationEmailWAF
    {
        public string FromEmail { get; set; }
        public string SmtpServer { get; set; }
        public int SmtpPort { get; set; }
        public bool SmtpEnabledSSL { get; set; }
        public string TemplateFileName { get; set; }
        public string InputFileName { get; set; }
        public string MailSubject { get; set; }

        [XmlIgnore]
        public SecureString Password {get; set;}

        [XmlIgnore]
        public string ConfigFileName { get; set; }              // name of file containing configuration information

        public ConfigurationEmailWAF()
        {
            string appData = Environment.GetEnvironmentVariable("APPDATA");
            string dataDir = Path.Combine(appData, "DonorStatement");
            if (!Directory.Exists(dataDir))
                Directory.CreateDirectory(dataDir);
            ConfigFileName = Path.Combine(dataDir, "ConfigurationEamilWithAttachment.xml");
        }




        public bool Serialize(string fileName)
        {
            try
            {
                using (TextWriter writer = new StreamWriter(fileName))
                {
                    XmlSerializer ser = new XmlSerializer(typeof(ConfigurationEmailWAF));
                    ser.Serialize(writer, this);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.ToString());
            }
            return true;
        }

        public static bool DeSerialize(string fileName, ref ConfigurationEmailWAF cfg)
        {
            if (!File.Exists(fileName))
                return true;
            try
            {
                using (FileStream fileStream = new FileStream(fileName, FileMode.Open))
                {
                    XmlSerializer ser = new XmlSerializer(typeof(ConfigurationEmailWAF));
                    cfg = (ConfigurationEmailWAF)ser.Deserialize(fileStream);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return true;
        }

    }
}
