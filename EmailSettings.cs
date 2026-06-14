using System.IO;
using System.Text.Json;

namespace EmailWithAttachedFile
{
    /// <summary>
    /// This class is a singelton.  All access is through Static methods.
    /// Call EmailSettings.Init() to initialize it.
    /// Check the return value of Init to see if it initialized.
    /// EmailSettings.Message will contain an error message if the init failed.
    /// </summary>
    public class EmailSettings
    {
        private record Values(
            string TenantId, 
            string ClientId, 
            string EmailAddress, 
            string EmailAuthor,
            string TemplateFileName,
            string InputFileName,
            string MailSubject,
            List<string> Attachments);
        private Values _values = new(string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, []);

        private static readonly EmailSettings _instance = new();
        private static string _filePath = string.Empty;
        public static string SettingsFileName { get => _filePath;}

        /// <summary>
        /// Read in the settings file
        /// </summary>
        /// <returns>If false check EmailSettings.Message for error message</returns>
        public static bool Init()
        {
            Message = string.Empty;

            try
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string exeName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name ?? "EmailWithAttachedFile";
                string dataDir = Path.Combine(appData, exeName);
                if (!Directory.Exists(dataDir))
                {
                    Directory.CreateDirectory(dataDir);
                }
                _filePath = Path.Combine(dataDir, exeName);
                _filePath = Path.ChangeExtension(_filePath, "json");
            }
            catch (Exception ex)
            {
                Message = $"Error accessing AppData folder: {ex.Message}";
                return false;
            }

            return _instance.InitSettings(_filePath);
        }

        public static string TenantId => _instance._values.TenantId;
        public static string ClientId => _instance._values.ClientId;
        public static string EmailAddress => _instance._values.EmailAddress;
        public static string EmailAuthor => _instance._values.EmailAuthor;
        public static string Message { get; private set; } = string.Empty;

        public static string TemplateFileName => _instance._values.TemplateFileName;
        public static string InputFileName => _instance._values.InputFileName;
        public static string MailSubject => _instance._values.MailSubject;
        public static List<string> Attachments => _instance._values.Attachments ?? [];

        public static void Update(string tenantId, string clientId, string emailAddress, string emailAuthor, string templateFile, string inputFile, string subject, List<string> attachments)
        {
            _instance._values = new Values(tenantId, clientId, emailAddress, emailAuthor, templateFile, inputFile, subject, attachments ?? []);
        }

        /// <summary>
        /// Save the current settings to the file used during Init.
        /// </summary>
        /// <returns>True if successful</returns>
        public static bool Save() => Save(_filePath);

        /// <summary>
        /// Save the current settings to a json file.
        /// </summary>
        /// <param name="settingFileName"></param>
        /// <returns>True if successful</returns>
        public static bool Save(string settingFileName)
        {
            try
            {
                string jsonString = JsonSerializer.Serialize(_instance._values, _instance.options);
                File.WriteAllText(settingFileName, jsonString);
                return true;
            }
            catch (Exception ex)
            {
                Message = ex.Message;
                return false;
            }
        }

        private readonly JsonSerializerOptions options = new() { PropertyNameCaseInsensitive = true, WriteIndented = true };
        private EmailSettings() { }
        private bool InitSettings(string settingFileName)
        {
            if (!File.Exists(settingFileName))
            {
                Message = $"Settings file does not exist for Email: {settingFileName}";
                return false;
            }
            try
            {
                string jsonString = File.ReadAllText(settingFileName);
                Values? values = JsonSerializer.Deserialize<Values>(jsonString, options);
                if (values == null)
                {
                    Message = $"No settings was found in {settingFileName}";
                    return false;
                }
                _values = values;
            }
            catch (Exception ex)
            {
                Message = ex.Message;
                return false;
            }

            return CheckForValues();
        }

        bool CheckForValues()
        {
            bool res = true;
            if (_values.Attachments == null)
            {
                _values = _values with { Attachments = [] };
            }

            if (string.IsNullOrEmpty(_values.EmailAddress))
            {
                res = false;
                Message += "EmailAddress is missing from file\n";
            }
            if (string.IsNullOrEmpty(_values.EmailAuthor))
            {
                res = false;
                Message += "EmailAuthor is missing from file\n";
            }
            if (string.IsNullOrEmpty(_values.ClientId))
            {
                res = false;
                Message += "ClientId is missing from file\n";
            }
            if (string.IsNullOrEmpty(_values.TenantId))
            {
                res = false;
                Message += "TenantID is missing from file\n";
            }
            if (string.IsNullOrEmpty(_values.TemplateFileName))
            {
                res = false;
                Message += "TemplateFileName is missing from file\n";
            }
            if (string.IsNullOrEmpty(_values.InputFileName))
            {
                res = false;
                Message += "InputFileName is missing from file\n";
            }
            if (string.IsNullOrEmpty(_values.MailSubject))
            {
                res = false;
                Message += "MailSubject is missing from file\n";
            }
            return res;
        }

    }
}
