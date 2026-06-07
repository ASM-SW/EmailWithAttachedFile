// Copyright © 2016-2022  ASM-SW
//asmeyers@outlook.com  https://github.com/asm-sw

using Microsoft.VisualBasic.FileIO;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Windows;

namespace EmailWithAttachedFile
{
    enum MsgStatus
    {
        NotSent = 0,
        Sent,
        Error,
        NoEmailAddress
    }

    public class EmailJob
    {
        public string Name { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, IDisposable
    {
        public MainWindow()
        {
            InitializeComponent();
            InitializeSettings();
        }

        private void InitializeSettings()
        {
            if (!EmailSettings.Init())
            {
                MessageBox.Show($"Error: {EmailSettings.Message}");
            }
            UpdateUiFromSettings();
        }

        private void UpdateUiFromSettings()
        {
            textEmailAddress.Text = EmailSettings.EmailAddress;
            textEmailAuthor.Text = EmailSettings.EmailAuthor;
            textTemplateFileName.Text = EmailSettings.TemplateFileName;
            textInputFileName.Text = EmailSettings.InputFileName;
            textMailSubject.Text = EmailSettings.MailSubject;

            listAttachments.Items.Clear();
            foreach (string file in EmailSettings.Attachments)
            {
                listAttachments.Items.Add(file);
            }

            // Enable Start button only if the required settings are present
            buttonStart.IsEnabled = !string.IsNullOrWhiteSpace(EmailSettings.EmailAddress) &&
                                   !string.IsNullOrWhiteSpace(EmailSettings.TemplateFileName);
        }

        /// <summary>
        /// This object is used to pass result status from the background worker thread by the ReportProgress method
        /// </summary>
        class ResultObject
        {
            public ResultObject()
            {
                MaxCount = 0;
                CountComplete = 0;
                NameComplete = string.Empty;
                ErrorMessage = string.Empty;
                IsOk = true;
            }
            public int MaxCount { get; set; }           // Number of emails being sent
            public int CountComplete { get; set; }      // Number of emails completed
            public string NameComplete { get; set; }    // Name of the last email completed
            public string ErrorMessage { get; set; }    // Error message if there was an issue
            public bool IsOk { get; set; }              // false indicates there was an error
        }

        private readonly EmailSender m_emailSender = new();
        private string m_outputFileName = string.Empty;
        private CancellationTokenSource? m_cancellationTokenSource;

        private async void ButtonStart_Click(object sender, RoutedEventArgs e)
        {
            if (CheckConfiguration())
            {
                buttonStart.IsEnabled = false;
                buttonStop.IsEnabled = true;
                listLog.Items.Clear();

                await StartEmailProcessAsync();
            }
        }

        /// <summary>
        /// Checks that the user input is OK.  Puts up a message box on error
        /// </summary>
        /// <returns>false if there was an error</returns>
        private static bool CheckConfiguration()
        {
            bool isOk = true;
            StringBuilder errMsg = new("ERROR:\n");
            isOk &= CheckString(EmailSettings.EmailAddress, "EmailAddress", ref errMsg);
            isOk &= CheckFile(EmailSettings.TemplateFileName, "Template File", ref errMsg);
            isOk &= CheckFile(EmailSettings.InputFileName, "Input File", ref errMsg);

            if (!isOk)
                MessageBox.Show(errMsg.ToString());

            return isOk;
        }

        /// <summary>
        /// Checks if a file exists.
        /// </summary>
        /// <param name="fileName">name of file</param>
        /// <param name="name">user friendly name to put in error message</param>
        /// <param name="errMsg">string builder for error message.  Message are appended.</param>
        /// <returns>true if OK</returns>
        private static bool CheckFile(string fileName, string name, ref StringBuilder errMsg)
        {
            bool isOk = true;
            if (!CheckString(fileName, name, ref errMsg))
            {
                isOk = false;
            }
            else if (!File.Exists(fileName))
            {
                errMsg.AppendFormat("\t{0} file \"{1}\" does not exist.\n", name, fileName);
                isOk = false;
            }
            return isOk;
        }

        /// <summary>
        /// Checks to see if a string is null or whitespace.
        /// </summary>
        /// <param name="value">string to check</param>
        /// <param name="name">user friendly name to put in error message</param>
        /// <param name="errMsg">string builder for error message.  Message are appended.</param>
        /// <returns>true if OK</returns>
        private static bool CheckString(string value, string name, ref StringBuilder errMsg)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                errMsg.AppendFormat("\t{0} is empty.\n", name);
                return false;
            }
            return true;
        }

        /// <summary>
        /// Creates and setsups backgroud worker threader.  Inits mail sender.
        /// </summary>
        private async Task StartEmailProcessAsync()
        {
            progressBar.Value = 0;
            progressText.Content = string.Empty;

            m_cancellationTokenSource = new CancellationTokenSource();
            Progress<ResultObject> progress = new(UpdateUiProgress);

            try
            {
                (bool success, string errMsg) = await m_emailSender.InitAsync();
                if (!success)
                {
                    MessageBox.Show(errMsg);
                    return;
                }

                if (!string.IsNullOrEmpty(errMsg))
                    Log(errMsg);

                await Task.Run(() => DoEmailWorkAsync(progress, m_cancellationTokenSource.Token), m_cancellationTokenSource.Token);
            }
            catch (OperationCanceledException)
            {
                Log("Process was cancelled by the user.");
            }
            catch (Exception ex)
            {
                Log($"An unexpected error occurred: {ex.Message}");
            }
            finally
            {
                OnProcessCompleted();
            }
        }

        private async Task DoEmailWorkAsync(IProgress<ResultObject> progress, CancellationToken token)
        {
            List<EmailJob> inputData = ReadInputFile();
            ResultObject results = new() { MaxCount = inputData.Count };

            foreach (EmailJob job in inputData)
            {
                token.ThrowIfCancellationRequested();

                string name = job.Name;
                string emailAddr = job.Email;

                if (string.IsNullOrWhiteSpace(emailAddr))
                {
                    results.IsOk = false;
                    results.ErrorMessage = "Email address is blank";
                    job.Status = MsgStatus.NoEmailAddress.ToString();
                }
                else
                {
                   (bool success, string errMsg) = await m_emailSender.SendMailAsync(name, emailAddr, job.FileName);

                    results.IsOk = success;
                    results.ErrorMessage = errMsg;
                    job.Status = results.IsOk ? MsgStatus.Sent.ToString() : MsgStatus.Error.ToString();
                }

                job.Message = results.ErrorMessage.Replace('\n', ';');
                results.NameComplete = name;
                results.CountComplete++;
                progress.Report(results);
            }

            m_outputFileName = Path.Combine(Path.GetDirectoryName(EmailSettings.InputFileName) ?? "",
                Path.GetFileNameWithoutExtension(EmailSettings.InputFileName) + "out.csv");

            try
            {
                await File.WriteAllTextAsync(m_outputFileName, inputData.ToCSV(), token);
            }
            catch (Exception ex)
            {
                Log($"Failed to write output file: {ex.Message}");
            }
        }

        private void UpdateUiProgress(ResultObject results)
        {
            int progressPercentage = results.MaxCount > 0
                ? Convert.ToInt32(((double)results.CountComplete / results.MaxCount) * 100)
                : 0;

            progressBar.Value = progressPercentage;

            if (string.IsNullOrEmpty(results.NameComplete))
                Log(results.ErrorMessage);
            else
            {
                progressText.Content = $"{results.CountComplete} of {results.MaxCount} complete";
                Log(results.IsOk ? $"Sent: {results.NameComplete}" : $"Not Sent: {results.NameComplete}\n\t{results.ErrorMessage}");
            }
        }

        private void Log(string msg)
        {
            listLog.Items.Add(msg);
            listLog.Items.MoveCurrentToLast();
            listLog.ScrollIntoView(listLog.Items.CurrentItem);
        }

        private void OnProcessCompleted()
        {
            Log("************ DONE *********************");
            Log("Check results in: " + m_outputFileName);
            Log("***************************************");
            buttonStop.IsEnabled = false;
            buttonStart.IsEnabled = true;
        }

        private static readonly string[] _commaDelimiter = [","];

        /// <summary>
        /// Parser for reading the CSV input file.
        /// </summary>
        private static List<EmailJob> ReadInputFile()
        {
            List<EmailJob> jobs = [];
            try
            {
                using TextFieldParser csvReader = new(EmailSettings.InputFileName);
                csvReader.SetDelimiters(_commaDelimiter);
                csvReader.HasFieldsEnclosedInQuotes = true;

                string[] colFields = csvReader.ReadFields() ?? [];
                int nameIdx = Array.IndexOf(colFields, "Name");
                int emailIdx = Array.IndexOf(colFields, "Email");
                int fileIdx = Array.IndexOf(colFields, "FileName");

                List<string> missingColumns = [];
                if (nameIdx == -1) missingColumns.Add("Name");
                if (emailIdx == -1) missingColumns.Add("Email");
                if (fileIdx == -1) missingColumns.Add("FileName");

                if (missingColumns.Count > 0)
                {
                    MessageBox.Show($"The following required columns are missing from the input file: {string.Join(", ", missingColumns)}", "Input Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return jobs;
                }

                while (!csvReader.EndOfData)
                {
                    string[]? fieldData = csvReader.ReadFields();
                    if (fieldData != null)
                    {
                        jobs.Add(new EmailJob
                        {
                            // existence of columns checked above, so using the index should be safe
                            Name = fieldData[nameIdx],
                            Email = fieldData[emailIdx],
                            FileName = fieldData[fileIdx]
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return jobs;
        }

        private void ButtonSettings_Click(object sender, RoutedEventArgs e)
        {
            EmailSettingsWindow settingsWindow = new()
            {
                Owner = this
            };
            if (settingsWindow.ShowDialog() == true)
            {
                UpdateUiFromSettings();
            }
        }

        private void ButtonStop_Click(object sender, RoutedEventArgs e)
        {
            m_cancellationTokenSource?.Cancel();
        }

        /// <summary>
        /// Saves configuration when shutting down
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainFormClosing(object sender, CancelEventArgs e)
        {
            EmailSettings.Save();
        }

        #region IDisposable Support
        private bool m_disposed = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!m_disposed)
            {
                if (disposing)
                {
                    m_cancellationTokenSource?.Dispose();
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                m_disposed = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion

        private void ButtonExit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
