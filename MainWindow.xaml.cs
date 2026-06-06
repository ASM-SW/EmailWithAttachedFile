// Copyright © 2016-2022  ASM-SW
//asmeyers@outlook.com  https://github.com/asm-sw

using Microsoft.VisualBasic.FileIO;
using System.ComponentModel;
using System.Data;
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
            EmailSenderTest.EmailSettings.Init();
            UpdateUiFromSettings();
        }

        private void UpdateUiFromSettings()
        {
            textEmailAddress.Text = EmailSenderTest.EmailSettings.EmailAddress;
            textEmailAuthor.Text = EmailSenderTest.EmailSettings.EmailAuthor;
            textTemplateFileName.Text = EmailSenderTest.EmailSettings.TemplateFileName;
            textInputFileName.Text = EmailSenderTest.EmailSettings.InputFileName;
            textMailSubject.Text = EmailSenderTest.EmailSettings.MailSubject;

            listAttachments.Items.Clear();
            foreach (string file in EmailSenderTest.EmailSettings.Attachments)
            {
                listAttachments.Items.Add(file);
            }

            // Enable Start button only if the required settings are present
            buttonStart.IsEnabled = !string.IsNullOrWhiteSpace(EmailSenderTest.EmailSettings.EmailAddress) &&
                                   !string.IsNullOrWhiteSpace(EmailSenderTest.EmailSettings.TemplateFileName);
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

        private readonly EmailSender m_emailSender = new EmailSender();
        private string m_outputFileName = string.Empty;
        private CancellationTokenSource? m_cts;

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
        private bool CheckConfiguration()
        {
            bool isOk = true;
            StringBuilder errMsg = new StringBuilder("ERROR:\n");
            isOk &= CheckString(EmailSenderTest.EmailSettings.EmailAddress, "EmailAddress", ref errMsg);
            isOk &= CheckFile(EmailSenderTest.EmailSettings.TemplateFileName, "Template File", ref errMsg);
            isOk &= CheckFile(EmailSenderTest.EmailSettings.InputFileName, "Input File", ref errMsg);

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
        private bool CheckFile(string fileName, string name, ref StringBuilder errMsg)
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
        private bool CheckString(string value, string name, ref StringBuilder errMsg)
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

            m_cts = new CancellationTokenSource();
            Progress<ResultObject> progress = new Progress<ResultObject>(UpdateUiProgress);

            try
            {
                (bool success, string errMsg) initResult = await m_emailSender.InitAsync();
                if (!initResult.success)
                {
                    MessageBox.Show(initResult.errMsg);
                    return;
                }

                if (!string.IsNullOrEmpty(initResult.errMsg))
                    Log(initResult.errMsg);

                await Task.Run(() => DoEmailWork(progress, m_cts.Token), m_cts.Token);
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

        private void DoEmailWork(IProgress<ResultObject> progress, CancellationToken token)
        {
            AdditionalEmailAddrs additionalEmailAddrs = new AdditionalEmailAddrs();
            StringBuilder msgParseAdditionalEmailAddress = new StringBuilder();

            ReadInputFile(out DataTable inputData);
            inputData.Columns.Add("Status", typeof(string));
            inputData.Columns.Add("Message", typeof(string));

            ResultObject results = new ResultObject { MaxCount = inputData.Rows.Count };

            foreach (DataRow row in inputData.Rows)
            {
                token.ThrowIfCancellationRequested();

                string name = row["Name"].ToString() ?? "Unknown";
                string emailAddr = row["Email"].ToString() ?? string.Empty;

                if (string.IsNullOrWhiteSpace(emailAddr))
                {
                    results.IsOk = false;
                    results.ErrorMessage = "Email address is blank";
                    row["Status"] = MsgStatus.NoEmailAddress.ToString();
                }
                else
                {
                    if (additionalEmailAddrs.GetAddtionalEmailAddresses(emailAddr, out string emailAddrAdditional))
                        emailAddr = emailAddrAdditional;

                    // Note: We use .GetAwaiter().GetResult() here ONLY because we are inside Task.Run 
                    // on a background thread where it is safe to block.
                    (bool success, string errMsg) sendResult = m_emailSender.SendMailAsync(name, emailAddr, row["FileName"].ToString() ?? "").GetAwaiter().GetResult();

                    results.IsOk = sendResult.success;
                    results.ErrorMessage = sendResult.errMsg;
                    row["Status"] = results.IsOk ? MsgStatus.Sent.ToString() : MsgStatus.Error.ToString();
                }

                row["Message"] = results.ErrorMessage.Replace('\n', ';');
                results.NameComplete = name;
                results.CountComplete++;
                progress.Report(results);
            }

            m_outputFileName = Path.Combine(Path.GetDirectoryName(EmailSenderTest.EmailSettings.InputFileName) ?? "",
                Path.GetFileNameWithoutExtension(EmailSenderTest.EmailSettings.InputFileName) + "out.csv");

            File.WriteAllText(m_outputFileName, inputData.ToCSV());
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

        /// <summary>
        /// Parser for reading the CSV input file.
        /// </summary>
        /// <param name="inputData">DataTabel containing the read in data</param>
        private void ReadInputFile(out DataTable inputData)
        {
            inputData = new DataTable();
            try
            {
                TextFieldParser csvReader = new(EmailSenderTest.EmailSettings.InputFileName);
                csvReader.SetDelimiters(new string[] { "," });
                csvReader.HasFieldsEnclosedInQuotes = true;

                string[]? colFields = csvReader.ReadFields();
                if (colFields != null)
                {
                    foreach (string item in colFields)
                    {
                        inputData.Columns.Add(item);
                    }
                }

                while (!csvReader.EndOfData)
                {
                    string[]? fieldData = csvReader.ReadFields();
                    if (fieldData != null)
                    {
                        inputData.Rows.Add(fieldData);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ButtonSettings_Click(object sender, RoutedEventArgs e)
        {
            EmailSettingsWindow settingsWindow = new EmailSettingsWindow();
            settingsWindow.Owner = this;
            if (settingsWindow.ShowDialog() == true)
            {
                UpdateUiFromSettings();
            }
        }

        private void ButtonStop_Click(object sender, RoutedEventArgs e)
        {
            m_cts?.Cancel();
        }

        /// <summary>
        /// Saves configuration when shutting down
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainFormClosing(object sender, CancelEventArgs e)
        {
            EmailSenderTest.EmailSettings.Save();
        }

        #region IDisposable Support
        private bool m_disposed = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!m_disposed)
            {
                if (disposing)
                {
                    m_cts?.Dispose();
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                m_disposed = true;
            }
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
        #endregion

    }
}
