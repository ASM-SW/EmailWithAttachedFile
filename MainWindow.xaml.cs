// Copyright © 2016-2020  ASM-SW
//asmeyers@outlook.com  https://github.com/asm-sw

using Microsoft.VisualBasic.FileIO;
using System;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace EmailWithAttachedFile
{
    enum MsgStatus
    {
        NotSent=0,
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

            ConfigurationEmailWAF.DeSerialize(m_configuration.ConfigFileName, ref m_configuration);

            textFromEmail.Text = m_configuration.FromEmail;
            textSmtpServer.Text = m_configuration.SmtpServer;
            textSmtpPort.Text = m_configuration.SmtpPort.ToString();
            checkSmtpEnableSsl.IsChecked = m_configuration.SmtpEnabledSSL;
            textMessageTemplateFileName.Text = m_configuration.TemplateFileName;
            textInputName.Text = m_configuration.InputFileName;
            textMailSubject.Text = m_configuration.MailSubject;
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

        private ConfigurationEmailWAF m_configuration = new ConfigurationEmailWAF();
        private BackgroundWorker m_bgWorker = new BackgroundWorker();
        private EmailSender m_emailSender = new EmailSender();
        private string m_outputFileName = string.Empty;

        private void ButtonStart_Click(object sender, RoutedEventArgs e)
        {
            GetValuesFromForm();

            if (CheckConfiguration())
            {
                buttonStart.IsEnabled = false;
                buttonStop.IsEnabled = true;
                listLog.Items.Clear();

                StartupBackgroudWorker();
            }
        }

        /// <summary>
        /// reads values from the form (user input) and puts them into the configuration object
        /// </summary>
        private void GetValuesFromForm()
        {
            m_configuration.FromEmail = textFromEmail.Text;
            m_configuration.SmtpServer = textSmtpServer.Text;

            int.TryParse(textSmtpPort.Text, out int port);

            m_configuration.SmtpPort = port;
            m_configuration.SmtpEnabledSSL = (bool)checkSmtpEnableSsl.IsChecked;
            m_configuration.Password = passwordBox.SecurePassword;
            m_configuration.TemplateFileName = textMessageTemplateFileName.Text;
            m_configuration.InputFileName = textInputName.Text;
            m_configuration.MailSubject = textMailSubject.Text;
        }

        /// <summary>
        /// Checks that the user input is OK.  Puts up a message box on error
        /// </summary>
        /// <returns>false if there was an error</returns>
        private bool CheckConfiguration()
        {
            bool isOk = true;
            StringBuilder errMsg = new StringBuilder("ERROR:\n");
            isOk &= CheckString(m_configuration.FromEmail, "FromEmail", ref errMsg);
            isOk &= CheckString(m_configuration.SmtpServer, "SMTP Server", ref errMsg);

            if (m_configuration.SmtpPort < 1 || m_configuration.SmtpPort > 65535)
            {
                errMsg.AppendLine("\tSMTP Port number must be 1 to 65535");
                isOk = false;
            }

            if (m_configuration.Password.Length < 1)
            {
                errMsg.AppendLine("\tPassword not entered");
                isOk = false;
            }

            isOk &= CheckFile(m_configuration.TemplateFileName, "Template File", ref errMsg);
            isOk &= CheckFile(m_configuration.InputFileName, "Input File", ref errMsg);

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
        private void StartupBackgroudWorker()
        {
            progressBar.Value = 0;
            progressText.Content = string.Empty;

            if (!m_emailSender.Init(ref m_configuration, out StringBuilder errMsg))
            {
                MessageBox.Show(errMsg.ToString());
                return;
            }

            m_bgWorker = new BackgroundWorker
            {
                WorkerReportsProgress = true
            };
            m_bgWorker.DoWork += Worker_DoWork;
            m_bgWorker.ProgressChanged += Worker_ProgressChanged;
            m_bgWorker.RunWorkerCompleted += Worker_RunWorkerCompleted;
            m_bgWorker.WorkerSupportsCancellation = true;
            m_bgWorker.RunWorkerAsync(null);
        }

        /// <summary>
        /// Background worker thread.  Reads the input file.  Sends email to each row in the file.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            DataTable inputData = new DataTable();
            ReadInputFile(out inputData);
            if (m_bgWorker.CancellationPending)
                return;

            inputData.Columns.Add("Status", typeof(string));
            inputData.Columns.Add("Message", typeof(string));
            foreach (DataRow row in inputData.Rows)
            {
                row["Status"] = MsgStatus.NotSent.ToString();
                row["Message"] = string.Empty;
            }

            ResultObject results = new ResultObject
            {
                MaxCount = inputData.Rows.Count,
                CountComplete = 0
            };

            foreach (DataRow row in inputData.Rows)
            {
                if (m_bgWorker.CancellationPending)
                {
                    e.Cancel = true;
                    break;
                }
                if (string.IsNullOrWhiteSpace(row["Email"].ToString()))
                {
                    results.IsOk = false;
                    results.ErrorMessage = "Email address is blank";
                    row["Status"] = MsgStatus.NoEmailAddress.ToString();
                    row["Message"] = results.ErrorMessage;
                }
                else
                {
                    // send the email
                    results.IsOk = m_emailSender.SendMail(row["Name"].ToString(), row["Email"].ToString(), row["FileName"].ToString(), out string errMsg);
                    results.ErrorMessage = errMsg;
                    if (results.IsOk)
                    {
                        row["Status"] = MsgStatus.Sent.ToString();
                    }
                    else
                    {
                        row["Status"] = MsgStatus.Error.ToString();
                        row["Message"] = results.ErrorMessage.Replace('\n', ';');
                    }
                }
                results.NameComplete = row["Name"].ToString();
                ++results.CountComplete;
                int progressPercentage = Convert.ToInt32(((double)results.CountComplete / results.MaxCount) * 100);
                (sender as BackgroundWorker).ReportProgress(progressPercentage, results);
                e.Result = results;

            }
            // write out results file
            m_outputFileName = Path.Combine(Path.GetDirectoryName(m_configuration.InputFileName),
                Path.GetFileNameWithoutExtension(m_configuration.InputFileName)) + "out.csv";
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(m_outputFileName))
            {
                file.Write(inputData.ToCSV());
            }
        }


        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
            if (e.UserState is ResultObject results)
            {
                progressText.Content = string.Format("{0} of {1} complete", results.CountComplete, results.MaxCount);
                if (results.IsOk)
                    Log("Sent: " + results.NameComplete);
                else
                {
                    Log("Not Sent: " + results.NameComplete + "\n\t" + results.ErrorMessage);
                }
            }
        }

        private void Log(string msg)
        {
            listLog.Items.Add(msg);
            listLog.Items.MoveCurrentToLast();
            listLog.ScrollIntoView(listLog.Items.CurrentItem);
        }

        void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Log("************ DONE *********************");
            Log("Check results in: " + m_outputFileName);
            Log("***************************************");
            buttonStop.IsEnabled = false;
            buttonStart.IsEnabled = true;
        }

        private void ButtonMessageTemplate_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog
            {
                // Set filter for file extension and default file extension 
                DefaultExt = ".txt",
                Filter = "Text Files (*.txt)|*.txt|All files (*.*)|*.*"
            };
            Nullable<bool> result = dlg.ShowDialog();

            if (result == true)
            {
                string filename = dlg.FileName;
                textMessageTemplateFileName.Text = filename;
            }
        }

        /// <summary>
        /// Checks to see if input is an integer
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private static bool IsInteger(string text)
        {
            Regex regex = new Regex("[0-9]+");
            return regex.IsMatch(text);
        }

        /// <summary>
        /// limit input to a integer during paste event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void IntegerOnlyPasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(String)))
            {
                String text = (String)e.DataObject.GetData(typeof(String));
                if (!IsInteger(text))
                {
                    e.CancelCommand();
                }
            }
            else
            {
                e.CancelCommand();
            }
        }

        /// <summary>
        /// limit input to an integer during text entry
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void IntegerOnly(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsInteger(e.Text);
        }

        private void ButtonInputFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog
            {

                // Set filter for file extension and default file extension 
                DefaultExt = ".csv",
                Filter = "CSV Files (*.txt)|*.csv|All files (*.*)|*.*"
            };
            Nullable<bool> result = dlg.ShowDialog();

            if (result == true)
            {
                string filename = dlg.FileName;
                textInputName.Text = filename;
            }
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
                using (TextFieldParser csvReader = new TextFieldParser(m_configuration.InputFileName))
                {
                    csvReader.SetDelimiters(new string[] { "," });
                    csvReader.HasFieldsEnclosedInQuotes = true;
                    string[] colFields = csvReader.ReadFields();
                    foreach (string item in colFields)
                    {
                        DataColumn column = new DataColumn(item);
                        inputData.Columns.Add(column);
                    }
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        inputData.Rows.Add(fieldData);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ButtonStop_Click(object sender, RoutedEventArgs e)
        {
            // button enable handled in worker_RunWorkerCompleted()
            //buttonStart.IsEnabled = true;
            //buttonStop.IsEnabled = false;

            if (m_bgWorker == null)
                return;
            if (m_bgWorker.IsBusy)
                m_bgWorker.CancelAsync();
        }

        /// <summary>
        /// Saves configuration when shutting down
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainFormClosing(object sender, CancelEventArgs e)
        {
            GetValuesFromForm();
            m_configuration.Serialize(m_configuration.ConfigFileName);
        }

        private void ShowPassword_Checked(object sender, RoutedEventArgs e)
        {
            passwordTxtBox.Text = passwordBox.Password;
            passwordBox.Visibility = Visibility.Collapsed;
            passwordTxtBox.Visibility = Visibility.Visible;
        }

        private void ShowPassword_Unchecked(object sender, RoutedEventArgs e)
        {
            passwordBox.Password = passwordTxtBox.Text;
            passwordTxtBox.Text = "";
            passwordTxtBox.Visibility = Visibility.Collapsed;
            passwordBox.Visibility = Visibility.Visible;
        }
        #region IDisposable Support
        private bool m_disposed = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!m_disposed)
            {
                if (disposing)
                {
                    m_bgWorker.Dispose();
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
