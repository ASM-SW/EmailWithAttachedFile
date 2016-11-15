﻿using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace EmailWithAttachedFile
{
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
            public int MaxCount { get; set; }
            public int CountComplete { get; set; }
            public string NameComplete { get; set; }
            public string ErrorMessage { get; set; }
            public bool IsOk { get; set; }
        }

        private ConfigurationEmailWAF m_configuration = new ConfigurationEmailWAF();
        private BackgroundWorker m_bgWorker = new BackgroundWorker();
        private EmailSender m_emailSender = new EmailSender();

        private void buttonStart_Click(object sender, RoutedEventArgs e)
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

        private void GetValuesFromForm()
        {
            m_configuration.FromEmail = textFromEmail.Text;
            m_configuration.SmtpServer = textSmtpServer.Text;

            int port = 0;
            int.TryParse(textSmtpPort.Text, out port);

            m_configuration.SmtpPort = port;
            m_configuration.SmtpEnabledSSL = (bool)checkSmtpEnableSsl.IsChecked;
            m_configuration.Password = password.SecurePassword;
            m_configuration.TemplateFileName = textMessageTemplateFileName.Text;
            m_configuration.InputFileName = textInputName.Text;
            m_configuration.MailSubject = textMailSubject.Text;
        }

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

        private bool CheckString(string value, string name, ref StringBuilder errMsg)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                errMsg.AppendFormat("\t{0} is empty.\n", name);
                return false;
            }
            return true;
        }
        private void StartupBackgroudWorker()
        {
            progressBar.Value = 0;
            progressText.Content = string.Empty;

            StringBuilder errMsg;
            if(!m_emailSender.Init(ref m_configuration, out errMsg))
            {
                MessageBox.Show(errMsg.ToString());
                return;
            }

            m_bgWorker = new BackgroundWorker();
            m_bgWorker.WorkerReportsProgress = true;
            m_bgWorker.DoWork += worker_DoWork;
            m_bgWorker.ProgressChanged += worker_ProgressChanged;
            m_bgWorker.RunWorkerCompleted += worker_RunWorkerCompleted;
            m_bgWorker.WorkerSupportsCancellation = true;
            m_bgWorker.RunWorkerAsync(null);
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            DataTable inputData;
            ReadInputFile(out inputData);
            if (m_bgWorker.CancellationPending)
                return;

            ResultObject results = new ResultObject();
            results.MaxCount = inputData.Rows.Count;
            results.CountComplete = 0;

            foreach (DataRow row in inputData.Rows)
            {
                if (m_bgWorker.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
                // send the email
                string errMsg;
                results.IsOk = m_emailSender.SendMail(row["Name"].ToString(), row["Email"].ToString(), row["FileName"].ToString(), out errMsg);
                results.ErrorMessage = errMsg;

                results.NameComplete = row["Name"].ToString();
                ++results.CountComplete;
                int progressPercentage = Convert.ToInt32(((double)results.CountComplete / results.MaxCount) * 100);
                (sender as BackgroundWorker).ReportProgress(progressPercentage, results);
                e.Result = results;

            }
            buttonStart.IsEnabled = true;
        }


        private void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
            ResultObject results = e.UserState as ResultObject;
            if (results != null)
            {
                progressText.Content = string.Format("{0} of {1} complete", results.CountComplete, results.MaxCount);
                if(results.IsOk)
                    Log(results.NameComplete + " Sent");
                else
                {
                    Log(results.ErrorMessage);
                }
            }
        }

        private void Log(string msg)
        {
            listLog.Items.Add(msg);
            listLog.Items.MoveCurrentToLast();
            listLog.ScrollIntoView(listLog.Items.CurrentItem);
        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            buttonStop.IsEnabled = false;
            buttonStart.IsEnabled = true;
        }
        private void buttonMessageTemplate_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".txt";
            dlg.Filter = "Text Files (*.txt)|*.txt|All files (*.*)|*.*";
            Nullable<bool> result = dlg.ShowDialog();

            if (result == true)
            {
                string filename = dlg.FileName;
                textMessageTemplateFileName.Text = filename;
            }
        }

        // check to see if text is a string
        private static bool IsInteger(string text)
        {
            Regex regex = new Regex("[0-9]+");
            return regex.IsMatch(text);
        }

        // limit input to a integer during past event
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

        // limit input to an integer during text entry
        private void IntegerOnly(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsInteger(e.Text);
        }

        private void buttonInputFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".csv";
            dlg.Filter = "CSV Files (*.txt)|*.csv|All files (*.*)|*.*";
            Nullable<bool> result = dlg.ShowDialog();

            if (result == true)
            {
                string filename = dlg.FileName;
                textInputName.Text = filename;
            }
        }
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

        private void buttonStop_Click(object sender, RoutedEventArgs e)
        {
            buttonStart.IsEnabled = true;
            buttonStop.IsEnabled = false;

            if (m_bgWorker == null)
                return;
            if (m_bgWorker.IsBusy)
                m_bgWorker.CancelAsync();
        }

        private void MainFormClosing(object sender, CancelEventArgs e)
        {
            GetValuesFromForm();
            m_configuration.Serialize(m_configuration.ConfigFileName);
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
