using Microsoft.Win32;
using System.Windows;

namespace EmailWithAttachedFile
{
    public partial class EmailSettingsWindow : Window
    {
        public EmailSettingsWindow()
        {
            InitializeComponent();
            LoadSettings();
        }

        private void LoadSettings()
        {
            txtTenantId.Text = EmailSettings.TenantId;
            txtClientId.Text = EmailSettings.ClientId;
            txtEmailAddress.Text = EmailSettings.EmailAddress;
            txtEmailAuthor.Text = EmailSettings.EmailAuthor;
            txtTemplateFile.Text = EmailSettings.TemplateFileName;
            txtInputFile.Text = EmailSettings.InputFileName;
            txtMailSubject.Text = EmailSettings.MailSubject;
            textSettingsFileName.Text = $"Settings file: {EmailSettings.SettingsFileName}";

            attachements.Items.Clear();
            foreach (string file in EmailSettings.Attachments)
            {
                attachements.Items.Add(file);
            }
        }

        private void BtnBrowseTemplate_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new()
            {
                DefaultExt = ".txt",
                Filter = "Text Files (*.txt)|*.txt|All files (*.*)|*.*"
            };
            if (dlg.ShowDialog() == true)
                txtTemplateFile.Text = dlg.FileName;
        }

        private void BtnBrowseInput_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new()
            {
                DefaultExt = ".csv",
                Filter = "CSV Files (*.csv)|*.csv|All files (*.*)|*.*"
            };
            if (dlg.ShowDialog() == true)
                txtInputFile.Text = dlg.FileName;
        }

        private void BtnAddAttachment_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new()
            {
                Filter = "All files (*.*)|*.*"
            };
            if (dlg.ShowDialog() == true)
                attachements.Items.Add(dlg.FileName);
        }

        private void BtnRemoveAttachment_Click(object sender, RoutedEventArgs e)
        {
            if (attachements.SelectedItem != null)
            {
                attachements.Items.Remove(attachements.SelectedItem);
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
#pragma warning disable IDE0305
            List<string> attachments = attachements.Items.Cast<string>().ToList();
#pragma warning restore
            EmailSettings.Update(
                txtTenantId.Text,
                txtClientId.Text,
                txtEmailAddress.Text,
                txtEmailAuthor.Text,
                txtTemplateFile.Text,
                txtInputFile.Text,
                txtMailSubject.Text,
                attachments);

            if (EmailSettings.Save())
            {
                this.DialogResult = true;
                this.Close();
            }
            else
            {
                MessageBox.Show($"Error saving settings: {EmailSettings.Message}", "Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}