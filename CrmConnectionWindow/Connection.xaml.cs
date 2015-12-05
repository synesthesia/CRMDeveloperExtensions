using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;
using OutputLogger;
using System;
using System.ServiceModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace CrmConnectionWindow
{
    public partial class Connection
    {
        private readonly Logger _logger;

        public string ConnectionName { get; set; }
        public string ConnectionString { get; set; }
        public string OrgId { get; set; }
        public string Version { get; set; }

        public Connection(string name, string connectionString)
        {
            InitializeComponent();

            _logger = new Logger();

            if (!string.IsNullOrEmpty(name))
            {
                Name.Text = name;
                Url.IsEnabled = false;
                ConnectionType.IsEnabled = false;
            }

            if (!string.IsNullOrEmpty(connectionString))
            {
                ConnectionString = connectionString;
                ParseConnection(connectionString);
                SetConnectionString();

                return;
            }

            if (string.IsNullOrEmpty(connectionString))
                ConnectionType.SelectedIndex = 0;
        }

        private void ParseConnection(string connectionString)
        {
            if (connectionString.ToUpper().Contains("DYNAMICS.COM"))
                ConnectionType.SelectedIndex = 0;
            else if (!connectionString.ToUpper().Contains("USERNAME"))
                ConnectionType.SelectedIndex = 2;
            else if (Url.Text.Contains("."))
                ConnectionType.SelectedIndex = 3;
            else
                ConnectionType.SelectedIndex = 1;

            string[] parts = connectionString.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts)
            {
                string[] s = part.Split('=');
                if (s[0] == "Url")
                    Url.Text = s[1];
                if (s[0] == "Domain")
                    Domain.Text = s[1];
                if (s[0] == "Username")
                    Username.Text = s[1];
                if (s[0] == "Password")
                    Password.Password = s[1];
            }
        }

        private async void Connect_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(Name.Text))
            {
                MessageBox.Show("Enter A Name");
                return;
            }

            if (string.IsNullOrEmpty(Url.Text))
            {
                MessageBox.Show("Enter The Url");
                return;
            }

            ComboBoxItem item = (ComboBoxItem)ConnectionType.SelectedItem;
            if (item == null) return;

            if (item.Content.ToString() != "On-premises using Windows integrated security")
            {
                if (string.IsNullOrEmpty(Username.Text) || string.IsNullOrEmpty(Password.Password))
                {
                    MessageBox.Show("Enter The Username And Password");
                    return;
                }
            }

            string value = "Url=" + Url.Text.Trim() + ";";
            switch (item.Content.ToString())
            {
                case "Online using Office 365":
                    value += "Username=" + Username.Text.Trim() + ";";
                    value += "Password=" + Password.Password.Trim() + ";";
                    break;
                case "On-premises with provided user credentials":
                    if (!string.IsNullOrEmpty(Domain.Text))
                        value += "Domain=" + Domain.Text.Trim() + ";";
                    value += "Username=" + Username.Text.Trim() + ";";
                    value += "Password=" + Password.Password.Trim() + ";";
                    break;
                case "On-premises using Windows integrated security":
                    break;
                case "On-premises (IFD) with claims":
                    if (!string.IsNullOrEmpty(Domain.Text))
                        value += "Domain=" + Domain.Text.Trim() + ";";
                    value += "Username=" + Username.Text.Trim() + ";";
                    value += "Password=" + Password.Password.Trim() + ";";
                    break;
            }

            ConnectionString = value;

            LockOverlay.Visibility = Visibility.Visible;

            RetrieveVersionResponse vResponse = await System.Threading.Tasks.Task.Run(() => ConnectToCrm(ConnectionString));

            LockOverlay.Visibility = Visibility.Hidden;

            if (vResponse == null)
            {
                MessageBox.Show("Error Connecting to CRM. See the Output Window for additional details.");
                return;
            }

            Version = vResponse.Version;
            ConnectionName = Name.Text;

            DialogResult = true;
            Close();
        }

        private RetrieveVersionResponse ConnectToCrm(string connectionString)
        {
            try
            {
                CrmConnection connection = CrmConnection.Parse(connectionString);
                using (OrganizationService orgService = new OrganizationService(connection))
                {
                    WhoAmIRequest wRequest = new WhoAmIRequest();
                    WhoAmIResponse wResponse = (WhoAmIResponse)orgService.Execute(wRequest);
                    _logger.WriteToOutputWindow("Connected To CRM Organization: " + wResponse.OrganizationId, Logger.MessageType.Info);

                    OrgId = wResponse.OrganizationId.ToString();

                    RetrieveVersionRequest vRequest = new RetrieveVersionRequest();
                    RetrieveVersionResponse vResponse = (RetrieveVersionResponse)orgService.Execute(vRequest);
                    _logger.WriteToOutputWindow("Version: " + vResponse.Version, Logger.MessageType.Info);

                    return vResponse;
                }
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                _logger.WriteToOutputWindow("Error Connecting To CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, Logger.MessageType.Error);
                return null;
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Connecting To CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
                return null;
            }
        }

        private void ConnectionType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBoxItem item = (ComboBoxItem)ConnectionType.SelectedItem;
            if (item == null) return;

            switch (item.Content.ToString())
            {
                case "Online using Office 365":
                    Url.Text = "https://orgname.crm.dynamics.com";
                    UsernameLabel.Foreground = Brushes.Red;
                    PasswordLabel.Foreground = Brushes.Red;
                    Domain.IsEnabled = false;
                    Domain.Text = null;
                    Username.IsEnabled = true;
                    Password.IsEnabled = true;
                    break;
                case "On-premises with provided user credentials":
                    Url.Text = "http://servername/orgname";
                    DomainLabel.Foreground = Brushes.Black;
                    UsernameLabel.Foreground = Brushes.Red;
                    PasswordLabel.Foreground = Brushes.Red;
                    Domain.IsEnabled = true;
                    Username.IsEnabled = true;
                    Password.IsEnabled = true;
                    break;
                case "On-premises using Windows integrated security":
                    Url.Text = "http://servername/orgname";
                    DomainLabel.Foreground = Brushes.Black;
                    UsernameLabel.Foreground = Brushes.Black;
                    PasswordLabel.Foreground = Brushes.Black;
                    Domain.IsEnabled = false;
                    Domain.Text = null;
                    Username.IsEnabled = false;
                    Username.Text = null;
                    Password.IsEnabled = false;
                    Password.Password = null;
                    break;
                case "On-premises (IFD) with claims":
                    Url.Text = "https://orgname.domain.com";
                    DomainLabel.Foreground = Brushes.Black;
                    UsernameLabel.Foreground = Brushes.Red;
                    PasswordLabel.Foreground = Brushes.Red;
                    Domain.IsEnabled = true;
                    Username.IsEnabled = true;
                    Password.IsEnabled = true;
                    break;
            }

            SetConnectionString();
        }

        private void SetConnectionString()
        {
            string value = "Url=" + Url.Text.Trim() + ";";

            ComboBoxItem item = (ComboBoxItem)ConnectionType.SelectedItem;
            if (item == null) return;

            switch (item.Content.ToString())
            {
                case "Online using Office 365":
                    value += "Username=" + Username.Text.Trim() + ";";
                    value += "Password=" + new String('*', Password.Password.Trim().Length) + ";";
                    break;
                case "On-premises with provided user credentials":
                    if (!string.IsNullOrEmpty(Domain.Text))
                        value += "Domain=" + Domain.Text.Trim() + ";";
                    value += "Username=" + Username.Text.Trim() + ";";
                    value += "Password=" + new String('*', Password.Password.Trim().Length) + ";";
                    break;
                case "On-premises using Windows integrated security":
                    break;
                case "On-premises (IFD) with claims":
                    if (!string.IsNullOrEmpty(Domain.Text))
                        value += "Domain=" + Domain.Text.Trim() + ";";
                    value += "Username=" + Username.Text.Trim() + ";";
                    value += "Password=" + new String('*', Password.Password.Trim().Length) + ";";
                    break;
            }

            ConnString.Text = value;
        }

        private void Url_TextChanged(object sender, TextChangedEventArgs e)
        {
            SetConnectionString();
        }

        private void Domain_TextChanged(object sender, TextChangedEventArgs e)
        {
            SetConnectionString();
        }

        private void Username_TextChanged(object sender, TextChangedEventArgs e)
        {
            SetConnectionString();
        }

        private void Password_PasswordChanged(object sender, RoutedEventArgs e)
        {
            SetConnectionString();
        }
    }
}
