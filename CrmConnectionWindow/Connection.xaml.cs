using EnvDTE;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using OutputLogger;
using System;
using System.ServiceModel;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace CrmConnectionWindow
{
    public partial class Connection
    {
        private readonly Logger _logger;
        private readonly DTE _dte;
        private readonly string _windowType;
        public string ConnectionName { get; set; }
        public string ConnectionString { get; set; }
        public string OrgId { get; set; }
        public string Version { get; set; }

        public Connection(string name, string connectionString, string windowType, DTE dte)
        {
            InitializeComponent();

            _logger = new Logger();
            _dte = dte;
            _windowType = windowType;

            if (!string.IsNullOrEmpty(name))
            {
                Name.Text = name;
                Name.IsEnabled = false;
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
            if (connectionString.ToUpper().Contains("AUTHTYPE=OAUTH"))
                ConnectionType.SelectedIndex = 4; //OAuth
            else if (connectionString.ToUpper().Contains("AUTHTYPE=OFFICE365"))
                ConnectionType.SelectedIndex = 0; //Online
            else if (connectionString.ToUpper().Contains("AUTHTYPE=IFD"))
                ConnectionType.SelectedIndex = 3; //IFD 
            else //AUTHTYPE=AD"
            {
                if (connectionString.ToUpper().Contains("USERNAME"))
                    ConnectionType.SelectedIndex = 1; //Credentials
                else
                    ConnectionType.SelectedIndex = 2; //Windows
            }


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
                {
                    string password = s[1];
                    if (password.StartsWith("'"))
                        password = password.Substring(1, password.Length - 1);
                    if (password.EndsWith("'"))
                        password = password.Substring(0, password.Length - 1);

                    Password.Password = password;
                }
                if (s[0] == "AppId")
                    AppId.Text = s[1];
            }

            SetFormProperties(false);
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

            if (Domain.IsEnabled && string.IsNullOrEmpty(Domain.Text))
            {
                MessageBox.Show("Enter The Domain");
                return;
            }

            if (item.Content.ToString() != "On-premises using Windows integrated security")
            {
                if (string.IsNullOrEmpty(Username.Text) || string.IsNullOrEmpty(Password.Password))
                {
                    MessageBox.Show("Enter The Username And Password");
                    return;
                }
            }

            if (item.Content.ToString() == "OAuth")
            {
                if (string.IsNullOrEmpty(AppId.Text))
                {
                    MessageBox.Show("Enter The AppId");
                    return;
                }
            }

            ConnectionString = CreateConnectionString(item);

            LockOverlay.Visibility = Visibility.Visible;

            string vResponse = await System.Threading.Tasks.Task.Run(() => ConnectToCrm(ConnectionString));

            LockOverlay.Visibility = Visibility.Hidden;

            if (vResponse == null)
            {
                MessageBox.Show("Error Connecting to CRM. See the Output Window for additional details.");
                return;
            }

            Version = vResponse;
            ConnectionName = Name.Text;

            DialogResult = true;
            Close();
        }

        private string CreateConnectionString(ComboBoxItem item)
        {
            var sb = new StringBuilder();

            sb.AppendFormat("Url={0};", Url.Text.Trim());

            switch (item.Content.ToString())
            {
                case "Online using Office 365":
                    sb.AppendFormat("Username={0};Password='{1}';", Username.Text.Trim(), Password.Password.Trim().Replace("'", "''"));
                    sb.Append("AuthType=Office365;");
                    break;
                case "On-premises (IFD) with claims":
                    sb.AppendFormat("Domain={0};", Domain.Text.Trim());
                    sb.AppendFormat("Username={0};Password='{1}';", Username.Text.Trim(), Password.Password.Trim().Replace("'", "''"));
                    sb.Append("AuthType=IFD;");
                    break;
                case "On-premises with provided user credentials":
                    sb.AppendFormat("Domain={0};", Domain.Text.Trim());
                    sb.AppendFormat("Username={0};Password={1};", Username.Text.Trim(), Password.Password.Trim().Replace("'", "''"));
                    sb.Append("AuthType=AD;");
                    break;
                case "On-premises using Windows integrated security":
                    sb.Append("AuthType=AD;");
                    break;
                case "OAuth":
                    sb.AppendFormat("Username={0};Password='{1}';", Username.Text.Trim(), Password.Password.Trim().Replace("'", "''"));
                    sb.AppendFormat("AppId={0};", AppId.Text.Trim());
                    sb.Append("LoginPrompt=Never;");
                    sb.Append("AuthType=OAuth;");
                    break;
            }

            sb.Append("RequireNewInstance=true;");
            return sb.ToString();
        }

        private string ConnectToCrm(string connectionString)
        {
            try
            {
                CrmServiceClient client = new CrmServiceClient(connectionString);

                WhoAmIRequest wRequest = new WhoAmIRequest();
                WhoAmIResponse wResponse = (WhoAmIResponse)client.Execute(wRequest);
                _logger.WriteToOutputWindow("Connected To CRM Organization: " + wResponse.OrganizationId, Logger.MessageType.Info);
                _logger.WriteToOutputWindow("Version: " + client.ConnectedOrgVersion, Logger.MessageType.Info);

                OrgId = wResponse.OrganizationId.ToString();

                Globals globals = _dte.Globals;
                globals["CurrentConnection" + _windowType] = client;

                if (client.ConnectedOrgVersion != null)
                    return client.ConnectedOrgVersion.ToString();

                _logger.WriteToOutputWindow("Error Connecting To CRM: Unable to determine org. version", Logger.MessageType.Error);
                return null;
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
            SetFormProperties(true);
        }

        private void SetFormProperties(bool setDefaults)
        {
            ComboBoxItem item = (ComboBoxItem)ConnectionType.SelectedItem;
            if (item == null) return;

            switch (item.Content.ToString())
            {
                case "Online using Office 365":
                    if (setDefaults)
                    {
                        Url.Text = "https://orgname.crm.dynamics.com";
                        Username.Text = "administrator@orgname.onmicrosoft.com";
                        Password.Password = "********";
                        Domain.Text = null;
                        AppId.Text = null;
                    }
                    UsernameLabel.Foreground = Brushes.Red;
                    PasswordLabel.Foreground = Brushes.Red;
                    DomainLabel.Foreground = Brushes.Black;
                    Domain.IsEnabled = false;
                    Username.IsEnabled = true;
                    Password.IsEnabled = true;
                    AppId.IsEnabled = false;
                    AppIdLabel.Foreground = Brushes.Black;
                    break;
                case "On-premises with provided user credentials":
                    if (setDefaults)
                    {
                        Url.Text = "http://servername/orgname";
                        Username.Text = "username";
                        Password.Password = "********";
                        Domain.Text = "domain";
                        AppId.Text = null;
                    }
                    DomainLabel.Foreground = Brushes.Red;
                    UsernameLabel.Foreground = Brushes.Red;
                    PasswordLabel.Foreground = Brushes.Red;
                    Domain.IsEnabled = true;
                    Username.IsEnabled = true;
                    Password.IsEnabled = true;
                    AppId.IsEnabled = false;
                    AppIdLabel.Foreground = Brushes.Black;
                    break;
                case "On-premises using Windows integrated security":
                    if (setDefaults)
                    {
                        Url.Text = "http://servername/orgname";
                        Username.Text = null;
                        Password.Password = null;
                        Domain.Text = null;
                        AppId.Text = null;
                    }
                    DomainLabel.Foreground = Brushes.Black;
                    UsernameLabel.Foreground = Brushes.Black;
                    PasswordLabel.Foreground = Brushes.Black;
                    Domain.IsEnabled = false;
                    Username.IsEnabled = false;
                    Password.IsEnabled = false;
                    AppId.IsEnabled = false;
                    AppIdLabel.Foreground = Brushes.Black;
                    break;
                case "On-premises (IFD) with claims":
                    if (setDefaults)
                    {
                        Url.Text = "https://host.domain.com/orgname";
                        Username.Text = "domain\\username";
                        Password.Password = "********";
                        Domain.Text = "domain";
                        AppId.Text = null;
                    }
                    DomainLabel.Foreground = Brushes.Red;
                    UsernameLabel.Foreground = Brushes.Red;
                    PasswordLabel.Foreground = Brushes.Red;
                    Domain.IsEnabled = true;
                    Username.IsEnabled = true;
                    Password.IsEnabled = true;
                    AppId.IsEnabled = false;
                    AppIdLabel.Foreground = Brushes.Black;
                    break;
                case "OAuth":
                    if (setDefaults)
                    {
                        Url.Text = "https://orgname.crm.dynamics.com";
                        Username.Text = "administrator@orgname.onmicrosoft.com";
                        Password.Password = "********";
                        Domain.Text = null;
                        AppId.Text = "00000000-0000-0000-0000-000000000000";
                    }
                    UsernameLabel.Foreground = Brushes.Red;
                    PasswordLabel.Foreground = Brushes.Red;
                    DomainLabel.Foreground = Brushes.Black;
                    Domain.IsEnabled = false;
                    Username.IsEnabled = true;
                    Password.IsEnabled = true;
                    AppId.IsEnabled = true;
                    AppIdLabel.Foreground = Brushes.Red;
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
                    value += "AuthType=Office365;";
                    break;
                case "On-premises with provided user credentials":
                    value += "Domain=" + Domain.Text.Trim() + ";";
                    value += "Username=" + Username.Text.Trim() + ";";
                    value += "Password=" + new String('*', Password.Password.Trim().Length) + ";";
                    value += "AuthType=AD;";
                    break;
                case "On-premises using Windows integrated security":
                    value += "AuthType=AD;";
                    break;
                case "On-premises (IFD) with claims":
                    value += "Domain=" + Domain.Text.Trim() + ";";
                    value += "Username=" + Username.Text.Trim() + ";";
                    value += "Password=" + new String('*', Password.Password.Trim().Length) + ";";
                    value += "AuthType=IFD;";
                    break;
                case "OAuth":
                    value += "Username=" + Username.Text.Trim() + ";";
                    value += "Password=" + new String('*', Password.Password.Trim().Length) + ";";
                    value += "AppId=" + AppId.Text.Trim() + ";";
                    value += "LoginPrompt=Never;";
                    value += "AuthType=OAuth;";
                    break;
            }

            value += "RequireNewInstance=true;";
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

        private void AppId_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            SetConnectionString();
        }

        private void Textbox_OnGotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (sender is TextBox)
            {
                ((TextBox)sender).SelectAll();
                return;
            }
            if (sender is PasswordBox)
            {
                ((PasswordBox)sender).SelectAll();
            }
        }
    }
}
