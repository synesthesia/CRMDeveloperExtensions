using System.Windows;
using System.Windows.Controls;

namespace TemplateWizards
{
    public partial class TypeScriptPicker
    {
        public string Version { get; set; }
        private readonly string _defaultSdkVersion;

        public TypeScriptPicker(string defaultSdkVersion)
        {
            _defaultSdkVersion = defaultSdkVersion;

            InitializeComponent();

            SetSdkVersion();
            ComboBoxItem itemSdk = (ComboBoxItem)SdkVersion.SelectedItem;
            Version = itemSdk.Content.ToString();
        }

        private void SetSdkVersion()
        {
            switch (_defaultSdkVersion)
            {
                case "CRM 2011 (5.0.X)":
                case "CRM 2013 (6.X.X)":
                case "CRM 2013 (6.1.X)":
                    SdkVersion.SelectedValue = "CRM 2013 (6.X.X)";
                    break;
                case "CRM 2015 (7.0.X)":
                    SdkVersion.SelectedValue = "CRM 2015 (7.0.X)";
                    break;
                case "CRM 2015 (7.1.X)":
                    SdkVersion.SelectedValue = "CRM 2015 (7.1.X)";
                    break;
                case "CRM 2016 (8.0.X)":
                case "CRM 2016 (8.1.X)":
                    SdkVersion.SelectedValue = "CRM 2016 (8.0.X)";
                    break;
            }
        }

        private void CreateTemplate_Click(object sender, RoutedEventArgs e)
        {
            ComboBoxItem itemVersion = (ComboBoxItem)SdkVersion.SelectedItem;
            Version = itemVersion.Content.ToString();

            DialogResult = true;
            Close();
        }
    }
}