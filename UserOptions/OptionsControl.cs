using System;
using System.IO;
using System.Windows.Forms;

namespace UserOptions
{
    public partial class OptionsControl : UserControl
    {
        public OptionsControl()
        {
            InitializeComponent();
        }

        internal OptionPageCustom DefaultCrmSdkVersion;
        internal OptionPageCustom DefaultProjectKeyFileName;
        internal OptionPageCustom UseDefaultWebBrowser;
        internal OptionPageCustom EnableCrmSdkSearch;

        public void Initialize()
        {
            DefaultSdkVersion.SelectedIndex = DefaultSdkVersion.FindStringExact(!string.IsNullOrEmpty(DefaultCrmSdkVersion.DefaultCrmSdkVersion)
                                                  ? DefaultCrmSdkVersion.DefaultCrmSdkVersion
                                                  : "CRM 2016 (8.2.X)");
            DefaultKeyFileName.Text = DefaultProjectKeyFileName.DefaultProjectKeyFileName;
            DefaultWebBrowser.Checked = UseDefaultWebBrowser.UseDefaultWebBrowser;
            EnableSdkSearch.Checked = EnableCrmSdkSearch.EnableCrmSdkSearch;
        }

        private void DefaultSdkVersion_SelectedIndexChanged(object sender, EventArgs e)
        {
            DefaultCrmSdkVersion.DefaultCrmSdkVersion = DefaultSdkVersion.SelectedItem.ToString();
        }

        private void DefaultKeyFileName_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(DefaultKeyFileName.Text))
            {
                HandleIllegalFileName();
                return;
            }

            string name = DefaultKeyFileName.Text + ".snk";
            if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                HandleIllegalFileName();
                return;
            }

            if (name.StartsWith("."))
            {
                HandleIllegalFileName();
                return;
            }

            if (name.ToUpper() == "CON.SNK" || name.ToUpper() == "AUX.SNK" || name.ToUpper() == "PRN.SNK" || name.ToUpper() == "COM1.SNK" || name.ToUpper() == "LPT2.SNK")
            {
                HandleIllegalFileName();
                return;
            }

            DefaultProjectKeyFileName.DefaultProjectKeyFileName = DefaultKeyFileName.Text.Trim();
        }

        private void HandleIllegalFileName()
        {
            MessageBox.Show("Illegal file name");
            DefaultProjectKeyFileName.DefaultProjectKeyFileName = "MyKey";
            DefaultKeyFileName.Text = "MyKey";
        }

        private void DefaultWebBrowser_CheckedChanged(object sender, EventArgs e)
        {
            UseDefaultWebBrowser.UseDefaultWebBrowser = DefaultWebBrowser.Checked;
        }

        private void EnableSdkSearch_CheckedChanged(object sender, EventArgs e)
        {
            EnableCrmSdkSearch.EnableCrmSdkSearch = EnableSdkSearch.Checked;
        }
    }
}
