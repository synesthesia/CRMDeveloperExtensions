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
        internal OptionPageCustom EnableXrmLogging;
        internal OptionPageCustom XrmLogPath;

        public void Initialize()
        {
            DefaultSdkVersion.SelectedIndex = DefaultSdkVersion.FindStringExact(!string.IsNullOrEmpty(DefaultCrmSdkVersion.DefaultCrmSdkVersion)
                                                  ? DefaultCrmSdkVersion.DefaultCrmSdkVersion
                                                  : "CRM 2016 (8.2.X)");
            DefaultKeyFileName.Text = DefaultProjectKeyFileName.DefaultProjectKeyFileName;
            DefaultWebBrowser.Checked = UseDefaultWebBrowser.UseDefaultWebBrowser;
            EnableSdkSearch.Checked = EnableCrmSdkSearch.EnableCrmSdkSearch;
            EnableLogging.Checked = EnableXrmLogging.EnableXrmToolingLogging;
            LogPath.Text = XrmLogPath.XrmToolingLogPath;
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
            MessageBox.Show(@"Folder does not exist or unable to access");
            DefaultProjectKeyFileName.DefaultProjectKeyFileName = "MyKey";
            DefaultKeyFileName.Text = @"MyKey";
        }

        private void DefaultWebBrowser_CheckedChanged(object sender, EventArgs e)
        {
            UseDefaultWebBrowser.UseDefaultWebBrowser = DefaultWebBrowser.Checked;
        }

        private void EnableSdkSearch_CheckedChanged(object sender, EventArgs e)
        {
            EnableCrmSdkSearch.EnableCrmSdkSearch = EnableSdkSearch.Checked;
        }

        private void EnableLogging_CheckedChanged(object sender, EventArgs e)
        {
            EnableXrmLogging.EnableXrmToolingLogging = EnableLogging.Checked;
        }

        private void OpenFolder_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog.ShowDialog();
            if (result != DialogResult.OK)
                return;

            string path = folderBrowserDialog.SelectedPath;

            if (string.IsNullOrEmpty(path))
            {
                XrmLogPath.XrmToolingLogPath = null;
                LogPath.Text = null;
                return;
            }

            if (!path.EndsWith("\\"))
                path += "\\";

            XrmLogPath.XrmToolingLogPath = path;
            LogPath.Text = path;
        }
    }
}
