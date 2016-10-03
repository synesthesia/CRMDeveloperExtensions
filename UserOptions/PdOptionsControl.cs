using System;
using System.Windows.Forms;

namespace UserOptions
{
    public partial class PdOptionsControl : UserControl
    {
        public PdOptionsControl()
        {
            InitializeComponent();
        }

        internal PdOptionPageCustom RegistraionToolPath;

        public void Initialize()
        {
            PrtName.Text = RegistraionToolPath.RegistrationToolPath;
        }

        private void OpenFolder_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog.ShowDialog();
            if (result != DialogResult.OK)
                return;

            string path = folderBrowserDialog.SelectedPath;

            if (string.IsNullOrEmpty(path))
            {
                RegistraionToolPath.RegistrationToolPath = null;
                PrtName.Text = null;
                return;
            }

            if (!path.EndsWith("\\"))
                path += "\\";

            RegistraionToolPath.RegistrationToolPath = path;
            PrtName.Text = path;
        }
    }
}
