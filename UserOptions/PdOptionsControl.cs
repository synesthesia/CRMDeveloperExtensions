using System;
using System.IO;
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

        private void PrtName_TextChanged(object sender, EventArgs e)
        {
            string path = PrtName.Text.Trim();

            if (string.IsNullOrEmpty(path))
            {
                RegistraionToolPath.RegistrationToolPath = null;
                return;
            }

            if (path.EndsWith(".exe", StringComparison.CurrentCultureIgnoreCase))
                path = Path.GetDirectoryName(path);

            if (path != null && !path.EndsWith("\\"))
                path += "\\";

            RegistraionToolPath.RegistrationToolPath = path;
        }
    }
}
