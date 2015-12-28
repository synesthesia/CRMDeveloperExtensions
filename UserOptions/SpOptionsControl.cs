using System;
using System.IO;
using System.Windows.Forms;

namespace UserOptions
{
    public partial class SpOptionsControl : UserControl
    {
        public SpOptionsControl()
        {
            InitializeComponent();
        }

        internal SpOptionPageCustom SolutionPackagerPath;
        internal SpOptionPageCustom SaveSolutionFiles;


        public void Initialize()
        {
            SpName.Text = SolutionPackagerPath.SolutionPackagerPath;
            SaveSolution.Checked = SaveSolutionFiles.SaveSolutionFiles;
        }

        private void SpName_TextChanged(object sender, EventArgs e)
        {
            string path = SpName.Text.Trim();

            if (string.IsNullOrEmpty(path))
            {
                SolutionPackagerPath.SolutionPackagerPath = null;
                return;
            }

            if (path.EndsWith(".exe", StringComparison.CurrentCultureIgnoreCase))
                path = Path.GetDirectoryName(path);

            if (path != null && !path.EndsWith("\\"))
                path += "\\";

            SolutionPackagerPath.SolutionPackagerPath = path;
        }

        private void SaveSolution_CheckedChanged(object sender, EventArgs e)
        {
            SaveSolutionFiles.SaveSolutionFiles = SaveSolution.Checked;
        }
    }
}
