using System;
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

        private void SaveSolution_CheckedChanged(object sender, EventArgs e)
        {
            SaveSolutionFiles.SaveSolutionFiles = SaveSolution.Checked;
        }

        private void OpenFolder_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog.ShowDialog();
            if (result != DialogResult.OK) 
                return;

            string path = folderBrowserDialog.SelectedPath;

            if (string.IsNullOrEmpty(path))
            {
                SolutionPackagerPath.SolutionPackagerPath = null;
                SpName.Text = null;
                return;
            }

            if (!path.EndsWith("\\"))
                path += "\\";

            SolutionPackagerPath.SolutionPackagerPath = path;
            SpName.Text = path;
        }
    }
}
