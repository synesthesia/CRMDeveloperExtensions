using System;
using System.Windows.Forms;

namespace UserOptions
{
    public partial class RdOptionsControl : UserControl
    {
        public RdOptionsControl()
        {
            InitializeComponent();
        }

        internal RdOptionPageCustom AllowPublishManagedReports;

        public void Initialize()
        {
            AllowPublishManagedRpts.Checked = AllowPublishManagedReports.AllowPublishManagedReports;
        }

        private void AllowPublishManagedRpts_CheckedChanged(object sender, EventArgs e)
        {
            AllowPublishManagedReports.AllowPublishManagedReports = AllowPublishManagedRpts.Checked;
        }
    }
}
