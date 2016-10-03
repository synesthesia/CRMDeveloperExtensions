using System;
using System.Windows.Forms;

namespace UserOptions
{
    public partial class WrdOptionsControl : UserControl
    {
        public WrdOptionsControl()
        {
            InitializeComponent();
        }

        internal WrdOptionPageCustom AllowPublishManagedWebResources;
        internal WrdOptionPageCustom EnableCrmWrContextTemplates;

        public void Initialize()
        {
            AllowPublishManaged.Checked = AllowPublishManagedWebResources.AllowPublishManagedWebResources;
            EnableWrContextTemplates.Checked = EnableCrmWrContextTemplates.EnableCrmWrContextTemplates;
        }

        private void AllowPublishManaged_CheckedChanged(object sender, EventArgs e)
        {
            AllowPublishManagedWebResources.AllowPublishManagedWebResources = AllowPublishManaged.Checked;
        }

        private void EnableWrContextTemplates_CheckedChanged(object sender, EventArgs e)
        {
            EnableCrmWrContextTemplates.EnableCrmWrContextTemplates = EnableWrContextTemplates.Checked;
        }
    }
}
