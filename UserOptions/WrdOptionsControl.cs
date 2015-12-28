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

        public void Initialize()
        {
            AllowPublishManaged.Checked = AllowPublishManagedWebResources.AllowPublishManagedWebResources;
        }

        private void AllowPublishManaged_CheckedChanged(object sender, EventArgs e)
        {
            AllowPublishManagedWebResources.AllowPublishManagedWebResources = AllowPublishManaged.Checked;
        }
    }
}
