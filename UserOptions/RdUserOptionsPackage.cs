using Microsoft.VisualStudio.Shell;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace UserOptions
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideMenuResource("Menus.ctmenu", 1), Guid(GuidList.GuidRdUserOptionsPkgString)]
    [ProvideOptionPage(typeof(RdOptionPageCustom), "CRM Developer Extensions", "Report Deployer", 101, 103, true)]
    public sealed class RdUserOptionsPackage : Package
    {
    }

    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("10151FAF-9292-40D1-8C07-7D8EB3D6CEEF")]
    public class RdOptionPageCustom : DialogPage
    {
        public bool AllowPublishManagedReports { get; set; }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        protected override IWin32Window Window
        {
            get
            {
                RdOptionsControl page = new RdOptionsControl
                {
                    AllowPublishManagedReports = this
                };
                page.Initialize();
                return page;
            }
        }
    }
}
