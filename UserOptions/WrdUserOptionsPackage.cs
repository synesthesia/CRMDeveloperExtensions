using Microsoft.VisualStudio.Shell;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace UserOptions
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideMenuResource("Menus.ctmenu", 1), Guid(GuidList.GuidWrdUserOptionsPkgString)]
    [ProvideOptionPage(typeof(WrdOptionPageCustom), "CRM Developer Extensions", "Web Resource Deployer", 101, 103, true)]
    public sealed class WrdUserOptionsPackage : Package
    {
    }

    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("0E8BE1E8-9AC3-4035-B4CD-2CA9064873DE")]
    public class WrdOptionPageCustom : DialogPage
    {
        public bool AllowPublishManagedWebResources { get; set; }
        public bool EnableCrmWrContextTemplates { get; set; }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        protected override IWin32Window Window
        {
            get
            {
                WrdOptionsControl page = new WrdOptionsControl
                {
                    AllowPublishManagedWebResources = this,
                    EnableCrmWrContextTemplates = this
                };
                page.Initialize();
                return page;
            }
        }
    }
}
