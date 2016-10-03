using Microsoft.VisualStudio.Shell;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace UserOptions
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideMenuResource("Menus.ctmenu", 1), Guid(GuidList.GuidPdUserOptionsPkgString)]
    [ProvideOptionPage(typeof(PdOptionPageCustom), "CRM Developer Extensions", "Plug-in Deployer", 101, 103, true)]
    public sealed class PdUserOptionsPackage : Package
    {
    }

    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("FA4FDB19-2A8F-448B-8A8A-F32744CD9DAB")]
    public class PdOptionPageCustom : DialogPage
    {
        public string RegistrationToolPath { get; set; }
        public bool EnableCrmPdContextTemplates { get; set; }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        protected override IWin32Window Window
        {
            get
            {
                PdOptionsControl page = new PdOptionsControl
                {
                    RegistraionToolPath = this,
                    EnableCrmPdContextTemplates = this
                };
                page.Initialize();
                return page;
            }
        }
    }
}
