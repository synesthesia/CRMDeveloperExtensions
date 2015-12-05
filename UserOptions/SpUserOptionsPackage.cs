using Microsoft.VisualStudio.Shell;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace UserOptions
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideMenuResource("Menus.ctmenu", 1), Guid(GuidList.GuidSpUserOptionsPkgString)]
    [ProvideOptionPage(typeof(SpOptionPageCustom), "CRM Developer Extensions", "Solution Packager", 101, 106, true)]
    public sealed class SpUserOptionsPackage : Package
    {
    }

    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("E65885EA-3A57-4A39-BC7B-7BDD897E9BBB")]
    public class SpOptionPageCustom : DialogPage
    {
        public string SolutionPackagerPath { get; set; }
        public bool SaveSolutionFiles { get; set; }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        protected override IWin32Window Window
        {
            get
            {
                SpOptionsControl page = new SpOptionsControl
                {
                    SolutionPackagerPath = this,
                    SaveSolutionFiles = this
                };
                page.Initialize();
                return page;
            }
        }
    }
}
