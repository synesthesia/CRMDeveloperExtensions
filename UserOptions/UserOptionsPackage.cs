using Microsoft.VisualStudio.Shell;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace UserOptions
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideMenuResource("Menus.ctmenu", 1), Guid(GuidList.GuidUserOptionsPkgString)]
    [ProvideOptionPage(typeof(OptionPageCustom), "CRM Developer Extensions", "General", 101, 102, true)]
    public sealed class UserOptionsPackage : Package
    {
    }

    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("1D9ECCF3-5D2F-4112-9B25-264596873DC9")]
    public class OptionPageCustom : DialogPage
    {
        private string _defaultCrmSdkVersion = "CRM 2016 (8.2.X)";
        private string _defaultProjectKeyFileName = "MyKey";
        private bool _enableCrmSdkSearch = true;

        public OptionPageCustom()
        {
            EnableXrmToolingLogging = false;
        }

        public string DefaultCrmSdkVersion
        {
            get { return _defaultCrmSdkVersion; }
            set { _defaultCrmSdkVersion = value; }
        }

        public string DefaultProjectKeyFileName
        {
            get { return _defaultProjectKeyFileName; }
            set { _defaultProjectKeyFileName = value; }
        }

        public bool EnableCrmSdkSearch
        {
            get { return _enableCrmSdkSearch; }
            set { _enableCrmSdkSearch = value; }
        }

        public bool UseDefaultWebBrowser { get; set; }

        public bool EnableXrmToolingLogging { get; set; }

        public string XrmToolingLogPath { get; set; }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        protected override IWin32Window Window
        {
            get
            {
                OptionsControl page = new OptionsControl
                {
                    DefaultCrmSdkVersion = this,
                    DefaultProjectKeyFileName = this,
                    UseDefaultWebBrowser = this,
                    EnableCrmSdkSearch = this,
                    EnableXrmLogging = this,
                    XrmLogPath = this
                };
                page.Initialize();
                return page;
            }
        }
    }
}
