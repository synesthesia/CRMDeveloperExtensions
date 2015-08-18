using Microsoft.VisualStudio.Shell;

namespace PluginDeployer
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    //[ProvideMenuResource("Menus.ctmenu", 1)]
    //[ProvideToolWindow(typeof(MyToolWindow))]
    //[Guid(GuidList.GuidPluginDeployerPkgString)]
    //[ProvideAutoLoad("f1536ef8-92ec-443c-9ed7-fdadf150da82")]
    public sealed class PluginDeployerPackage : Package
    {
        protected override void Initialize()
        {
            base.Initialize();
        }
    }
}
