using Microsoft.VisualStudio.Shell;

namespace PluginDeployer
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    public sealed class PluginDeployerPackage : Package
    {
        protected override void Initialize()
        {
            base.Initialize();
        }
    }
}
