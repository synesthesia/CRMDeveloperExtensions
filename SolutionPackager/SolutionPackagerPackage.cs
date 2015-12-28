using Microsoft.VisualStudio.Shell;
using System.Runtime.InteropServices;

namespace SolutionPackager
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.GuidSolutionPackagerPkgString)]
    [ProvideAutoLoad("f1536ef8-92ec-443c-9ed7-fdadf150da82")]
    public sealed class SolutionPackagerPackage : Package
    {
    }
}
