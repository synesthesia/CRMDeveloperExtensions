using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using OutputLogger;

namespace SolutionPackager
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.GuidSolutionPackagerPkgString)]
    [ProvideAutoLoad("f1536ef8-92ec-443c-9ed7-fdadf150da82")]
    public sealed class SolutionPackagerPackage : Package
    {
        //private DTE _dte;
        //private Logger _logger;

        protected override void Initialize()
        {
            base.Initialize();

            //_logger = new Logger();

            //_dte = GetGlobalService(typeof(DTE)) as DTE;
            //if (_dte == null)
            //    return;

            //OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            //if (mcs != null)
            //{
            //    CommandID toolwndCommandID = new CommandID(GuidList.GuidSolutionPackagerCmdSet, (int)PkgCmdIDList.CmdidSolutionPackager);
            //    MenuCommand menuToolWin = new MenuCommand(ShowToolWindow, toolwndCommandID);
            //    mcs.AddCommand(menuToolWin);
            //}
        }
    }
}
