using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using PluginDeployer;
using ReportDeployer;
using SolutionPackager;
using System;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using CommonResources;
using WebResourceDeployer;

namespace CRMDeveloperExtensions
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.3.4.1", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideToolWindow(typeof(WrdWindow))]
    [ProvideToolWindow(typeof(ReportWindow))]
    [ProvideToolWindow(typeof(PluginWindow))]
    [ProvideToolWindow(typeof(SPWindow))]
    [Guid(GuidList.GuidCrmDeveloperExtensionsPkgString)]
    [ProvideAutoLoad("f1536ef8-92ec-443c-9ed7-fdadf150da82")]
    public sealed class CRMDeveloperExtensionsPackage : Package
    {
        private DTE _dte;

        protected override void Initialize()
        {
            base.Initialize();

            _dte = GetGlobalService(typeof(DTE)) as DTE;
            if (_dte == null)
                return;

            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (mcs == null) return;

            //Item Templates
            ItemTemplateCommands itemCommands = new ItemTemplateCommands(_dte);
            CommandID pluginMenuCommandId1 = new CommandID(GuidList.GuidItMenuCommandsCmdSet, (int)PkgCmdIdList.CmdidAddItem1);
            OleMenuCommand pluginMenuItem1 = new OleMenuCommand(itemCommands.MenuItem1Callback, pluginMenuCommandId1);
            pluginMenuItem1.BeforeQueryStatus += itemCommands.MenuItem1_BeforeQueryStatus;
            pluginMenuItem1.Visible = false;
            mcs.AddCommand(pluginMenuItem1);

            CommandID pluginMenuCommandId2 = new CommandID(GuidList.GuidItMenuCommandsCmdSet, (int)PkgCmdIdList.CmdidAddItem2);
            OleMenuCommand pluginMenuItem2 = new OleMenuCommand(itemCommands.MenuItem2Callback, pluginMenuCommandId2);
            pluginMenuItem2.BeforeQueryStatus += itemCommands.MenuItem2_BeforeQueryStatus;
            pluginMenuItem2.Visible = false;
            mcs.AddCommand(pluginMenuItem2);

            CommandID pluginMenuCommandId3 = new CommandID(GuidList.GuidItMenuCommandsCmdSet, (int)PkgCmdIdList.CmdidAddItem3);
            OleMenuCommand pluginMenuItem3 = new OleMenuCommand(itemCommands.MenuItem3Callback, pluginMenuCommandId3);
            pluginMenuItem3.BeforeQueryStatus += itemCommands.MenuItem3_BeforeQueryStatus;
            pluginMenuItem3.Visible = false;
            mcs.AddCommand(pluginMenuItem3);

            //Web Resource Deployer
            CommandID wrdWindowCommandId = new CommandID(GuidList.GuidCrmDevExCmdSet, (int)PkgCmdIdList.CmdidWebResourceDeployerWindow);
            OleMenuCommand wrdWindowItem = new OleMenuCommand(ShowWrdToolWindow, wrdWindowCommandId);
            mcs.AddCommand(wrdWindowItem);

            //Report Deployer
            CommandID reportWindowCommandId = new CommandID(GuidList.GuidCrmDevExCmdSet, (int)PkgCmdIdList.CmdidReportDeployerWindow);
            OleMenuCommand reportWindowItem = new OleMenuCommand(ShowReportToolWindow, reportWindowCommandId);
            mcs.AddCommand(reportWindowItem);

            //Plug-in Deployer
            CommandID pluginWindowCommandId = new CommandID(GuidList.GuidCrmDevExCmdSet, (int)PkgCmdIdList.CmdidPluginDeployerWindow);
            OleMenuCommand pluginWindowItem = new OleMenuCommand(ShowPluginToolWindow, pluginWindowCommandId);
            mcs.AddCommand(pluginWindowItem);

            //Solution Packager
            CommandID packagerWindowCommandId = new CommandID(GuidList.GuidCrmDevExCmdSet, (int)PkgCmdIdList.CmdidSolutionPackagerWindow);
            OleMenuCommand packageWindowItem = new OleMenuCommand(ShowSolutionPackagerToolWindow, packagerWindowCommandId);
            mcs.AddCommand(packageWindowItem);

            //Enable Logging
            SharedConnection.EnableLogging(_dte);
        }

        private void ShowWrdToolWindow(object sender, EventArgs e)
        {
            ToolWindowPane window = FindToolWindow(typeof(WrdWindow), 0, true);
            if ((null == window) || (null == window.Frame))
            {
                throw new NotSupportedException("Cannot create tool window.");
            }
            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }

        private void ShowReportToolWindow(object sender, EventArgs e)
        {
            ToolWindowPane window = FindToolWindow(typeof(ReportWindow), 0, true);
            if ((null == window) || (null == window.Frame))
            {
                throw new NotSupportedException("Cannot create tool window.");
            }
            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }

        private void ShowPluginToolWindow(object sender, EventArgs e)
        {
            ToolWindowPane window = FindToolWindow(typeof(PluginWindow), 0, true);
            if ((null == window) || (null == window.Frame))
            {
                throw new NotSupportedException("Cannot create tool window.");
            }
            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }

        private void ShowSolutionPackagerToolWindow(object sender, EventArgs e)
        {
            ToolWindowPane window = FindToolWindow(typeof(SPWindow), 0, true);
            if ((null == window) || (null == window.Frame))
            {
                throw new NotSupportedException("Cannot create tool window.");
            }
            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }
    }
}
