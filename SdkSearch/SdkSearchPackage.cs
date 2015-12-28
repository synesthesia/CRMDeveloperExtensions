using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.Net;
using System.Runtime.InteropServices;

namespace SdkSearch
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.GuidSdkSearchPkgString)]
    [ProvideAutoLoad("f1536ef8-92ec-443c-9ed7-fdadf150da82")]
    public sealed class SdkSearchPackage : Package
    {
        private DTE _dte;

        protected override void Initialize()
        {
            base.Initialize();

            _dte = GetGlobalService(typeof(DTE)) as DTE;
            if (_dte == null)
                return;

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (mcs != null)
            {
                // Create the command for the menu item.
                CommandID sdkSearchCommandId = new CommandID(GuidList.GuidSdkSearchCmdSet, (int)PkgCmdIdList.CmdidCrmSdkSearch);
                OleMenuCommand sdkSearchMenuItem = new OleMenuCommand(SearchMenuItemCallback, sdkSearchCommandId);
                sdkSearchMenuItem.BeforeQueryStatus += SdkSearch_BeforeQueryStatus;
                sdkSearchMenuItem.Visible = false;
                mcs.AddCommand(sdkSearchMenuItem);
            }
        }

        private void SdkSearch_BeforeQueryStatus(object sender, EventArgs eventArgs)
        {
            OleMenuCommand menuCommand = sender as OleMenuCommand;
            if (menuCommand == null) return;

            //Check SDK Search is enabled
            var props = _dte.Properties["CRM Developer Extensions", "General"];
            bool enableCrmSdkSearch = (bool)props.Item("EnableCrmSdkSearch").Value;

            if (!enableCrmSdkSearch)
            {
                menuCommand.Visible = false;
                return;
            }

            //Check for selected text
            TextSelection selection = (TextSelection)_dte.ActiveDocument.Selection;
            string searchText = selection.Text;
            if (string.IsNullOrEmpty(searchText))
            {
                menuCommand.Visible = false;
                return;
            }

            menuCommand.Visible = true;
        }

        private void SearchMenuItemCallback(object sender, EventArgs e)
        {
            TextSelection selection = (TextSelection)_dte.ActiveDocument.Selection;
            string searchText = selection.Text;

            if (string.IsNullOrEmpty(searchText)) return;

            var props = _dte.Properties["CRM Developer Extensions", "General"];
            bool useDefaultWebBrowser = (bool)props.Item("UseDefaultWebBrowser").Value;

            string url =
                "https://social.msdn.microsoft.com/Search/en-US/dynamics/crm?query=" + WebUtility.UrlEncode(searchText) + "&Refinement=241";

            if (useDefaultWebBrowser) //User's default browser
                System.Diagnostics.Process.Start(url);
            else //Internal VS browser
                _dte.ItemOperations.Navigate(url);
        }
    }
}
