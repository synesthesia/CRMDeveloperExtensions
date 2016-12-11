using CommonResources;
using CommonResources.Models;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Tooling.Connector;
using OutputLogger;
using System;
using System.ComponentModel.Design;
using System.IO;
using System.Runtime.InteropServices;
using System.ServiceModel;
using System.Xml;
using Window = EnvDTE.Window;

namespace ReportDeployer
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.GuidReportDeployerPkgString)]
    [ProvideAutoLoad("f1536ef8-92ec-443c-9ed7-fdadf150da82")]
    public sealed class ReportDeployerPackage : Package
    {
        private DTE _dte;
        private Logger _logger;
        private const string WindowType = "ReportDeployer";

        protected override void Initialize()
        {
            base.Initialize();

            _logger = new Logger();

            _dte = GetGlobalService(typeof(DTE)) as DTE;
            if (_dte == null)
                return;

            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (mcs != null)
            {
                CommandID publishCommandId = new CommandID(GuidList.GuidItemMenuCommandsCmdSet, (int)PkgCmdIdList.CmdidReportDeployerPublish);
                OleMenuCommand publishMenuItem = new OleMenuCommand(PublishItemCallback, publishCommandId);
                publishMenuItem.BeforeQueryStatus += PublishItem_BeforeQueryStatus;
                publishMenuItem.Visible = false;
                mcs.AddCommand(publishMenuItem);
            }
        }

        private void PublishItem_BeforeQueryStatus(object sender, EventArgs eventArgs)
        {
            OleMenuCommand menuCommand = sender as OleMenuCommand;
            if (menuCommand == null) return;

            bool windowOpen = false;
            foreach (Window window in _dte.Windows)
            {
                if (window.Caption != Resources.ResourceManager.GetString("ToolWindowTitle")) continue;

                windowOpen = window.Visible;
                break;
            }

            if (!windowOpen)
            {
                menuCommand.Visible = false;
                return;
            }

            if (_dte.SelectedItems.Count != 1)
            {
                menuCommand.Visible = false;
                return;
            }

            SelectedItem item = _dte.SelectedItems.Item(1);
            ProjectItem projectItem = item.ProjectItem;

            CrmConn selectedConnection = (CrmConn)SharedGlobals.GetGlobal("SelectedConnection", _dte);
            if (selectedConnection == null)
            {
                menuCommand.Visible = false;
                return;
            }

            Guid reportId = GetMapping(projectItem, selectedConnection);
            menuCommand.Visible = reportId != Guid.Empty;
        }

        private void PublishItemCallback(object sender, EventArgs e)
        {
            if (_dte.SelectedItems.Count != 1) return;

            SelectedItem item = _dte.SelectedItems.Item(1);
            ProjectItem projectItem = item.ProjectItem;

            CrmConn selectedConnection = (CrmConn)SharedGlobals.GetGlobal("SelectedConnection", _dte);
            if (selectedConnection == null) return;

            Guid reportId = GetMapping(projectItem, selectedConnection);
            if (reportId == Guid.Empty) return;

            CrmServiceClient client = SharedConnection.GetCurrentConnection(selectedConnection.ConnectionString, WindowType, _dte);

            UpdateAndPublishSingle(client, projectItem, reportId);
        }

        private void UpdateAndPublishSingle(CrmServiceClient client, ProjectItem projectItem, Guid reportId)
        {
            try
            {
                _dte.StatusBar.Text = "Deploying report...";
                _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationDeploy);

                Entity report = new Entity("report") { Id = reportId };
                if (!File.Exists(projectItem.FileNames[1])) return;

                report["bodytext"] = File.ReadAllText(projectItem.FileNames[1]);

                UpdateRequest request = new UpdateRequest { Target = report };
                client.Execute(request);
                _logger.WriteToOutputWindow("Deployed Report", Logger.MessageType.Info);
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                _logger.WriteToOutputWindow("Error Deploying Report To CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, Logger.MessageType.Error);
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Deploying Report To CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
            }

            _dte.StatusBar.Clear();
            _dte.StatusBar.Animate(false, vsStatusAnimation.vsStatusAnimationDeploy);
        }

        private Guid GetMapping(ProjectItem projectItem, CrmConn selectedConnection)
        {
            try
            {
                Project project = projectItem.ContainingProject;
                var projectPath = Path.GetDirectoryName(project.FullName);
                if (projectPath == null) return Guid.Empty;
                var boundName = projectItem.FileNames[1].Replace(projectPath, String.Empty).Replace("\\", "/");

                if (!File.Exists(projectItem.FileNames[1])) return Guid.Empty;

                var path = Path.GetDirectoryName(project.FullName);
                if (!SharedConfigFile.ConfigFileExists(project))
                {
                    _logger.WriteToOutputWindow("Error Getting Mapping: Missing CRMDeveloperExtensions.config File", Logger.MessageType.Error);
                    return Guid.Empty;
                }

                XmlDocument doc = new XmlDocument();
                doc.Load(path + "\\CRMDeveloperExtensions.config");

                if (string.IsNullOrEmpty(selectedConnection.ConnectionString)) return Guid.Empty;
                if (string.IsNullOrEmpty(selectedConnection.OrgId)) return Guid.Empty;

                var props = _dte.Properties["CRM Developer Extensions", "Report Deployer"];
                bool allowPublish = (bool)props.Item("AllowPublishManagedReports").Value;

                //Get the mapped file info
                XmlNodeList mappedFiles = doc.GetElementsByTagName("File");
                foreach (XmlNode file in mappedFiles)
                {
                    XmlNode orgIdNode = file["OrgId"];
                    if (orgIdNode == null) continue;
                    if (orgIdNode.InnerText.ToUpper() != selectedConnection.OrgId.ToUpper()) continue;

                    XmlNode pathNode = file["Path"];
                    if (pathNode == null) continue;
                    if (pathNode.InnerText.ToUpper() != boundName.ToUpper()) continue;

                    XmlNode isManagedNode = file["IsManaged"];
                    if (isManagedNode == null) continue;

                    bool isManaged;
                    bool isBool = Boolean.TryParse(isManagedNode.InnerText, out isManaged);
                    if (!isBool) continue;
                    if (isManaged && !allowPublish) return Guid.Empty;

                    XmlNode reportIdNode = file["ReportId"];
                    if (reportIdNode == null) return Guid.Empty;

                    Guid reportId;
                    bool isGuid = Guid.TryParse(reportIdNode.InnerText, out reportId);
                    if (!isGuid) return Guid.Empty;

                    return reportId;
                }

                return Guid.Empty;
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Getting Mapping: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
                return Guid.Empty;
            }
        }
    }
}
