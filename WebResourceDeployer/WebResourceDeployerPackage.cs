using CommonResources;
using CommonResources.Models;
using EnvDTE;
using Microsoft.Crm.Sdk.Messages;
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
using System.Text;
using System.Windows;
using System.Xml;
using Window = EnvDTE.Window;

namespace WebResourceDeployer
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.GuidWebResourceDeployerPkgString)]
    [ProvideAutoLoad("f1536ef8-92ec-443c-9ed7-fdadf150da82")]
    public sealed class WebResourceDeployerPackage : Package
    {
        private DTE _dte;
        private Logger _logger;
        private const string WindowType = "WebResourceDeployer";

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
                CommandID publishCommandId = new CommandID(GuidList.GuidItemMenuCommandsCmdSet, (int)PkgCmdIdList.CmdidWebResourceDeployerPublish);
                OleMenuCommand publishMenuItem = new OleMenuCommand(PublishItemCallback, publishCommandId);
                publishMenuItem.BeforeQueryStatus += PublishItem_BeforeQueryStatus;
                publishMenuItem.Visible = false;
                mcs.AddCommand(publishMenuItem);

                CommandID editorPublishCommandId = new CommandID(GuidList.GuidEditorCommandsCmdSet, (int)PkgCmdIdList.CmdidWebResourceEditorPublish);
                OleMenuCommand editorPublishMenuItem = new OleMenuCommand(PublishItemCallback, editorPublishCommandId);
                editorPublishMenuItem.BeforeQueryStatus += PublishItem_BeforeQueryStatus;
                editorPublishMenuItem.Visible = false;
                mcs.AddCommand(editorPublishMenuItem);
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

            Guid webResourceId = GetMapping(projectItem, selectedConnection);
            menuCommand.Visible = webResourceId != Guid.Empty;
        }

        private void PublishItemCallback(object sender, EventArgs e)
        {
            if (_dte.SelectedItems.Count != 1) return;

            SelectedItem item = _dte.SelectedItems.Item(1);
            ProjectItem projectItem = item.ProjectItem;

            if (projectItem.IsDirty)
            {
                MessageBoxResult result = MessageBox.Show("Save item and publish?", "Unsaved Item",
                    MessageBoxButton.YesNo);

                if (result != MessageBoxResult.Yes) return;

                projectItem.Save();
            }

            //Build TypeScript project
            if (projectItem.Name.ToUpper().EndsWith("TS"))
            {
                SolutionBuild solutionBuild = _dte.Solution.SolutionBuild;
                solutionBuild.BuildProject(_dte.Solution.SolutionBuild.ActiveConfiguration.Name, projectItem.ContainingProject.UniqueName, true);
            }

            CrmConn selectedConnection = (CrmConn)SharedGlobals.GetGlobal("SelectedConnection", _dte);
            if (selectedConnection == null) return;

            Guid webResourceId = GetMapping(projectItem, selectedConnection);
            if (webResourceId == Guid.Empty) return;

            CrmServiceClient client = SharedConnection.GetCurrentConnection(selectedConnection.ConnectionString, WindowType, _dte);

            //Check if < CRM 2011 UR12 (ExecuteMutliple)
            Version version = Version.Parse(selectedConnection.Version);
            if (version.Major == 5 && version.Revision < 3200)
                UpdateAndPublishSingle(client, projectItem, webResourceId);
            else
                UpdateAndPublishMultiple(client, projectItem, webResourceId);
        }

        private void UpdateAndPublishMultiple(CrmServiceClient client, ProjectItem projectItem, Guid webResourceId)
        {
            try
            {
                ExecuteMultipleRequest emRequest = new ExecuteMultipleRequest
                {
                    Requests = new OrganizationRequestCollection(),
                    Settings = new ExecuteMultipleSettings
                    {
                        ContinueOnError = false,
                        ReturnResponses = true
                    }
                };

                OrganizationRequestCollection requests = new OrganizationRequestCollection();

                string publishXml = "<importexportxml><webresources>";
                Entity webResource = new Entity("webresource") { Id = webResourceId };

                string extension = Path.GetExtension(projectItem.FileNames[1]);
                string content = extension != null && (extension.ToUpper() != ".TS")
                    ? File.ReadAllText(projectItem.FileNames[1])
                    : File.ReadAllText(Path.ChangeExtension(projectItem.FileNames[1], ".js"));
                webResource["content"] = EncodeString(content);

                UpdateRequest request = new UpdateRequest { Target = webResource };
                requests.Add(request);

                publishXml += "<webresource>{" + webResource.Id + "}</webresource>";
                publishXml += "</webresources></importexportxml>";

                PublishXmlRequest pubRequest = new PublishXmlRequest { ParameterXml = publishXml };
                requests.Add(pubRequest);
                emRequest.Requests = requests;

                _dte.StatusBar.Text = "Updating & publishing web resource...";
                _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationDeploy);

                ExecuteMultipleResponse emResponse = (ExecuteMultipleResponse)client.Execute(emRequest);

                bool wasError = false;
                foreach (var responseItem in emResponse.Responses)
                {
                    if (responseItem.Fault == null) continue;

                    _logger.WriteToOutputWindow(
                        "Error Updating And Publishing Web Resources To CRM: " + responseItem.Fault.Message + Environment.NewLine + responseItem.Fault.TraceText,
                        Logger.MessageType.Error);
                    wasError = true;
                }

                if (wasError)
                    MessageBox.Show("Error Updating And Publishing Web Resources To CRM. See the Output Window for additional details.");
                else
                    _logger.WriteToOutputWindow("Updated And Published Web Resource", Logger.MessageType.Info);
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                _logger.WriteToOutputWindow("Error Updating And Publishing Web Resource To CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, Logger.MessageType.Error);
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Updating And Publishing Web Resource To CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
            }

            _dte.StatusBar.Clear();
            _dte.StatusBar.Animate(false, vsStatusAnimation.vsStatusAnimationDeploy);
        }

        private void UpdateAndPublishSingle(CrmServiceClient client, ProjectItem projectItem, Guid webResourceId)
        {
            try
            {
                _dte.StatusBar.Text = "Updating & publishing web resource...";
                _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationDeploy);
                string publishXml = "<importexportxml><webresources>";
                Entity webResource = new Entity("webresource") { Id = webResourceId };

                string extension = Path.GetExtension(projectItem.FileNames[1]);
                string content = extension != null && (extension.ToUpper() != ".TS")
                    ? File.ReadAllText(projectItem.FileNames[1])
                    : File.ReadAllText(Path.ChangeExtension(projectItem.FileNames[1], ".js"));
                webResource["content"] = EncodeString(content);

                UpdateRequest request = new UpdateRequest { Target = webResource };
                client.Execute(request);
                _logger.WriteToOutputWindow("Uploaded Web Resource", Logger.MessageType.Info);

                publishXml += "<webresource>{" + webResource.Id + "}</webresource>";
                publishXml += "</webresources></importexportxml>";

                PublishXmlRequest pubRequest = new PublishXmlRequest { ParameterXml = publishXml };

                client.Execute(pubRequest);
                _logger.WriteToOutputWindow("Published Web Resource", Logger.MessageType.Info);
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                _logger.WriteToOutputWindow("Error Updating And Publishing Web Resource To CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, Logger.MessageType.Error);
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Updating And Publishing Web Resource To CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
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

                var props = _dte.Properties["CRM Developer Extensions", "Web Resource Deployer"];
                bool allowPublish = (bool)props.Item("AllowPublishManagedWebResources").Value;

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

                    XmlNode webResourceIdNode = file["WebResourceId"];
                    if (webResourceIdNode == null) return Guid.Empty;

                    Guid webResourceId;
                    bool isGuid = Guid.TryParse(webResourceIdNode.InnerText, out webResourceId);
                    if (!isGuid) return Guid.Empty;

                    return webResourceId;
                }

                return Guid.Empty;
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Getting Mapping: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
                return Guid.Empty;
            }
        }

        private string EncodeString(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
        }
    }
}