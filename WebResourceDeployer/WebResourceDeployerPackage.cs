using EnvDTE;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using OutputLogger;
using System;
using System.ComponentModel.Design;
using System.IO;
using System.Runtime.InteropServices;
using System.ServiceModel;
using System.Text;
using System.Windows;
using System.Xml;
using WebResourceDeployer.Models;
using Window = EnvDTE.Window;

namespace WebResourceDeployer
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideToolWindow(typeof(WrdWindow))]
    [Guid(GuidList.GuidWebResourceDeployerPkgString)]
    [ProvideAutoLoad("f1536ef8-92ec-443c-9ed7-fdadf150da82")]
    public sealed class WebResourceDeployerPackage : Package
    {
        private DTE _dte;
        private Logger _logger;

        protected override void Initialize()
        {
            base.Initialize();

            _logger = new Logger();

            _dte = GetGlobalService(typeof(DTE)) as DTE;
            if (_dte == null)
                return;

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (mcs != null)
            {
                // Create the command for the tool window
                CommandID windowCommandId = new CommandID(GuidList.GuidWebResourceDeployerCmdSet, (int)PkgCmdIdList.CmdidWebResourceDeployerWindow);
                OleMenuCommand windowItem = new OleMenuCommand(ShowToolWindow, windowCommandId);
                mcs.AddCommand(windowItem);

                // Create the command for the menu item.
                CommandID publishCommandId = new CommandID(GuidList.GuidItemMenuCommandsCmdSet, (int)PkgCmdIdList.CmdidWebResourceDeployerPublish);
                OleMenuCommand publishMenuItem = new OleMenuCommand(PublishItemCallback, publishCommandId);
                publishMenuItem.BeforeQueryStatus += PublishItem_BeforeQueryStatus;
                publishMenuItem.Visible = false;
                mcs.AddCommand(publishMenuItem);

                // Create the command for the editor menu item.
                CommandID editorPublishCommandId = new CommandID(GuidList.GuidEditorCommandsCmdSet, (int)PkgCmdIdList.CmdidWebResourceEditorPublish);
                OleMenuCommand editorPublishMenuItem = new OleMenuCommand(PublishItemCallback, editorPublishCommandId);
                editorPublishMenuItem.BeforeQueryStatus += PublishItem_BeforeQueryStatus;
                editorPublishMenuItem.Visible = false;
                mcs.AddCommand(editorPublishMenuItem);
            }
        }

        private void ShowToolWindow(object sender, EventArgs e)
        {
            ToolWindowPane window = FindToolWindow(typeof(WrdWindow), 0, true);
            if ((null == window) || (null == window.Frame))
            {
                throw new NotSupportedException(Resources.CanNotCreateWindow);
            }
            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            ErrorHandler.ThrowOnFailure(windowFrame.Show());
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

            CrmConn selectedConnection = GetSelectedConnection(projectItem);
            if (selectedConnection == null)
            {
                menuCommand.Visible = false;
                return;
            }

            Guid webResourceId = GetMapping(projectItem, selectedConnection);
            menuCommand.Visible = webResourceId != Guid.Empty;
        }

        private async void PublishItemCallback(object sender, EventArgs e)
        {
            if (_dte.SelectedItems.Count != 1) return;

            SelectedItem item = _dte.SelectedItems.Item(1);
            ProjectItem projectItem = item.ProjectItem;

            if (projectItem.IsDirty)
            {
                MessageBoxResult result = MessageBox.Show("Save items and publish?", "Unsaved Items",
                    MessageBoxButton.YesNo);

                if (result != MessageBoxResult.Yes) return;

                projectItem.Save();
            }

            CrmConn selectedConnection = GetSelectedConnection(projectItem);
            if (selectedConnection == null) return;

            Guid webResourceId = GetMapping(projectItem, selectedConnection);
            if (webResourceId == Guid.Empty) return;

            CrmConnection connection = CrmConnection.Parse(selectedConnection.ConnectionString);

            //Check if < CRM 2011 UR12 (ExecuteMutliple)
            Version version = Version.Parse(selectedConnection.Version);
            if (version.Major == 5 && version.Revision < 3200)
                await System.Threading.Tasks.Task.Run(() => UpdateAndPublishSingle(connection, projectItem, webResourceId));
            else
                await System.Threading.Tasks.Task.Run(() => UpdateAndPublishMultiple(connection, projectItem, webResourceId));
        }

        private void UpdateAndPublishMultiple(CrmConnection connection, ProjectItem projectItem, Guid webResourceId)
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

                string content = File.ReadAllText(projectItem.FileNames[1]);
                webResource["content"] = EncodeString(content);

                UpdateRequest request = new UpdateRequest { Target = webResource };
                requests.Add(request);

                publishXml += "<webresource>{" + webResource.Id + "}</webresource>";
                publishXml += "</webresources></importexportxml>";

                PublishXmlRequest pubRequest = new PublishXmlRequest { ParameterXml = publishXml };
                requests.Add(pubRequest);
                emRequest.Requests = requests;

                using (OrganizationService orgService = new OrganizationService(connection))
                {
                    _dte.StatusBar.Text = "Updating & publishing web resource...";
                    ExecuteMultipleResponse emResponse = (ExecuteMultipleResponse)orgService.Execute(emRequest);

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
        }

        private void UpdateAndPublishSingle(CrmConnection connection, ProjectItem projectItem, Guid webResourceId)
        {
            try
            {
                _dte.StatusBar.Text = "Updating & publishing web resource...";
                _dte.StatusBar.Highlight(true);

                using (OrganizationService orgService = new OrganizationService(connection))
                {
                    string publishXml = "<importexportxml><webresources>";
                    Entity webResource = new Entity("webresource") { Id = webResourceId };

                    string content = File.ReadAllText(projectItem.FileNames[1]);
                    webResource["content"] = EncodeString(content);

                    UpdateRequest request = new UpdateRequest { Target = webResource };
                    orgService.Execute(request);
                    _logger.WriteToOutputWindow("Uploaded Web Resource", Logger.MessageType.Info);

                    publishXml += "<webresource>{" + webResource.Id + "}</webresource>";
                    publishXml += "</webresources></importexportxml>";

                    PublishXmlRequest pubRequest = new PublishXmlRequest { ParameterXml = publishXml };

                    orgService.Execute(pubRequest);
                    _logger.WriteToOutputWindow("Published Web Resource", Logger.MessageType.Info);
                }
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                _logger.WriteToOutputWindow("Error Updating And Publishing Web Resource To CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, Logger.MessageType.Error);
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Updating And Publishing Web Resource To CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
            }

            _dte.StatusBar.Text = "Published";
        }

        private CrmConn GetSelectedConnection(ProjectItem projectItem)
        {
            CrmConn selectedConnection = new CrmConn();
            Project project = projectItem.ContainingProject;
            var projectPath = Path.GetDirectoryName(project.FullName);
            if (projectPath == null) return selectedConnection;

            var path = Path.GetDirectoryName(project.FullName);
            XmlDocument doc = new XmlDocument();
            doc.Load(path + "\\CRMDeveloperExtensions.config");

            XmlNodeList connections = doc.GetElementsByTagName("Connection");
            if (connections.Count == 0) return selectedConnection;

            //Get the selected Connection info
            foreach (XmlNode node in connections)
            {
                XmlNode selectedNode = node["Selected"];
                if (selectedNode == null) continue;

                bool selected;
                bool isBool = Boolean.TryParse(selectedNode.InnerText, out selected);
                if (!isBool) continue;
                if (!selected) continue;

                XmlNode connectionStringNode = node["ConnectionString"];
                if (connectionStringNode == null) continue;

                selectedConnection.ConnectionString = DecodeString(connectionStringNode.InnerText);

                XmlNode orgIdNode = node["OrgId"];
                if (orgIdNode == null) continue;

                selectedConnection.OrgId = orgIdNode.InnerText;

                XmlNode vesionNode = node["Version"];
                if (vesionNode == null) continue;

                selectedConnection.Version = vesionNode.InnerText;

                break;
            }

            return selectedConnection;
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
                if (!ConfigFileExists(project))
                {
                    _logger.WriteToOutputWindow("Error Getting Mapping: Missing CRMDeveloperExtensions.config File", Logger.MessageType.Error);
                    return Guid.Empty;
                }

                XmlDocument doc = new XmlDocument();
                doc.Load(path + "\\CRMDeveloperExtensions.config");

                if (string.IsNullOrEmpty(selectedConnection.ConnectionString)) return Guid.Empty;
                if (string.IsNullOrEmpty(selectedConnection.OrgId)) return Guid.Empty;

                var props = _dte.Properties["CRM Developer Extensions", "Settings"];
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

        private string DecodeString(string value)
        {
            byte[] data = Convert.FromBase64String(value);
            return Encoding.UTF8.GetString(data);
        }

        private string EncodeString(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
        }

        private bool ConfigFileExists(Project project)
        {
            var path = Path.GetDirectoryName(project.FullName);
            return File.Exists(path + "/CRMDeveloperExtensions.config");
        }
    }
}
