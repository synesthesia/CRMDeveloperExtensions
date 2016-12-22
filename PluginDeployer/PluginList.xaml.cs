using CommonResources;
using EnvDTE;
using InfoWindow;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using NuGet.VisualStudio;
using OutputLogger;
using PluginDeployer.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Xml;
using VSLangProj;
using Path = System.IO.Path;
using Task = System.Threading.Tasks.Task;
using Window = EnvDTE.Window;

namespace PluginDeployer
{
    public partial class PluginList
    {
        private readonly DTE _dte;
        private readonly Solution _solution;
        private readonly Logger _logger;
        private bool _isIlMergeInstalled;
        private const string SolutionFolder = "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}";
        private const string WindowType = "PluginDeployer";

        public PluginList()
        {
            InitializeComponent();

            _logger = new Logger();

            _dte = Package.GetGlobalService(typeof(DTE)) as DTE;
            if (_dte == null)
                return;

            _solution = _dte.Solution;
            if (_solution == null)
                return;

            var events = _dte.Events;
            var windowEvents = events.WindowEvents;
            windowEvents.WindowActivated += WindowEventsOnWindowActivated;
            var solutionEvents = events.SolutionEvents;
            solutionEvents.BeforeClosing += SolutionBeforeClosing;
            solutionEvents.ProjectRemoved += SolutionProjectRemoved;

            SelectedAssemblyItem.PropertyChanged += SelectedAssemblyItem_PropertyChanged;
        }

        private void ConnPane_OnProjectChanged(object sender, EventArgs e)
        {
            Publish.IsEnabled = false;
            Customizations.IsEnabled = false;
            Solutions.IsEnabled = false;
            Assemblies.IsEnabled = false;
            Assemblies.ItemsSource = null;

            _isIlMergeInstalled = IsIlMergeInstalled();
            SetIlMergeTooltip(_isIlMergeInstalled);

            var vsproject = ConnPane.SelectedProject.Object as VSProject;
            if (vsproject == null) return;

            vsproject.Events.ReferencesEvents.ReferenceAdded += ReferencesEvents_ReferenceAdded;
        }

        private void ReferencesEvents_ReferenceAdded(Reference reference)
        {
            if (_isIlMergeInstalled)
                SetReferenceCopyLocal(false);
        }

        private void ConnPane_OnConnectionChanged(object sender, ConnectionSelectedEventArgs e)
        {
            if (e.SelectedConnection != null)
            {
                if (e.ConnectionAdded)
                {
                    Customizations.IsEnabled = true;
                    Solutions.IsEnabled = true;
                }
                else
                {
                    Customizations.IsEnabled = false;
                    Solutions.IsEnabled = false;
                    Assemblies.ItemsSource = null;
                }
            }
            else
            {
                Customizations.IsEnabled = false;
                Solutions.IsEnabled = false;
            }

            Publish.IsEnabled = false;
            Assemblies.IsEnabled = false;
        }

        private void ConnPane_OnConnectionStarted(object sender, EventArgs e)
        {
            _dte.StatusBar.Text = "Connecting to CRM...";
            _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationSync);
        }

        private async void ConnPane_OnConnected(object sender, ConnectEventArgs e)
        {
            bool gotPlugins = await GetPlugins(e.ConnectionString);

            if (!gotPlugins)
            {
                Customizations.IsEnabled = false;
                Solutions.IsEnabled = false;
                return;
            }

            Customizations.IsEnabled = true;
            Solutions.IsEnabled = true;
        }

        private async void ConnPane_OnConnectionAdded(object sender, ConnectionAddedEventArgs e)
        {
            bool gotPlugins = await GetPlugins(e.AddedConnection.ConnectionString);

            if (!gotPlugins)
            {
                Customizations.IsEnabled = false;
                Solutions.IsEnabled = false;
                return;
            }

            Customizations.IsEnabled = true;
            Solutions.IsEnabled = true;
        }

        private async void ConnPane_OnConnectionModified(object sender, ConnectionModifiedEventArgs e)
        {
            bool gotPlugins = await GetPlugins(e.ModifiedConnection.ConnectionString);

            if (!gotPlugins)
            {
                Customizations.IsEnabled = false;
                Solutions.IsEnabled = false;
                return;
            }

            Customizations.IsEnabled = true;
            Solutions.IsEnabled = true;
        }

        private void Info_OnClick(object sender, RoutedEventArgs e)
        {
            Info info = new Info();
            info.ShowDialog();
        }

        private void Publish_Click(object sender, RoutedEventArgs e)
        {
            if (Assemblies.SelectedItem == null) return;

            AssemblyItem assemblyItem = (AssemblyItem)Assemblies.SelectedItem;
            UpdateAssembly(assemblyItem);
        }

        private async void UpdateAssembly(AssemblyItem assemblyItem)
        {
            string projectName = ConnPane.SelectedProject.Name;
            Project project = GetProjectByName(projectName);
            if (project == null) return;

            string connString = ConnPane.SelectedConnection.ConnectionString;
            if (connString == null) return;
            CrmServiceClient client = SharedConnection.GetCurrentConnection(connString, WindowType, _dte);

            LockMessage.Content = "Updating...";
            LockOverlay.Visibility = Visibility.Visible;

            bool success = await Task.Run(() => UpdateCrmAssembly(assemblyItem, client));

            LockOverlay.Visibility = Visibility.Hidden;

            if (success) { return; }

            MessageBox.Show("Error Updating Assembly. See the Output Window for additional details.");
            _dte.StatusBar.Clear();
        }

        private bool UpdateCrmAssembly(AssemblyItem assemblyItem, CrmServiceClient client)
        {
            _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationDeploy);

            try
            {
                string outputFileName = ConnPane.SelectedProject.Properties.Item("OutputFileName").Value.ToString();
                string path = GetOutputPath() + outputFileName;

                //Build the project
                SolutionBuild solutionBuild = _dte.Solution.SolutionBuild;
                solutionBuild.BuildProject(_dte.Solution.SolutionBuild.ActiveConfiguration.Name, ConnPane.SelectedProject.UniqueName, true);

                if (solutionBuild.LastBuildInfo > 0)
                    return false;

                //Make sure Major and Minor versions match
                Version assemblyVersion = Version.Parse(ConnPane.SelectedProject.Properties.Item("AssemblyVersion").Value.ToString());
                if (assemblyItem.Version.Major != assemblyVersion.Major ||
                    assemblyItem.Version.Minor != assemblyVersion.Minor)
                {
                    _logger.WriteToOutputWindow("Error Updating Assembly In CRM: Changes To Major & Minor Versions Require Redeployment", Logger.MessageType.Error);
                    return false;
                }

                //Make sure assembly names match
                string assemblyName = ConnPane.SelectedProject.Properties.Item("AssemblyName").Value.ToString();
                if (assemblyName.ToUpper() != assemblyItem.Name.ToUpper())
                {
                    _logger.WriteToOutputWindow("Error Updating Assembly In CRM: Changes To Assembly Name Require Redeployment", Logger.MessageType.Error);
                    return false;
                }

                //Update CRM

                Entity crmAssembly = new Entity("pluginassembly") { Id = assemblyItem.AssemblyId };
                crmAssembly["content"] = Convert.ToBase64String(File.ReadAllBytes(path));

                client.Update(crmAssembly);

                //Update assembly name and version numbers
                assemblyItem.Version = assemblyVersion;
                assemblyItem.Name = ConnPane.SelectedProject.Properties.Item("AssemblyName").Value.ToString();
                assemblyItem.DisplayName = assemblyItem.Name + " (" + assemblyVersion + ")";
                assemblyItem.DisplayName += (assemblyItem.IsWorkflowActivity) ? " [Workflow]" : " [Plug-in]";

                return true;
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                _logger.WriteToOutputWindow("Error Updating Assembly In CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, Logger.MessageType.Error);
                return false;
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Updating Assembly In CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
                return false;
            }
            finally
            {
                _dte.StatusBar.Clear();
                _dte.StatusBar.Animate(false, vsStatusAnimation.vsStatusAnimationDeploy);
            }
        }

        private string GetOutputPath()
        {
            ConfigurationManager configurationManager = ConnPane.SelectedProject.ConfigurationManager;
            if (configurationManager == null) return null;

            Configuration activeConfiguration = configurationManager.ActiveConfiguration;
            string outputPath = activeConfiguration.Properties.Item("OutputPath").Value.ToString();
            string absoluteOutputPath = String.Empty;
            string projectFolder;

            if (outputPath.StartsWith(Path.DirectorySeparatorChar.ToString() + Path.DirectorySeparatorChar))
            {
                absoluteOutputPath = outputPath;
            }
            else if (outputPath.Length >= 2 && outputPath[0] == Path.VolumeSeparatorChar)
            {
                absoluteOutputPath = outputPath;
            }
            else if (outputPath.IndexOf("..\\", StringComparison.Ordinal) != -1)
            {
                projectFolder = Path.GetDirectoryName(ConnPane.SelectedProject.FullName);

                while (outputPath.StartsWith("..\\"))
                {
                    outputPath = outputPath.Substring(3);
                    projectFolder = Path.GetDirectoryName(projectFolder);
                }

                if (projectFolder != null) absoluteOutputPath = Path.Combine(projectFolder, outputPath);
            }
            else
            {
                projectFolder = Path.GetDirectoryName(ConnPane.SelectedProject.FullName);
                if (projectFolder != null)
                    absoluteOutputPath = Path.Combine(projectFolder, outputPath);
            }

            return absoluteOutputPath;
        }

        private void WindowEventsOnWindowActivated(Window gotFocus, Window lostFocus)
        {
            //No solution loaded
            if (_solution.Count == 0)
            {
                ResetForm();
                return;
            }

            //Lost focus
            if (gotFocus.Caption != PluginDeployer.Resources.ResourceManager.GetString("ToolWindowTitle")) return;

            _isIlMergeInstalled = IsIlMergeInstalled();
            SetIlMergeTooltip(_isIlMergeInstalled);
        }

        private void ResetForm()
        {
            Publish.IsEnabled = false;
            Customizations.IsEnabled = false;
            Solutions.IsEnabled = false;
        }

        private void SolutionProjectRemoved(Project project)
        {
            if (ConnPane.SelectedProject != null)
            {
                // TODO: Make sure project.FullName gets a value set
                if (ConnPane.SelectedProject.FullName == project.FullName)
                {
                    Assemblies.ItemsSource = null;
                    Publish.IsEnabled = false;
                    Customizations.IsEnabled = false;
                    Solutions.IsEnabled = false;
                }
            }
        }

        private void SolutionBeforeClosing()
        {
            //Close the Web Plug-in Deployer window - forces having to reopen for a new solution
            foreach (Window window in _dte.Windows)
            {
                if (window.Caption != PluginDeployer.Resources.ResourceManager.GetString("ToolWindowTitle")) continue;

                ResetForm();
                _logger.DeleteOutputWindow();
                window.Close();
                return;
            }
        }

        private async Task<bool> GetPlugins(string connString)
        {
            _dte.StatusBar.Text = "Connecting to CRM...";
            _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationSync);
            LockMessage.Content = "Working...";
            LockOverlay.Visibility = Visibility.Visible;

            CrmServiceClient client = SharedConnection.GetCurrentConnection(ConnPane.SelectedConnection.ConnectionString, WindowType, _dte);

            _dte.StatusBar.Text = "Getting assemblies...";

            EntityCollection results = await Task.Run(() => RetrieveAssembliesFromCrm(client));
            if (results == null)
            {
                SharedConnection.ClearCurrentConnection(WindowType, _dte);
                _dte.StatusBar.Clear();
                _dte.StatusBar.Animate(false, vsStatusAnimation.vsStatusAnimationSync);
                LockOverlay.Visibility = Visibility.Hidden;
                MessageBox.Show("Error Retrieving Assemblies. See the Output Window for additional details.");
                return false;
            }

            _logger.WriteToOutputWindow("Retrieved Assemblies From CRM", Logger.MessageType.Info);

            ObservableCollection<AssemblyItem> assemblies = new ObservableCollection<AssemblyItem>();

            AssemblyItem emptyItem = new AssemblyItem
            {
                AssemblyId = Guid.Empty,
                Name = String.Empty
            };
            assemblies.Add(emptyItem);

            foreach (var entity in results.Entities)
            {
                AssemblyItem aItem = new AssemblyItem
                {
                    AssemblyId = entity.Id,
                    Name = entity.GetAttributeValue<string>("name"),
                    Version = Version.Parse(entity.GetAttributeValue<string>("version")),
                    DisplayName = entity.GetAttributeValue<string>("name") + " (" + entity.GetAttributeValue<string>("version") + ")",
                    IsWorkflowActivity = false
                };

                //Only need to process the 1st assembly/type combination returned 
                if (assemblies.Count(a => a.AssemblyId == aItem.AssemblyId) > 0)
                    continue;

                if (entity.Contains("plugintype.isworkflowactivity"))
                    aItem.IsWorkflowActivity = (bool)entity.GetAttributeValue<AliasedValue>("plugintype.isworkflowactivity").Value;

                aItem.DisplayName += (aItem.IsWorkflowActivity) ? " [Workflow]" : " [Plug-in]";

                assemblies.Add(aItem);
            }

            assemblies = HandleMappings(assemblies);
            Assemblies.ItemsSource = assemblies;
            if (assemblies.Count(a => !string.IsNullOrEmpty(a.BoundProject)) > 0)
                Assemblies.SelectedItem = assemblies.First(a => !string.IsNullOrEmpty(a.BoundProject));
            Assemblies.IsEnabled = true;

            _dte.StatusBar.Clear();
            _dte.StatusBar.Animate(false, vsStatusAnimation.vsStatusAnimationSync);
            LockOverlay.Visibility = Visibility.Hidden;

            return true;
        }

        private EntityCollection RetrieveAssembliesFromCrm(CrmServiceClient client)
        {
            try
            {
                QueryExpression query = new QueryExpression
                {
                    EntityName = "pluginassembly",
                    ColumnSet = new ColumnSet("name", "version"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
						{
							new ConditionExpression
							{
								AttributeName = "ismanaged",
								Operator = ConditionOperator.Equal,
								Values = { false }
							}
						}
                    },
                    LinkEntities =
					    {
						    new LinkEntity
						    {
							    Columns = new ColumnSet("isworkflowactivity"),
							    LinkFromEntityName = "pluginassembly",
							    LinkFromAttributeName = "pluginassemblyid",
							    LinkToEntityName = "plugintype",
							    LinkToAttributeName = "pluginassemblyid",
							    EntityAlias = "plugintype",
							    JoinOperator = JoinOperator.LeftOuter
						    }
					    },
                    Orders =
					    {
						    new OrderExpression
						    {
							    AttributeName = "name",
							    OrderType = OrderType.Ascending
						    }
					    }
                };

                return client.RetrieveMultiple(query);
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                _logger.WriteToOutputWindow("Error Retrieving Assemblies From CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, Logger.MessageType.Error);
                return null;
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Retrieving Assemblies From CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
                return null;
            }
        }

        private ObservableCollection<AssemblyItem> HandleMappings(ObservableCollection<AssemblyItem> aItems)
        {
            try
            {
                string projectName = ConnPane.SelectedProject.Name;
                Project project = GetProjectByName(projectName);
                if (project == null)
                    return new ObservableCollection<AssemblyItem>();

                var path = Path.GetDirectoryName(project.FullName);
                if (!SharedConfigFile.ConfigFileExists(project))
                {
                    _logger.WriteToOutputWindow("Error Updating Mapping In Config File: Missing CRMDeveloperExtensions.config File", Logger.MessageType.Error);
                    return new ObservableCollection<AssemblyItem>();
                }

                XmlDocument doc = new XmlDocument();
                doc.Load(path + "\\CRMDeveloperExtensions.config");

                List<string> projectNames = new List<string>();
                foreach (Project p in ConnPane.Projects)
                {
                    projectNames.Add(p.Name.ToUpper());
                }

                XmlNodeList assemblyNodes = doc.GetElementsByTagName("Assembly");
                List<XmlNode> nodesToRemove = new List<XmlNode>();

                foreach (AssemblyItem aItem in aItems)
                {
                    foreach (XmlNode assemblyNode in assemblyNodes)
                    {
                        XmlNode orgIdNode = assemblyNode["OrgId"];
                        if (orgIdNode == null) continue;
                        if (orgIdNode.InnerText.ToUpper() != ConnPane.SelectedConnection.OrgId.ToUpper()) continue;

                        XmlNode assemblyId = assemblyNode["AssemblyId"];
                        if (assemblyId == null) continue;
                        if (assemblyId.InnerText.ToUpper() != aItem.AssemblyId.ToString().ToUpper()) continue;

                        XmlNode projectNameNode = assemblyNode["ProjectName"];
                        if (projectNameNode == null) continue;

                        if (!projectNames.Contains(projectNameNode.InnerText.ToUpper()))
                            //Remove mappings for projects that might have been deleted from the solution
                            nodesToRemove.Add(projectNameNode);
                        else
                            aItem.BoundProject = projectNameNode.InnerText;
                    }
                }

                //Remove mappings for assemblies that might have been deleted from CRM
                assemblyNodes = doc.GetElementsByTagName("Assembly");
                foreach (XmlNode assembly in assemblyNodes)
                {
                    XmlNode orgIdNode = assembly["OrgId"];
                    if (orgIdNode == null) continue;
                    if (orgIdNode.InnerText.ToUpper() != ConnPane.SelectedConnection.OrgId.ToUpper()) continue;

                    XmlNode assemblyId = assembly["AssemblyId"];
                    if (assemblyId == null) continue;

                    var count = aItems.Count(a => a.AssemblyId.ToString().ToUpper() == assemblyId.InnerText.ToUpper());
                    if (count == 0)
                        nodesToRemove.Add(assembly);
                }

                //Remove the invalid mappings
                if (nodesToRemove.Count <= 0)
                    return aItems;

                XmlNode projects = nodesToRemove[0].ParentNode;
                foreach (XmlNode xmlNode in nodesToRemove)
                {
                    if (projects != null && projects.ParentNode != null)
                        projects.RemoveChild(xmlNode);
                }

                if (SharedConfigFile.IsConfigReadOnly(path + "\\CRMDeveloperExtensions.config"))
                {
                    FileInfo file = new FileInfo(path + "\\CRMDeveloperExtensions.config") { IsReadOnly = false };
                }

                doc.Save(path + "\\CRMDeveloperExtensions.config");
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Updating Mappings In Config File: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
            }

            return aItems;
        }

        private Project GetProjectByName(string projectName)
        {
            foreach (Project project in ConnPane.Projects)
            {
                if (project.Kind.ToUpper() == SolutionFolder)
                {
                    List<Project> list = new List<Project>();
                    list.AddRange(GetFolderProjects(project));
                    foreach (Project subProject in list)
                    {
                        if (subProject.Name != projectName)
                            return subProject;
                    }
                }
                else
                    if (project.Name != projectName) continue;

                return project;
            }

            return null;
        }

        private static IEnumerable<Project> GetFolderProjects(Project folder)
        {
            var list = new List<Project>();

            foreach (ProjectItem item in folder.ProjectItems)
            {
                var subProject = item.SubProject;

                if (subProject == null) continue;

                if (subProject.Kind.ToUpper() == SolutionFolder)
                    list.AddRange(GetFolderProjects(subProject));
                else
                    list.Add(subProject);
            }

            return list;
        }

        private void Assemblies_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ConnPane.SelectedProject == null || string.IsNullOrEmpty(ConnPane.SelectedProject.Name)) return;

            AssemblyItem item = (AssemblyItem)Assemblies.SelectedItem;
            if (item == null)
            {
                Publish.IsEnabled = false;
                SelectedAssemblyItem.Item = null;
                return;
            }

            item.BoundProject = ConnPane.SelectedProject.Name;
            AddOrUpdateMapping(item);

            Publish.IsEnabled = item.AssemblyId != Guid.Empty;
            SelectedAssemblyItem.Item = item;
        }

        private void SelectedAssemblyItem_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (SelectedAssemblyItem.Item == null) return;

            if (SelectedAssemblyItem.Item.AssemblyId != Guid.Empty)
            {
                ((AssemblyItem)Assemblies.SelectedItem).DisplayName = SelectedAssemblyItem.Item.Name + " (" + SelectedAssemblyItem.Item.Version + ")" +
                    ((SelectedAssemblyItem.Item.IsWorkflowActivity) ? " [Workflow]" : " [Plug-in]");
            }
        }

        private void AddOrUpdateMapping(AssemblyItem item)
        {
            try
            {
                var path = Path.GetDirectoryName(ConnPane.SelectedProject.FullName);
                if (!SharedConfigFile.ConfigFileExists(ConnPane.SelectedProject))
                {
                    _logger.WriteToOutputWindow("Error Updating Mappings In Config File: Missing CRMDeveloperExtensions.config File", Logger.MessageType.Error);
                    return;
                }

                XmlDocument doc = new XmlDocument();
                doc.Load(path + "\\CRMDeveloperExtensions.config");

                //Update or delete existing mapping
                XmlNodeList assemblyNodes = doc.GetElementsByTagName("Assembly");
                if (assemblyNodes.Count > 0)
                {
                    foreach (XmlNode node in assemblyNodes)
                    {
                        bool changed = false;
                        XmlNode orgId = node["OrgId"];
                        if (orgId != null && orgId.InnerText.ToUpper() != ConnPane.SelectedConnection.OrgId.ToUpper()) continue;

                        XmlNode projectNameNode = node["ProjectName"];
                        if (projectNameNode != null && projectNameNode.InnerText.ToUpper() != item.BoundProject.ToUpper())
                            continue;

                        if (string.IsNullOrEmpty(item.BoundProject) || item.AssemblyId == Guid.Empty)
                        {
                            //Delete
                            var parentNode = node.ParentNode;
                            if (parentNode != null)
                            {
                                parentNode.RemoveChild(node);
                                changed = true;
                            }
                        }
                        else
                        {
                            //Update
                            XmlNode assemblyIdNode = node["AssemblyId"];
                            if (assemblyIdNode != null)
                            {
                                string oldAssemblyId = assemblyIdNode.InnerText;
                                if (oldAssemblyId != item.AssemblyId.ToString())
                                {
                                    assemblyIdNode.InnerText = item.AssemblyId.ToString();
                                    changed = true;
                                }
                            }
                        }

                        if (!changed) return;

                        if (SharedConfigFile.IsConfigReadOnly(path + "\\CRMDeveloperExtensions.config"))
                        {
                            FileInfo file = new FileInfo(path + "\\CRMDeveloperExtensions.config") { IsReadOnly = false };
                        }

                        doc.Save(path + "\\CRMDeveloperExtensions.config");
                        return;
                    }
                }

                //Create new mapping
                XmlNodeList projects = doc.GetElementsByTagName("Assemblies");
                if (projects.Count <= 0)
                    return;

                XmlNode assembly = doc.CreateElement("Assembly");
                XmlNode org = doc.CreateElement("OrgId");
                org.InnerText = ConnPane.SelectedConnection.OrgId;
                assembly.AppendChild(org);
                XmlNode projectNameNode2 = doc.CreateElement("ProjectName");
                projectNameNode2.InnerText = item.BoundProject;
                assembly.AppendChild(projectNameNode2);
                XmlNode assemblyId = doc.CreateElement("AssemblyId");
                assemblyId.InnerText = item.AssemblyId.ToString();
                assembly.AppendChild(assemblyId);
                projects[0].AppendChild(assembly);

                if (SharedConfigFile.IsConfigReadOnly(path + "\\CRMDeveloperExtensions.config"))
                {
                    FileInfo file = new FileInfo(path + "\\CRMDeveloperExtensions.config") { IsReadOnly = false };
                }

                doc.Save(path + "\\CRMDeveloperExtensions.config");
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Updating Mappings In Config File: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
            }
        }

        private void Customizations_OnClick(object sender, RoutedEventArgs e)
        {
            SharedWindow.OpenCrmPage("tools/solution/edit.aspx?id=%7bfd140aaf-4df4-11dd-bd17-0019b9312238%7d",
                ConnPane.SelectedConnection, _dte);
        }

        private void Solutions_OnClick(object sender, RoutedEventArgs e)
        {
            SharedWindow.OpenCrmPage("tools/Solution/home_solution.aspx?etc=7100&sitemappath=Settings|Customizations|nav_solution",
                ConnPane.SelectedConnection, _dte);
        }

        private void RegistrationTool_OnClick(object sender, RoutedEventArgs e)
        {
            var props = _dte.Properties["CRM Developer Extensions", "Plug-in Deployer"];
            string prtPath = (string)props.Item("RegistrationToolPath").Value;

            if (string.IsNullOrEmpty(prtPath))
            {
                MessageBox.Show("Set Plug-in Registraion Tool path under Tools -> Options -> CRM Developer Extensions");
                return;
            }

            if (!prtPath.EndsWith("\\"))
                prtPath += "\\";

            if (!File.Exists(prtPath + "PluginRegistration.exe"))
            {
                MessageBox.Show("PluginRegistration.exe not found at: " + prtPath);
                return;
            }

            try
            {
                System.Diagnostics.Process.Start(prtPath + "PluginRegistration.exe");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error launching Plug-in Registration Tool: " + Environment.NewLine + Environment.NewLine + ex.Message);
            }
        }

        private void IlMerge_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var vsproject = ConnPane.SelectedProject.Object as VSProject;
                if (vsproject == null) return;

                var componentModel = (IComponentModel)Package.GetGlobalService(typeof(SComponentModel));
                if (componentModel == null) return;

                if (!_isIlMergeInstalled)
                {
                    InstallIlMerge(componentModel);

                    //CRM Assemblies shouldn't be copied local to prevent merging
                    SetReferenceCopyLocal(false);
                }
                else
                {
                    UninstallIlMerge(componentModel);

                    // Reset CRM Assemblies to copy local
                    SetReferenceCopyLocal(true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error installing : " + Environment.NewLine + Environment.NewLine + ex.Message);
            }
        }

        private bool IsIlMergeInstalled()
        {
            if (ConnPane.SelectedProject == null) return false;

            var componentModel = (IComponentModel)Package.GetGlobalService(typeof(SComponentModel));
            if (componentModel == null) return false;

            var installerService = componentModel.GetService<IVsPackageInstallerServices>();
            return installerService.IsPackageInstalled(ConnPane.SelectedProject, "MSBuild.ILMerge.Task");
        }

        private void SetReferenceCopyLocal(bool copyLocal)
        {
            List<string> excludedAssemblies = new List<string>()
                    {
                        "Microsoft.Xrm.Sdk",
                        "Microsoft.Crm.Sdk.Proxy",
                        "Microsoft.Xrm.Sdk.Deployment",
                        "Microsoft.Xrm.Client",
                        "Microsoft.Xrm.Portal",
                        "Microsoft.Xrm.Sdk.Workflow"
                    };

            var vsproject = ConnPane.SelectedProject.Object as VSProject;
            if (vsproject == null) return;

            foreach (Reference reference in vsproject.References)
            {
                if (reference.SourceProject != null) continue;

                if (excludedAssemblies.Contains(reference.Name))
                    reference.CopyLocal = copyLocal;
            }
        }

        private void InstallIlMerge(IComponentModel componentModel)
        {
            try
            {
                var installer = componentModel.GetService<IVsPackageInstaller>();

                _dte.StatusBar.Text = @"Installing MSBuild.ILMerge.Task...";

                installer.InstallPackage("http://packages.nuget.org", ConnPane.SelectedProject, "MSBuild.ILMerge.Task",
                    (Version)null, false);

                _dte.StatusBar.Clear();

                SetIlMergeTooltip(true);
                _isIlMergeInstalled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error installing MSBuild.ILMerge.Task" + Environment.NewLine + Environment.NewLine + ex.Message);
            }
        }

        private void UninstallIlMerge(IComponentModel componentModel)
        {
            try
            {
                var uninstaller = componentModel.GetService<IVsPackageUninstaller>();

                _dte.StatusBar.Text = @"Uninstalling MSBuild.ILMerge.Task...";

                uninstaller.UninstallPackage(ConnPane.SelectedProject, "MSBuild.ILMerge.Task", true);

                _dte.StatusBar.Clear();

                SetIlMergeTooltip(false);
                _isIlMergeInstalled = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error uninstalling MSBuild.ILMerge.Task" + Environment.NewLine + Environment.NewLine + ex.Message);
            }
        }

        private void SetIlMergeTooltip(bool installed)
        {
            IlMerge.ToolTip = installed ? "Remove ILMerge" : "ILMerge Referenced Assemblies";
        }
    }
}