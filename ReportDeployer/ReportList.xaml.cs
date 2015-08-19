using CrmConnectionWindow;
using EnvDTE;
using InfoWindow;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using OutputLogger;
using ReportDeployer.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Xml;
using Button = System.Windows.Controls.Button;
using CheckBox = System.Windows.Controls.CheckBox;
using Window = EnvDTE.Window;

namespace ReportDeployer
{
    public partial class ReportList
    {
        //All DTE objects need to be declared here or else things stop working
        private readonly DTE _dte;
        private readonly Solution _solution;
        private readonly Events _events;
        private readonly SolutionEvents _solutionEvents;
        private Projects _projects;
        private VsSolutionEvents _vsSolutionEvents;
        private readonly IVsSolution _vsSolution;
        private uint _solutionEventsCookie;

        private CrmConn _selectedConn;
        private Project _selectedProject;
        private bool _projectEventsRegistered;
        private bool _connectionAdded;
        private readonly Logger _logger;

        public ReportList()
        {
            InitializeComponent();

            _logger = new Logger();

            _dte = Package.GetGlobalService(typeof(DTE)) as DTE;
            if (_dte == null)
                return;

            _solution = _dte.Solution;
            if (_solution == null)
                return;

            _events = _dte.Events;
            var windowEvents = _events.WindowEvents;
            windowEvents.WindowActivated += WindowEventsOnWindowActivated;
            _solutionEvents = _events.SolutionEvents;
            _solutionEvents.BeforeClosing += BeforeSolutionClosing;
            _solutionEvents.Opened += SolutionOpened;
            _solutionEvents.BeforeClosing += SolutionBeforeClosing;
            _solutionEvents.ProjectAdded += SolutionProjectAdded;
            _solutionEvents.ProjectRemoved += SolutionProjectRemoved;
            _solutionEvents.ProjectRenamed += SolutionProjectRenamed;

            _vsSolutionEvents = new VsSolutionEvents(this);
            _vsSolution = (IVsSolution)ServiceProvider.GlobalProvider.GetService(typeof(SVsSolution));
            _vsSolution.AdviseSolutionEvents(_vsSolutionEvents, out _solutionEventsCookie);
        }

        private void WindowEventsOnWindowActivated(Window gotFocus, Window lostFocus)
        {
            if (_projects == null)
                _projects = _dte.Solution.Projects;

            //No solution loaded
            if (_solution.Count == 0)
            {
                ResetForm();
                return;
            }

            //Lost focus
            if (gotFocus.Caption != ReportDeployer.Resources.ResourceManager.GetString("ToolWindowTitle")) return;

            Projects.IsEnabled = true;
            AddConnection.IsEnabled = true;
            Connections.IsEnabled = true;

            foreach (Project project in _projects)
            {
                SolutionProjectAdded(project);
            }

            if (!_projectEventsRegistered)
            {
                RegisterProjectEvents();
                _projectEventsRegistered = true;
            }
        }

        private void RegisterProjectEvents()
        {
            //Manually register the OnAfterOpenProject event on the existing projects as they are already opened by the time the event would normally be registered
            foreach (Project project in _projects)
            {
                IVsHierarchy projectHierarchy;
                if (_vsSolution.GetProjectOfUniqueName(project.UniqueName, out projectHierarchy) != VSConstants.S_OK)
                    continue;

                IVsSolutionEvents vsSolutionEvents = new VsSolutionEvents(this);
                vsSolutionEvents.OnAfterOpenProject(projectHierarchy, 1);
            }
        }

        private void SolutionBeforeClosing()
        {
            //Close the Report Deployer window - forces having to reopen for a new solution
            foreach (Window window in _dte.Windows)
            {
                if (window.Caption != ReportDeployer.Resources.ResourceManager.GetString("ToolWindowTitle")) continue;

                ResetForm();
                _logger.DeleteOutputWindow();
                window.Close();
                return;
            }
        }

        public void ProjectItemRenamed(ProjectItem projectItem)
        {
            if (!IsItemInSelectedProject(projectItem)) return;

            List<ReportItem> reports = (List<ReportItem>)ReportGrid.ItemsSource;
            if (reports == null) return;

            var fullname = projectItem.FileNames[1];
            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null) return;

            if (projectItem.Kind.ToUpper() == "{F14B399A-7131-4C87-9E4B-1186C45EF12D}") //File
            {
                string[] files = GetRdlFiles(_selectedProject.ProjectItems);
                List<string> shortFiles = new List<string>();
                foreach (string file in files)
                {
                    var fileName = Path.GetFileName(file);
                    if (fileName != null)
                        shortFiles.Add("/" + fileName.ToUpper());
                }

                string oldName = null;
                foreach (ComboBoxItem item in reports[0].ProjectFiles)
                {
                    if (string.IsNullOrEmpty(item.Content.ToString())) continue;

                    if (!shortFiles.Contains(item.Content.ToString().ToUpper()))
                    {
                        oldName = item.Content.ToString();
                        break;
                    }
                }

                var newItemName = fullname.Replace(projectPath, String.Empty).Replace("\\", "/");

                if (projectItem.Name == null) return;

                var oldItemName = newItemName.Replace(Path.GetFileName(projectItem.Name), oldName).Replace("//", "/");

                foreach (var reportItem in reports)
                {
                    string boundFile = reportItem.BoundFile;
                    reportItem.ProjectFiles = GetProjectFiles(_selectedProject.Name);
                    reportItem.BoundFile = boundFile;

                    if (reportItem.BoundFile != oldItemName) continue;

                    reportItem.BoundFile = newItemName;
                }
            }
        }

        public void ProjectItemRemoved(ProjectItem projectItem, uint itemid)
        {
            if (!IsItemInSelectedProject(projectItem)) return;

            List<ReportItem> reports = (List<ReportItem>)ReportGrid.ItemsSource;
            if (reports == null) return;

            var fullname = projectItem.FileNames[1];
            var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
            if (projectPath == null) return;

            if (projectItem.Kind.ToUpper() == "{F14B399A-7131-4C87-9E4B-1186C45EF12D}") //File
            {
                var name = fullname.Replace(projectPath, String.Empty).Replace("\\", "/");

                foreach (ReportItem reportItem in reports)
                {
                    if (!string.IsNullOrEmpty(reportItem.BoundFile) && reportItem.BoundFile == name)
                    {
                        reportItem.BoundFile = null;
                        reportItem.Publish = false;
                    }

                    if (!string.IsNullOrEmpty(reportItem.BoundFile))
                    {
                        var boundFile = reportItem.BoundFile;
                        bool publish = reportItem.Publish;
                        reportItem.ProjectFiles.Clear();
                        reportItem.ProjectFiles = GetProjectFiles(projectItem.ContainingProject.Name);
                        reportItem.BoundFile = boundFile;
                        reportItem.Publish = publish;
                    }

                    SetPublishAll();

                    foreach (ComboBoxItem comboBoxItem in reportItem.ProjectFiles.ToList())
                    {
                        if (comboBoxItem.Content.ToString() == name)
                            reportItem.ProjectFiles.Remove(comboBoxItem);
                    }

                    //If there is only 1 item left it must be the empty item - so remove it
                    if (reportItem.ProjectFiles.ToList().Count == 1)
                        reportItem.ProjectFiles.RemoveAt(0);
                }
            }
        }

        public void ProjectItemAdded(ProjectItem projectItem, uint itemid)
        {
            if (!IsItemInSelectedProject(projectItem)) return;

            List<ReportItem> reports = (List<ReportItem>)ReportGrid.ItemsSource;
            if (reports == null) return;

            if (projectItem.Kind.ToUpper() == "{F14B399A-7131-4C87-9E4B-1186C45EF12D}") //File
            {
                var fullname = projectItem.FileNames[1];
                //Don't want to include files being added here from the temp folder during a diff operation
                if (fullname.ToUpper().Contains(Path.GetTempPath().ToUpper()))
                    return;

                var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
                if (projectPath == null) return;

                foreach (ReportItem reportItem in reports)
                {
                    string boundFile = reportItem.BoundFile;
                    bool publish = reportItem.Publish;
                    reportItem.ProjectFiles = GetProjectFiles(projectItem.ContainingProject.Name);
                    reportItem.BoundFile = boundFile;
                    reportItem.Publish = publish;
                }
            }
        }

        private bool IsItemInSelectedProject(ProjectItem projectItem)
        {
            Project project = projectItem.ContainingProject;
            return _selectedProject == project;
        }

        private void SolutionOpened()
        {
            _projects = _dte.Solution.Projects;
        }

        private void SolutionProjectRenamed(Project project, string oldName)
        {
            string name = Path.GetFileNameWithoutExtension(oldName);
            foreach (ComboBoxItem comboBoxItem in Projects.Items)
            {
                if (string.IsNullOrEmpty(comboBoxItem.Content.ToString())) continue;
                if (name != null && comboBoxItem.Content.ToString().ToUpper() != name.ToUpper()) continue;

                comboBoxItem.Content = project.Name;
            }

            _projects = _dte.Solution.Projects;
        }

        private void SolutionProjectRemoved(Project project)
        {
            foreach (ComboBoxItem comboBoxItem in Projects.Items)
            {
                if (string.IsNullOrEmpty(comboBoxItem.Content.ToString())) continue;
                if (comboBoxItem.Content.ToString().ToUpper() != project.Name.ToUpper()) continue;

                Projects.Items.Remove(comboBoxItem);
                break;
            }

            if (_selectedProject != null)
            {
                if (_selectedProject.FullName == project.FullName)
                {
                    ReportGrid.ItemsSource = null;
                    Connections.ItemsSource = null;
                    Connections.Items.Clear();
                    Connections.IsEnabled = false;
                    AddConnection.IsEnabled = false;
                    Publish.IsEnabled = false;
                    Customizations.IsEnabled = false;
                    Reports.IsEnabled = false;
                    Solutions.IsEnabled = false;
                    AddReport.IsEnabled = false;
                }
            }

            _projects = _dte.Solution.Projects;
        }

        private void SolutionProjectAdded(Project project)
        {
            //Don't want to include the VS Miscellaneous Files Project - which appears occasionally and during a diff operation
            if (project.Name.ToUpper() == "MISCELLANEOUS FILES")
                return;

            bool addProject = true;
            foreach (ComboBoxItem projectItem in Projects.Items)
            {
                if (projectItem.Content.ToString().ToUpper() != project.Name.ToUpper()) continue;

                addProject = false;
                break;
            }

            if (addProject)
            {
                ComboBoxItem item = new ComboBoxItem() { Content = project.Name, Tag = project };
                Projects.Items.Add(item);
            }

            if (Projects.SelectedIndex == -1)
                Projects.SelectedIndex = 0;

            _projects = _dte.Solution.Projects;
        }

        private void BeforeSolutionClosing()
        {
            ResetForm();
        }

        private void ResetForm()
        {
            ReportGrid.ItemsSource = null;
            Connections.ItemsSource = null;
            Connections.Items.Clear();
            Connections.IsEnabled = false;
            Projects.ItemsSource = null;
            Projects.Items.Clear();
            Projects.IsEnabled = false;
            AddConnection.IsEnabled = false;
            Publish.IsEnabled = false;
            Customizations.IsEnabled = false;
            Reports.IsEnabled = false;
            Solutions.IsEnabled = false;
            AddReport.IsEnabled = false;
        }

        private void AddConnection_Click(object sender, RoutedEventArgs e)
        {
            Connection connection = new Connection(null, null);
            bool? result = connection.ShowDialog();

            if (!result.HasValue || !result.Value) return;

            var configExists = ConfigFileExists(_selectedProject);
            if (!configExists)
                CreateConfigFile(_selectedProject);

            Expander.IsExpanded = false;
            Customizations.IsEnabled = true;
            Solutions.IsEnabled = true;
            Reports.IsEnabled = true;

            bool change = AddOrUpdateConnection(_selectedProject, connection.ConnectionName, connection.ConnectionString, connection.OrgId, connection.Version, true);
            if (!change) return;

            GetConnections();
            foreach (CrmConn conn in Connections.Items)
            {
                if (conn.Name != connection.ConnectionName) continue;

                Connections.SelectedItem = conn;
                GetReports(connection.ConnectionString);
                break;
            }
        }

        private bool AddOrUpdateConnection(Project vsProject, string connectionName, string connString, string orgId, string versionNum, bool showPrompt)
        {
            try
            {
                var path = Path.GetDirectoryName(vsProject.FullName);
                if (!ConfigFileExists(vsProject))
                {
                    _logger.WriteToOutputWindow("Error Adding Or Updating Connection: Missing CRMDeveloperExtensions.config File", Logger.MessageType.Error);
                    return false;
                }

                XmlDocument doc = new XmlDocument();
                doc.Load(path + "\\CRMDeveloperExtensions.config");

                //Check if connection alredy exists for project
                XmlNodeList connectionStrings = doc.GetElementsByTagName("ConnectionString");
                if (connectionStrings.Count > 0)
                {
                    foreach (XmlNode node in connectionStrings)
                    {
                        string decodedString = DecodeString(node.InnerText);
                        if (decodedString != connString) continue;

                        if (showPrompt)
                        {
                            MessageBoxResult result = MessageBox.Show("Update Connection?", "Connection Already Added",
                                MessageBoxButton.YesNo);

                            //Update existing connection
                            if (result != MessageBoxResult.Yes)
                                return false;
                        }

                        XmlNode connectionU = node.ParentNode;
                        if (connectionU != null)
                        {
                            XmlNode nameNode = connectionU["Name"];
                            if (nameNode != null)
                                nameNode.InnerText = connectionName;
                            XmlNode versionNode = connectionU["Version"];
                            if (versionNode != null)
                                versionNode.InnerText = versionNum;
                        }

                        doc.Save(path + "\\CRMDeveloperExtensions.config");
                        return true;
                    }
                }

                //Add the connection elements
                XmlNodeList connections = doc.GetElementsByTagName("Connections");
                XmlElement connection = doc.CreateElement("Connection");
                XmlElement name = doc.CreateElement("Name");
                name.InnerText = connectionName;
                connection.AppendChild(name);
                XmlElement org = doc.CreateElement("OrgId");
                org.InnerText = orgId;
                connection.AppendChild(org);
                XmlElement connectionString = doc.CreateElement("ConnectionString");
                connectionString.InnerText = EncodeString(connString);
                connection.AppendChild(connectionString);
                XmlElement version = doc.CreateElement("Version");
                version.InnerText = versionNum;
                connection.AppendChild(version);
                XmlElement selected = doc.CreateElement("Selected");
                selected.InnerText = "True";
                connection.AppendChild(selected);
                connections[0].AppendChild(connection);

                _connectionAdded = true;

                doc.Save(path + "\\CRMDeveloperExtensions.config");
                return true;
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Adding Or Updating Connection: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
                return false;
            }
        }

        private void CreateConfigFile(Project vsProject)
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                XmlElement reportDeployer = doc.CreateElement("ReportDeployer");
                XmlElement connections = doc.CreateElement("Connections");
                XmlElement files = doc.CreateElement("Files");
                reportDeployer.AppendChild(connections);
                reportDeployer.AppendChild(files);
                doc.AppendChild(reportDeployer);

                var path = Path.GetDirectoryName(vsProject.FullName);
                doc.Save(path + "/CRMDeveloperExtensions.config");
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Creating Config File: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
            }
        }

        private bool ConfigFileExists(Project project)
        {
            var path = Path.GetDirectoryName(project.FullName);
            return File.Exists(path + "/CRMDeveloperExtensions.config");
        }

        private Project GetProjectByName(string projectName)
        {
            foreach (Project project in _projects)
            {
                if (project.Name != projectName) continue;

                return project;
            }

            return null;
        }

        private ObservableCollection<ComboBoxItem> GetProjectFiles(string projectName)
        {
            ObservableCollection<ComboBoxItem> projectFiles = new ObservableCollection<ComboBoxItem>();
            Project project = GetProjectByName(projectName);
            if (project == null)
                return projectFiles;

            var projectItems = project.ProjectItems;
            for (int i = 1; i <= projectItems.Count; i++)
            {
                var files = GetFiles(projectItems.Item(i), String.Empty);
                foreach (var comboBoxItem in files)
                {
                    projectFiles.Add(comboBoxItem);
                }
            }

            if (projectFiles.Count > 0)
                projectFiles.Insert(0, new ComboBoxItem() { Content = String.Empty });

            return projectFiles;
        }

        private ObservableCollection<ComboBoxItem> GetFiles(ProjectItem projectItem, string path)
        {
            ObservableCollection<ComboBoxItem> projectFiles = new ObservableCollection<ComboBoxItem>();
            if (projectItem.Kind != "{6BB5F8EF-4483-11D3-8BCF-00C04F8EC28C}") // VS Folder 
            {
                string ex = Path.GetExtension(projectItem.Name);
                if (ex != null && ex.ToUpper() == ".RDL")
                    projectFiles.Add(new ComboBoxItem() { Content = path + "/" + projectItem.Name, Tag = projectItem });
            }
            else
            {
                for (int i = 1; i <= projectItem.ProjectItems.Count; i++)
                {
                    var files = GetFiles(projectItem.ProjectItems.Item(i), path + "/" + projectItem.Name);
                    foreach (var comboBoxItem in files)
                    {
                        projectFiles.Add(comboBoxItem);
                    }
                }
            }
            return projectFiles;
        }

        private void GetConnections()
        {
            Connections.ItemsSource = null;

            var path = Path.GetDirectoryName(_selectedProject.FullName);
            XmlDocument doc = new XmlDocument();

            if (!ConfigFileExists(_selectedProject))
            {
                _logger.WriteToOutputWindow("Error Retrieving Connections From Config File: Missing CRMDeveloperExtensions.config file", Logger.MessageType.Error);
                return;
            }

            doc.Load(path + "\\CRMDeveloperExtensions.config");
            XmlNodeList connections = doc.GetElementsByTagName("Connection");
            if (connections.Count == 0) return;

            List<CrmConn> crmConnections = new List<CrmConn>();

            foreach (XmlNode node in connections)
            {
                CrmConn conn = new CrmConn();
                XmlNode nameNode = node["Name"];
                if (nameNode != null)
                    conn.Name = nameNode.InnerText;
                XmlNode connectionStringNode = node["ConnectionString"];
                if (connectionStringNode != null)
                    conn.ConnectionString = DecodeString(connectionStringNode.InnerText);
                XmlNode orgIdNode = node["OrgId"];
                if (orgIdNode != null)
                    conn.OrgId = orgIdNode.InnerText;
                XmlNode versionNode = node["Version"];
                if (versionNode != null)
                    conn.Version = versionNode.InnerText;

                crmConnections.Add(conn);
            }

            Connections.ItemsSource = crmConnections;

            if (Connections.SelectedIndex == -1 && crmConnections.Count > 0)
                Connections.SelectedIndex = 0;
        }

        private string EncodeString(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
        }

        private string DecodeString(string value)
        {
            byte[] data = Convert.FromBase64String(value);
            return Encoding.UTF8.GetString(data);
        }

        private void Connect_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedConn == null) return;

            string connString = _selectedConn.ConnectionString;
            if (string.IsNullOrEmpty(connString)) return;

            UpdateSelectedConnection(true);
            GetReports(connString);

            Expander.IsExpanded = false;
            Customizations.IsEnabled = true;
            Solutions.IsEnabled = true;
            Reports.IsEnabled = true;
        }

        private async void GetReports(string connString)
        {
            string projectName = _selectedProject.Name;
            CrmConnection connection = CrmConnection.Parse(connString);

            _dte.StatusBar.Text = "Connecting to CRM and getting reports...";
            _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationSync);
            LockMessage.Content = "Working...";
            LockOverlay.Visibility = Visibility.Visible;

            EntityCollection results = await System.Threading.Tasks.Task.Run(() => RetrieveReportsFromCrm(connection));
            if (results == null)
            {
                _dte.StatusBar.Clear();
                LockOverlay.Visibility = Visibility.Hidden;
                MessageBox.Show("Error Retrieving Reports. See the Output Window for additional details.");
                return;
            }

            _logger.WriteToOutputWindow("Retrieved Reports From CRM", Logger.MessageType.Info);

            List<ReportItem> rItems = new List<ReportItem>();
            foreach (var entity in results.Entities)
            {
                ReportItem rItem = new ReportItem
                {
                    Publish = false,
                    ReportId = entity.Id,
                    Name = entity.GetAttributeValue<string>("name"),
                    IsManaged = entity.GetAttributeValue<bool>("ismanaged"),
                    AllowPublish = false,
                    ProjectFiles = GetProjectFiles(projectName),
                };

                rItem.PropertyChanged += ReportItem_PropertyChanged;
                rItems.Add(rItem);
            }

            rItems = HandleMappings(rItems);
            ReportGrid.ItemsSource = rItems;
            FilterReports();
            ReportGrid.IsEnabled = true;
            ShowManaged.IsEnabled = true;
            AddReport.IsEnabled = true;

            _dte.StatusBar.Clear();
            _dte.StatusBar.Animate(false, vsStatusAnimation.vsStatusAnimationSync);
            LockOverlay.Visibility = Visibility.Hidden;
        }

        private EntityCollection RetrieveReportsFromCrm(CrmConnection connection)
        {
            try
            {
                using (OrganizationService orgService = new OrganizationService(connection))
                {
                    QueryExpression query = new QueryExpression
                    {
                        EntityName = "report",
                        ColumnSet = new ColumnSet("name", "ismanaged"),
                        Criteria = new FilterExpression
                        {
                            Conditions =
                            {
                                new ConditionExpression
                                {
                                    AttributeName = "iscustomizable",
                                    Operator = ConditionOperator.Equal,
                                    Values = { true }
                                }
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

                    return orgService.RetrieveMultiple(query);
                }
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                _logger.WriteToOutputWindow("Error Retrieving Reports From CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, Logger.MessageType.Error);
                return null;
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Retrieving Reports From CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
                return null;
            }
        }

        private List<ReportItem> HandleMappings(List<ReportItem> rItems)
        {
            try
            {
                string projectName = _selectedProject.Name;
                Project project = GetProjectByName(projectName);
                if (project == null)
                    return new List<ReportItem>();

                var path = Path.GetDirectoryName(project.FullName);
                if (!ConfigFileExists(project))
                {
                    _logger.WriteToOutputWindow("Error Updating Mappings In Config File: Missing CRMDeveloperExtensions.config File", Logger.MessageType.Error);
                    return rItems;
                }

                XmlDocument doc = new XmlDocument();
                doc.Load(path + "\\CRMDeveloperExtensions.config");

                var props = _dte.Properties["CRM Developer Extensions", "Settings"];
                bool allowPublish = (bool)props.Item("AllowPublishManagedReports").Value;

                XmlNodeList mappedFiles = doc.GetElementsByTagName("File");
                List<XmlNode> nodesToRemove = new List<XmlNode>();

                foreach (ReportItem rItem in rItems)
                {
                    foreach (XmlNode file in mappedFiles)
                    {
                        XmlNode orgIdNode = file["OrgId"];
                        if (orgIdNode == null) continue;
                        if (orgIdNode.InnerText.ToUpper() != _selectedConn.OrgId.ToUpper()) continue;

                        XmlNode reportId = file["ReportId"];
                        if (reportId == null) continue;
                        if (reportId.InnerText.ToUpper() != rItem.ReportId.ToString().ToUpper()) continue;

                        XmlNode filePartialPath = file["Path"];
                        if (filePartialPath == null) continue;

                        string filePath = Path.GetDirectoryName(project.FullName) +
                                          filePartialPath.InnerText.Replace("/", "\\");
                        if (!File.Exists(filePath))
                            //Remove mappings for files that might have been deleted from the project
                            nodesToRemove.Add(file);
                        else
                        {
                            rItem.BoundFile = filePartialPath.InnerText;
                            rItem.AllowPublish = allowPublish || !rItem.IsManaged;
                        }
                    }
                }

                //Remove mappings for files that might have been deleted from CRM
                mappedFiles = doc.GetElementsByTagName("File");
                foreach (XmlNode file in mappedFiles)
                {
                    XmlNode orgIdNode = file["OrgId"];
                    if (orgIdNode == null) continue;
                    if (orgIdNode.InnerText.ToUpper() != _selectedConn.OrgId.ToUpper()) continue;

                    XmlNode reportId = file["ReportId"];
                    if (reportId == null) continue;

                    var count = rItems.Count(w => w.ReportId.ToString().ToUpper() == reportId.InnerText.ToUpper());
                    if (count == 0)
                        nodesToRemove.Add(file);
                }

                //Remove the invalid mappings
                if (nodesToRemove.Count <= 0)
                    return rItems;

                XmlNode files = nodesToRemove[0].ParentNode;
                foreach (XmlNode xmlNode in nodesToRemove)
                {
                    if (files != null && files.ParentNode != null)
                        files.RemoveChild(xmlNode);
                }
                doc.Save(path + "\\CRMDeveloperExtensions.config");
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Updating Mappings In Config File: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
            }

            return rItems;
        }

        private void ReportItem_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "BoundFile")
            {
                ReportItem item = (ReportItem)sender;
                AddOrUpdateMapping(item);
            }

            if (e.PropertyName == "Publish")
            {
                List<ReportItem> reports = (List<ReportItem>)ReportGrid.ItemsSource;
                if (reports == null) return;

                Publish.IsEnabled = reports.Count(w => w.Publish) > 0;

                SetPublishAll();
            }
        }

        private void SetPublishAll()
        {
            List<ReportItem> reports = (List<ReportItem>)ReportGrid.ItemsSource;
            if (reports == null) return;

            //Set Publish All
            CheckBox publishAll = FindVisualChildren<CheckBox>(ReportGrid).FirstOrDefault(t => t.Name == "PublishSelectAll");
            if (publishAll == null) return;

            publishAll.IsChecked = reports.Count(w => w.Publish) == reports.Count(w => w.AllowPublish);
        }

        public static IEnumerable<T> FindVisualChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj == null) yield break;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(depObj, i);
                if (child != null && child is T)
                {
                    yield return (T)child;
                }

                foreach (T childOfChild in FindVisualChildren<T>(child))
                {
                    yield return childOfChild;
                }
            }
        }

        private void AddOrUpdateMapping(ReportItem item)
        {
            try
            {
                var projectPath = Path.GetDirectoryName(_selectedProject.FullName);
                if (!ConfigFileExists(_selectedProject))
                {
                    _logger.WriteToOutputWindow("Error Updating Mappings In Config File: Missing CRMDeveloperExtensions.config File", Logger.MessageType.Error);
                    return;
                }

                XmlDocument doc = new XmlDocument();
                doc.Load(projectPath + "\\CRMDeveloperExtensions.config");

                //Update or delete existing mapping
                XmlNodeList fileNodes = doc.GetElementsByTagName("File");
                if (fileNodes.Count > 0)
                {
                    foreach (XmlNode node in fileNodes)
                    {
                        XmlNode orgId = node["OrgId"];
                        if (orgId != null && orgId.InnerText.ToUpper() != _selectedConn.OrgId.ToUpper()) continue;

                        XmlNode reportId = node["ReportId"];
                        if (reportId != null && reportId.InnerText.ToUpper() !=
                            item.ReportId.ToString()
                                .ToUpper()
                                .Replace("{", String.Empty)
                                .Replace("}", String.Empty))
                            continue;

                        if (string.IsNullOrEmpty(item.BoundFile))
                        {
                            //Delete
                            var parentNode = node.ParentNode;
                            if (parentNode != null)
                            {
                                parentNode.RemoveChild(node);

                                item.Publish = false;
                                item.AllowPublish = false;
                            }
                        }
                        else
                        {
                            //Update
                            XmlNode path = node["Path"];
                            if (path != null)
                                path.InnerText = item.BoundFile;
                        }

                        doc.Save(projectPath + "\\CRMDeveloperExtensions.config");
                        return;
                    }
                }

                //Create new mapping
                XmlNodeList files = doc.GetElementsByTagName("Files");
                if (files.Count > 0)
                {
                    XmlNode file = doc.CreateElement("File");
                    XmlNode org = doc.CreateElement("OrgId");
                    org.InnerText = _selectedConn.OrgId;
                    file.AppendChild(org);
                    XmlNode path = doc.CreateElement("Path");
                    path.InnerText = item.BoundFile;
                    file.AppendChild(path);
                    XmlNode reportId = doc.CreateElement("ReportId");
                    reportId.InnerText = item.ReportId.ToString();
                    file.AppendChild(reportId);
                    XmlNode name = doc.CreateElement("Name");
                    name.InnerText = item.Name;
                    file.AppendChild(name);
                    XmlNode isManaged = doc.CreateElement("IsManaged");
                    isManaged.InnerText = item.IsManaged.ToString();
                    file.AppendChild(isManaged);
                    files[0].AppendChild(file);

                    doc.Save(projectPath + "\\CRMDeveloperExtensions.config");

                    item.AllowPublish = true;
                }
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Updating Mappings In Config File: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
            }
        }

        private void GetReport_OnClick(object sender, RoutedEventArgs e)
        {
            Guid reportId = new Guid(((Button)sender).CommandParameter.ToString());
            ReportItem reportItem =
                ((List<ReportItem>)ReportGrid.ItemsSource)
                    .FirstOrDefault(r => r.ReportId == reportId);

            string folder = String.Empty;
            if (reportItem != null && !string.IsNullOrEmpty(reportItem.BoundFile))
            {
                var directoryName = Path.GetDirectoryName(reportItem.BoundFile);
                if (directoryName != null)
                    folder = directoryName.Replace("\\", "/");
                if (folder == "/")
                    folder = String.Empty;
            }

            string connString = _selectedConn.ConnectionString;
            string projectName = ((ComboBoxItem)Projects.SelectedItem).Content.ToString();
            DownloadReport(reportId, folder, connString, projectName);
        }

        private void DownloadReport(Guid reportId, string folder, string connString, string projectName)
        {
            _dte.StatusBar.Text = "Downloading file...";
            _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationSync);

            try
            {
                CrmConnection connection = CrmConnection.Parse(connString);
                using (OrganizationService orgService = new OrganizationService(connection))
                {
                    Entity report = orgService.Retrieve("report", reportId,
                        new ColumnSet("bodytext", "filename"));

                    _logger.WriteToOutputWindow("Downloaded Report: " + report.Id, Logger.MessageType.Info);

                    Project project = GetProjectByName(projectName);
                    string[] name = report.GetAttributeValue<string>("filename").Split('/');
                    folder = folder.Replace("/", "\\");
                    var path = Path.GetDirectoryName(project.FullName) +
                               ((folder != "\\") ? folder : String.Empty) +
                               "\\" + name[name.Length - 1];

                    if (File.Exists(path))
                    {
                        MessageBoxResult result = MessageBox.Show("OK to overwrite?", "Report Download",
                            MessageBoxButton.YesNo);
                        if (result != MessageBoxResult.Yes)
                        {
                            _dte.StatusBar.Clear();
                            return;
                        }
                    }

                    File.WriteAllText(path, report.GetAttributeValue<string>("bodytext"));

                    ProjectItem projectItem = project.ProjectItems.AddFromFile(path);

                    IVsHierarchy projectHierarchy;
                    if (_vsSolution.GetProjectOfUniqueName(project.UniqueName, out projectHierarchy) == VSConstants.S_OK)
                    {
                        uint itemId;
                        if (projectHierarchy.ParseCanonicalName(path, out itemId) == VSConstants.S_OK)
                            ProjectItemAdded(projectItem, itemId);
                    }

                    var fullname = projectItem.FileNames[1];
                    var projectPath = Path.GetDirectoryName(projectItem.ContainingProject.FullName);
                    if (projectPath == null) return;

                    var boundName = fullname.Replace(projectPath, String.Empty).Replace("\\", "/");

                    List<ReportItem> items = (List<ReportItem>)ReportGrid.ItemsSource;
                    ReportItem item = items.FirstOrDefault(w => w.ReportId == reportId);
                    if (item != null)
                    {
                        item.BoundFile = boundName;

                        CheckBox publishAll =
                            FindVisualChildren<CheckBox>(ReportGrid)
                                .FirstOrDefault(t => t.Name == "PublishSelectAll");
                        if (publishAll == null) return;

                        if (publishAll.IsChecked == true)
                            item.Publish = true;
                    }
                }
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                _logger.WriteToOutputWindow(
                    "Error Downloading Report From CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace,
                    Logger.MessageType.Error);
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow(
                    "Error Downloading Report From CRM: " + ex.Message + Environment.NewLine + ex.StackTrace,
                    Logger.MessageType.Error);
            }
            finally
            {
                _dte.StatusBar.Clear();
                _dte.StatusBar.Animate(false, vsStatusAnimation.vsStatusAnimationSync);
            }
        }

        private void Publish_Click(object sender, RoutedEventArgs e)
        {
            List<ReportItem> items = (List<ReportItem>)ReportGrid.ItemsSource;
            List<ReportItem> selectedItems = items.Where(w => w.Publish).ToList();

            UpdateReports(selectedItems);
        }

        private async void UpdateReports(List<ReportItem> items)
        {
            string projectName = _selectedProject.Name;
            Project project = GetProjectByName(projectName);
            if (project == null) return;

            string connString = _selectedConn.ConnectionString;
            if (connString == null) return;
            CrmConnection connection = CrmConnection.Parse(connString);

            LockMessage.Content = "Deploying...";
            LockOverlay.Visibility = Visibility.Visible;

            bool success;
            //Check if < CRM 2011 UR12 (ExecuteMutliple)
            Version version = Version.Parse(_selectedConn.Version);
            if (version.Major == 5 && version.Revision < 3200)
                success = await System.Threading.Tasks.Task.Run(() => UpdateAndPublishSingle(items, project, connection));
            else
                success = await System.Threading.Tasks.Task.Run(() => UpdateAndPublishMultiple(items, project, connection));

            LockOverlay.Visibility = Visibility.Hidden;

            if (success) return;

            MessageBox.Show("Error Updating Reports. See the Output Window for additional details.");
            _dte.StatusBar.Clear();
        }

        private bool UpdateAndPublishMultiple(List<ReportItem> items, Project project, CrmConnection connection)
        {
            //CRM 2011 UR12+
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

                foreach (var reportItem in items)
                {
                    Entity report = new Entity("report") { Id = reportItem.ReportId };

                    string filePath = Path.GetDirectoryName(project.FullName) +
                                      reportItem.BoundFile.Replace("/", "\\");
                    if (!File.Exists(filePath)) continue;

                    report["bodytext"] = File.ReadAllText(filePath);

                    UpdateRequest request = new UpdateRequest { Target = report };
                    requests.Add(request);
                }

                emRequest.Requests = requests;

                bool wasError = false;
                using (OrganizationService orgService = new OrganizationService(connection))
                {
                    _dte.StatusBar.Text = "Updating report(s)...";
                    _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationDeploy);

                    ExecuteMultipleResponse emResponse = (ExecuteMultipleResponse)orgService.Execute(emRequest);

                    foreach (var responseItem in emResponse.Responses)
                    {
                        if (responseItem.Fault == null) continue;

                        _logger.WriteToOutputWindow(
                            "Error Updating Report(s) To CRM: " + responseItem.Fault.Message +
                            Environment.NewLine + responseItem.Fault.TraceText, Logger.MessageType.Error);
                        wasError = true;
                    }

                    if (wasError)
                    {
                        MessageBox.Show(
                            "Error Updating Report(s) To CRM. See the Output Window for additional details.");
                        _dte.StatusBar.Clear();
                        return false;
                    }
                }

                _logger.WriteToOutputWindow("Updated Report(s)", Logger.MessageType.Info);

                return true;
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                _logger.WriteToOutputWindow("Error Updating Report(s) To CRM: " +
                    crmEx.Message + Environment.NewLine + crmEx.StackTrace, Logger.MessageType.Error);
                return false;
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Updating Report(s) To CRM: " +
                    ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
                return false;
            }
            finally
            {
                _dte.StatusBar.Clear();
                _dte.StatusBar.Animate(false, vsStatusAnimation.vsStatusAnimationDeploy);
            }
        }

        private bool UpdateAndPublishSingle(List<ReportItem> items, Project project, CrmConnection connection)
        {
            //CRM 2011 < UR12
            _dte.StatusBar.Text = "Updating report(s)...";
            _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationDeploy);

            try
            {
                using (OrganizationService orgService = new OrganizationService(connection))
                {
                    foreach (var reportItem in items)
                    {
                        Entity report = new Entity("report") { Id = reportItem.ReportId };

                        string filePath = Path.GetDirectoryName(project.FullName) +
                                          reportItem.BoundFile.Replace("/", "\\");
                        if (!File.Exists(filePath)) continue;

                        report["bodytext"] = File.ReadAllText(filePath);

                        UpdateRequest request = new UpdateRequest { Target = report };
                        orgService.Execute(request);
                        _logger.WriteToOutputWindow("Uploaded Report", Logger.MessageType.Info);
                    }
                }

                return true;
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                _logger.WriteToOutputWindow("Error Updating Report(s) To CRM: " + crmEx.Message + Environment.NewLine +
                    crmEx.StackTrace, Logger.MessageType.Error);
                return false;
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Updating Report(s) To CRM: " + ex.Message + Environment.NewLine +
                    ex.StackTrace, Logger.MessageType.Error);
                return false;
            }
            finally
            {
                _dte.StatusBar.Clear();
                _dte.StatusBar.Animate(false, vsStatusAnimation.vsStatusAnimationDeploy);
            }
        }

        private void Connections_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _selectedConn = (CrmConn)Connections.SelectedItem;
            if (_selectedConn != null)
            {
                Connect.IsEnabled = !string.IsNullOrEmpty(_selectedConn.Name);
                Delete.IsEnabled = !string.IsNullOrEmpty(_selectedConn.Name);
                ModifyConnection.IsEnabled = !string.IsNullOrEmpty(_selectedConn.Name);

                if (_connectionAdded)
                    _connectionAdded = false;
                else
                    UpdateSelectedConnection(false);
            }
            else
            {
                Connect.IsEnabled = false;
                Delete.IsEnabled = false;
                ModifyConnection.IsEnabled = false;
            }

            ReportGrid.ItemsSource = null;
            ShowManaged.IsEnabled = false;
            Publish.IsEnabled = false;
            Customizations.IsEnabled = false;
            Solutions.IsEnabled = false;
            Reports.IsEnabled = false;
            AddReport.IsEnabled = false;
            ReportGrid.IsEnabled = false;
        }

        private void UpdateSelectedConnection(bool makeSelected)
        {
            try
            {
                var path = Path.GetDirectoryName(_selectedProject.FullName);
                if (!ConfigFileExists(_selectedProject))
                {
                    _logger.WriteToOutputWindow("Error Updating Connection: Missing CRMDeveloperExtensions.config File", Logger.MessageType.Error);
                    return;
                }

                XmlDocument doc = new XmlDocument();
                doc.Load(path + "\\CRMDeveloperExtensions.config");

                XmlNodeList connections = doc.GetElementsByTagName("Connection");
                if (connections.Count > 0)
                {
                    foreach (XmlNode node in connections)
                    {
                        XmlNode name = node["Name"];
                        if (name == null) continue;

                        XmlNode selected = node["Selected"];
                        if (selected == null) continue;

                        if (makeSelected)
                            selected.InnerText = name.InnerText != _selectedConn.Name ? "False" : "True";
                        else
                            selected.InnerText = "False";
                    }

                    doc.Save(path + "\\CRMDeveloperExtensions.config");
                }
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Updating Connection: Missing CRMDeveloperExtensions.config File: " + ex.Message, Logger.MessageType.Error);
            }
        }

        private void Projects_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //No solution loaded
            if (_solution.Count == 0) return;

            ReportGrid.ItemsSource = null;
            ShowManaged.IsEnabled = false;
            Publish.IsEnabled = false;
            Customizations.IsEnabled = false;
            Solutions.IsEnabled = false;
            Reports.IsEnabled = false;
            AddReport.IsEnabled = false;
            ReportGrid.IsEnabled = false;

            ComboBoxItem item = (ComboBoxItem)Projects.SelectedItem;
            if (item == null) return;
            if (string.IsNullOrEmpty(item.Content.ToString())) return;

            _selectedProject = (Project)((ComboBoxItem)Projects.SelectedItem).Tag;
            GetConnections();
        }

        private void ModifyConnection_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedConn == null) return;
            if (string.IsNullOrEmpty(_selectedConn.ConnectionString)) return;

            string name = _selectedConn.Name;
            Connection connection = new Connection(name, _selectedConn.ConnectionString);
            bool? result = connection.ShowDialog();

            if (!result.HasValue || !result.Value) return;

            var configExists = ConfigFileExists(_selectedProject);
            if (!configExists)
                CreateConfigFile(_selectedProject);

            Expander.IsExpanded = false;

            AddOrUpdateConnection(_selectedProject, connection.ConnectionName, connection.ConnectionString, connection.OrgId, connection.Version, false);

            GetConnections();
            foreach (CrmConn conn in Connections.Items)
            {
                if (conn.Name != connection.ConnectionName) continue;

                Connections.SelectedItem = conn;
                GetReports(connection.ConnectionString);
                break;
            }
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MessageBoxResult result = MessageBox.Show("Are you sure?" + Environment.NewLine + Environment.NewLine +
                    "This will delete the connection information and all associated mappings.", "Delete Connection", MessageBoxButton.YesNo);
                if (result != MessageBoxResult.Yes) return;

                if (_selectedConn == null) return;
                if (string.IsNullOrEmpty(_selectedConn.ConnectionString)) return;

                var path = Path.GetDirectoryName(_selectedProject.FullName);
                if (!ConfigFileExists(_selectedProject))
                {
                    _logger.WriteToOutputWindow("Error Deleting Connection: Missing CRMDeveloperExtensions.config File", Logger.MessageType.Error);
                    return;
                }

                if (!ConfigFileExists(_selectedProject)) return;

                //Delete Connection
                XmlDocument doc = new XmlDocument();
                doc.Load(path + "\\CRMDeveloperExtensions.config");
                XmlNodeList connections = doc.GetElementsByTagName("Connection");
                if (connections.Count == 0) return;

                List<XmlNode> nodesToRemove = new List<XmlNode>();
                foreach (XmlNode connection in connections)
                {
                    XmlNode orgId = connection["OrgId"];
                    if (orgId == null) continue;
                    if (orgId.InnerText.ToUpper() != _selectedConn.OrgId.ToUpper()) continue;

                    nodesToRemove.Add(connection);
                }

                foreach (XmlNode xmlNode in nodesToRemove)
                {
                    if (xmlNode.ParentNode != null)
                        xmlNode.ParentNode.RemoveChild(xmlNode);
                }
                doc.Save(path + "\\CRMDeveloperExtensions.config");

                //Delete related Files
                doc.Load(path + "\\CRMDeveloperExtensions.config");
                XmlNodeList files = doc.GetElementsByTagName("File");
                if (files.Count > 0)
                {
                    nodesToRemove = new List<XmlNode>();
                    foreach (XmlNode file in files)
                    {
                        XmlNode orgId = file["OrgId"];
                        if (orgId == null) continue;
                        if (orgId.InnerText.ToUpper() != _selectedConn.OrgId.ToUpper()) continue;

                        nodesToRemove.Add(file);
                    }

                    foreach (XmlNode xmlNode in nodesToRemove)
                    {
                        if (xmlNode.ParentNode != null)
                            xmlNode.ParentNode.RemoveChild(xmlNode);
                    }
                    doc.Save(path + "\\CRMDeveloperExtensions.config");
                }

                ReportGrid.ItemsSource = null;
                ShowManaged.IsEnabled = false;
                Publish.IsEnabled = false;
                Customizations.IsEnabled = false;
                Solutions.IsEnabled = false;
                Reports.IsEnabled = false;
                AddReport.IsEnabled = false;
                ReportGrid.IsEnabled = false;

                GetConnections();
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Deleting Connection: Missing CRMDeveloperExtensions.config File: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
            }
        }

        private void OpenReport_OnClick(object sender, RoutedEventArgs e)
        {
            Button button = (Button)sender;
            string page = "reportproperty.aspx";
            if (button.Name == "RunReport")
                page = "viewer/viewer.aspx";

            Guid reportId = new Guid(((Button)sender).CommandParameter.ToString());

            OpenCrmPage("crmreports/" + page + "?id=%7b" + reportId + "%7d");
        }

        private void ShowManaged_Checked(object sender, RoutedEventArgs e)
        {
            FilterReports();
        }

        private void FilterReports()
        {
            bool showManaged = ShowManaged.IsChecked != null && ShowManaged.IsChecked.Value;

            List<ReportItem> items = (List<ReportItem>)ReportGrid.ItemsSource;
            foreach (ReportItem reportItem in items)
            {
                if (reportItem.IsManaged && !showManaged)
                    reportItem.Publish = false;
            }

            //Filter the items
            ICollectionView icv = CollectionViewSource.GetDefaultView(ReportGrid.ItemsSource);
            if (icv == null) return;

            if (showManaged)
                //Show managed + unmanaged
                icv.Filter = null;
            else
            {
                icv.Filter = o =>
                {
                    ReportItem r = o as ReportItem;
                    //Show unmanaged only
                    return r != null && !r.IsManaged;
                };
            }


            //Item Count
            CollectionView cv = (CollectionView)icv;
            ItemCount.Text = cv.Count + " Items";
        }

        private void ReportGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //Make rows unselectable
            ReportGrid.UnselectAllCells();
        }

        private void AddReport_Click(object sender, RoutedEventArgs e)
        {
            NewReport newReport = new NewReport(_selectedConn, _selectedProject, GetProjectFiles(_selectedProject.Name));
            bool? result = newReport.ShowDialog();

            if (result != true) return;

            // Add new item
            ReportItem rItem = new ReportItem
            {
                Publish = false,
                ReportId = newReport.NewId,
                Name = newReport.NewName,
                IsManaged = false,
                AllowPublish = true,
                ProjectFiles = GetProjectFiles(_selectedProject.Name)
            };

            rItem.PropertyChanged += ReportItem_PropertyChanged;
            //Needs to be after setting the property changed event
            rItem.BoundFile = newReport.NewBoudndFile;

            List<ReportItem> items = (List<ReportItem>)ReportGrid.ItemsSource;
            items.Add(rItem);
            ReportGrid.ItemsSource = items.OrderBy(w => w.Name).ToList();

            var showManaged = ShowManaged.IsChecked;

            FilterReports();

            ShowManaged.IsChecked = showManaged;

            ReportGrid.ScrollIntoView(rItem);
        }

        private void Info_OnClick(object sender, RoutedEventArgs e)
        {
            Info info = new Info();
            info.Show();
        }

        private void UpdateAllPublishChecks(bool publish)
        {
            List<ReportItem> reports = (List<ReportItem>)ReportGrid.ItemsSource;
            foreach (ReportItem reportItem in reports)
            {
                if (reportItem.AllowPublish)
                    reportItem.Publish = publish;
            }
        }

        private void PublishSelectAll_OnClick(object sender, RoutedEventArgs e)
        {
            CheckBox publishAll = (CheckBox)sender;
            bool? isChecked = publishAll.IsChecked;

            if (isChecked != null && isChecked.Value)
                UpdateAllPublishChecks(true);
            else
                UpdateAllPublishChecks(false);
        }

        private void DeleteCache_OnClick(object sender, RoutedEventArgs e)
        {
            _dte.StatusBar.Text = "Clearing cache...";
            _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationGeneral);

            string[] items = GetRdlFiles(_selectedProject.ProjectItems);

            foreach (string file in items)
            {
                string cachePath = file + ".data";
                if (!File.Exists(cachePath)) continue;

                File.SetAttributes(cachePath, FileAttributes.Normal);
                File.Delete(cachePath);
            }

            _dte.StatusBar.Clear();
            _dte.StatusBar.Animate(false, vsStatusAnimation.vsStatusAnimationGeneral);
        }

        public static string[] GetRdlFiles(ProjectItems projectItems)
        {
            if (projectItems == null) return new string[] { };

            List<string> items = new List<string>();
            foreach (ProjectItem pi in projectItems)
            {
                if (pi.SubProject != null)
                    items.AddRange(GetRdlFiles(pi.SubProject.ProjectItems));
                else if (pi.Name.ToLower().EndsWith(".rdl"))
                    items.Add(pi.FileNames[1]);

                items.AddRange(GetRdlFiles(pi.ProjectItems));
            }
            return items.ToArray();
        }

        private void Customizations_OnClick(object sender, RoutedEventArgs e)
        {
            OpenCrmPage("tools/solution/edit.aspx?id=%7bfd140aaf-4df4-11dd-bd17-0019b9312238%7d");
        }

        private void Solutions_OnClick(object sender, RoutedEventArgs e)
        {
            OpenCrmPage("tools/Solution/home_solution.aspx?etc=7100&sitemappath=Settings|Customizations|nav_solution");
        }

        private void Reports_OnClick(object sender, RoutedEventArgs e)
        {
            OpenCrmPage("main.aspx?area=nav_reports&etc=9100&page=CS&pageType=EntityList&web=false");
        }

        private void OpenCrmPage(string url)
        {
            if (_selectedConn == null) return;
            string connString = _selectedConn.ConnectionString;
            if (string.IsNullOrEmpty(connString)) return;

            string[] connParts = connString.Split(';');
            string urlPart = connParts.FirstOrDefault(s => s.ToUpper().StartsWith("URL="));
            if (!string.IsNullOrEmpty(urlPart))
            {
                string[] urlParts = urlPart.Split('=');
                string baseUrl = (urlParts[1].EndsWith("/")) ? urlParts[1] : urlParts[1] + "/";

                var props = _dte.Properties["CRM Developer Extensions", "Settings"];
                bool useDefaultWebBrowser = (bool)props.Item("UseDefaultWebBrowser").Value;

                if (useDefaultWebBrowser) //User's default browser
                    System.Diagnostics.Process.Start(baseUrl + url);
                else //Internal VS browser
                    _dte.ItemOperations.Navigate(baseUrl + url);
            }
        }
    }
}