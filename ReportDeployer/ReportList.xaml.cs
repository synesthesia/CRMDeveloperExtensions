using CommonResources;
using EnvDTE;
using InfoWindow;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using OutputLogger;
using ReportDeployer.Models;
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
using System.Windows.Data;
using System.Windows.Media;
using System.Xml;
using Button = System.Windows.Controls.Button;
using CheckBox = System.Windows.Controls.CheckBox;
using Task = System.Threading.Tasks.Task;
using Window = EnvDTE.Window;

namespace ReportDeployer
{
    public partial class ReportList
    {
        //All DTE objects need to be declared here or else things stop working
        private readonly DTE _dte;
        private readonly Solution _solution;
        private readonly IVsSolution _vsSolution;
        private bool _projectEventsRegistered;
        private readonly Logger _logger;
        private const string WindowType = "ReportDeployer";

        public ReportList()
        {
            uint solutionEventsCookie;
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

            var vsSolutionEvents = new VsSolutionEvents(this);
            _vsSolution = (IVsSolution)ServiceProvider.GlobalProvider.GetService(typeof(SVsSolution));
            _vsSolution.AdviseSolutionEvents(vsSolutionEvents, out solutionEventsCookie);
        }

        private void WindowEventsOnWindowActivated(Window gotFocus, Window lostFocus)
        {
            //No solution loaded
            if (_solution.Count == 0)
            {
                ResetForm();
                return;
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
            foreach (Project project in ConnPane.Projects)
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
                string[] files = GetRdlFiles(ConnPane.SelectedProject.ProjectItems);
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
                    reportItem.ProjectFiles = GetProjectFiles(ConnPane.SelectedProject.Name);
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
            return ConnPane.SelectedProject == project;
        }

        private void SolutionProjectRemoved(Project project)
        {
            if (ConnPane.SelectedProject == null || ConnPane.SelectedProject.FullName != project.FullName) return;

            ReportGrid.ItemsSource = null;
            Publish.IsEnabled = false;
            Customizations.IsEnabled = false;
            Reports.IsEnabled = false;
            Solutions.IsEnabled = false;
            AddReport.IsEnabled = false;
        }

        private void ResetForm()
        {
            ReportGrid.ItemsSource = null;
            Publish.IsEnabled = false;
            Customizations.IsEnabled = false;
            Reports.IsEnabled = false;
            Solutions.IsEnabled = false;
            AddReport.IsEnabled = false;
        }

        private async void ConnPane_OnConnectionAdded(object sender, ConnectionAddedEventArgs e)
        {
            bool gotReports = await GetReports(e.AddedConnection.ConnectionString);

            if (!gotReports)
            {
                Customizations.IsEnabled = false;
                Solutions.IsEnabled = false;
                Reports.IsEnabled = false;
                return;
            }

            Customizations.IsEnabled = true;
            Solutions.IsEnabled = true;
            Reports.IsEnabled = true;
        }

        private Project GetProjectByName(string projectName)
        {
            foreach (Project project in ConnPane.Projects)
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
            if (projectItem.Kind.ToUpper() != "{6BB5F8EF-4483-11D3-8BCF-00C04F8EC28C}") // VS Folder 
            {
                string ex = Path.GetExtension(projectItem.Name);
                if (ex != null && ex.ToUpper() == ".RDL")
                    projectFiles.Add(new ComboBoxItem { Content = path + "/" + projectItem.Name, Tag = projectItem });
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

        private void ConnPane_OnConnectionStarted(object sender, EventArgs e)
        {
            _dte.StatusBar.Text = "Connecting to CRM...";
            _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationSync);
        }

        private async void ConnPane_OnConnected(object sender, ConnectEventArgs e)
        {
            bool gotReports = await GetReports(e.ConnectionString);

            if (!gotReports)
            {
                Customizations.IsEnabled = false;
                Solutions.IsEnabled = false;
                Reports.IsEnabled = false;
                return;
            }

            Customizations.IsEnabled = true;
            Solutions.IsEnabled = true;
            Reports.IsEnabled = true;
        }

        private async Task<bool> GetReports(string connString)
        {
            string projectName = ConnPane.SelectedProject.Name;
            _dte.StatusBar.Text = "Connecting to CRM...";
            CrmServiceClient client = SharedConnection.GetCurrentConnection(connString, WindowType, _dte);

            _dte.StatusBar.Text = "Getting reports...";
            _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationSync);
            LockMessage.Content = "Working...";
            LockOverlay.Visibility = Visibility.Visible;

            EntityCollection results = await Task.Run(() => RetrieveReportsFromCrm(client));
            if (results == null)
            {
                SharedConnection.ClearCurrentConnection(WindowType, _dte);
                _dte.StatusBar.Clear();
                _dte.StatusBar.Animate(false, vsStatusAnimation.vsStatusAnimationSync);
                LockOverlay.Visibility = Visibility.Hidden;
                MessageBox.Show("Error Retrieving Reports. See the Output Window for additional details.");
                return false;
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

            return true;
        }

        private EntityCollection RetrieveReportsFromCrm(CrmServiceClient client)
        {
            try
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

                return client.RetrieveMultiple(query);
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
                string projectName = ConnPane.SelectedProject.Name;
                Project project = GetProjectByName(projectName);
                if (project == null)
                    return new List<ReportItem>();

                var path = Path.GetDirectoryName(project.FullName);
                if (!SharedConfigFile.ConfigFileExists(project))
                {
                    _logger.WriteToOutputWindow("Error Updating Mappings In Config File: Missing CRMDeveloperExtensions.config File", Logger.MessageType.Error);
                    return rItems;
                }

                XmlDocument doc = new XmlDocument();
                doc.Load(path + "\\CRMDeveloperExtensions.config");

                var props = _dte.Properties["CRM Developer Extensions", "Report Deployer"];
                bool allowPublish = (bool)props.Item("AllowPublishManagedReports").Value;

                XmlNodeList mappedFiles = doc.GetElementsByTagName("File");
                List<XmlNode> nodesToRemove = new List<XmlNode>();

                foreach (ReportItem rItem in rItems)
                {
                    foreach (XmlNode file in mappedFiles)
                    {
                        XmlNode orgIdNode = file["OrgId"];
                        if (orgIdNode == null) continue;
                        if (orgIdNode.InnerText.ToUpper() != ConnPane.SelectedConnection.OrgId.ToUpper()) continue;

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
                    if (orgIdNode.InnerText.ToUpper() != ConnPane.SelectedConnection.OrgId.ToUpper()) continue;

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
                var projectPath = Path.GetDirectoryName(ConnPane.SelectedProject.FullName);
                if (!SharedConfigFile.ConfigFileExists(ConnPane.SelectedProject))
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
                        bool changed = false;
                        XmlNode orgId = node["OrgId"];
                        if (orgId != null && orgId.InnerText.ToUpper() != ConnPane.SelectedConnection.OrgId.ToUpper()) continue;

                        XmlNode exReportId = node["ReportId"];
                        if (exReportId != null && exReportId.InnerText.ToUpper() !=
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

                                changed = true;
                                item.Publish = false;
                                item.AllowPublish = false;
                            }
                        }
                        else
                        {
                            //Update
                            XmlNode exPath = node["Path"];
                            if (exPath != null)
                            {
                                string oldPath = exPath.InnerText;
                                if (oldPath != item.BoundFile)
                                {
                                    exPath.InnerText = item.BoundFile;
                                    changed = true;
                                }
                            }
                        }

                        if (!changed)
                            return;

                        if (SharedConfigFile.IsConfigReadOnly(projectPath + "\\CRMDeveloperExtensions.config"))
                        {
                            FileInfo file = new FileInfo(projectPath + "\\CRMDeveloperExtensions.config") { IsReadOnly = false };
                        }

                        doc.Save(projectPath + "\\CRMDeveloperExtensions.config");
                        return;
                    }
                }

                //Create new mapping
                XmlNodeList files = doc.GetElementsByTagName("Files");
                if (files.Count <= 0)
                    return;

                XmlNode fileNode = doc.CreateElement("File");
                XmlNode org = doc.CreateElement("OrgId");
                org.InnerText = ConnPane.SelectedConnection.OrgId;
                fileNode.AppendChild(org);
                XmlNode newPath = doc.CreateElement("Path");
                newPath.InnerText = item.BoundFile;
                fileNode.AppendChild(newPath);
                XmlNode newReportId = doc.CreateElement("ReportId");
                newReportId.InnerText = item.ReportId.ToString();
                fileNode.AppendChild(newReportId);
                XmlNode name = doc.CreateElement("Name");
                name.InnerText = item.Name;
                fileNode.AppendChild(name);
                XmlNode isManaged = doc.CreateElement("IsManaged");
                isManaged.InnerText = item.IsManaged.ToString();
                fileNode.AppendChild(isManaged);
                files[0].AppendChild(fileNode);

                if (SharedConfigFile.IsConfigReadOnly(projectPath + "\\CRMDeveloperExtensions.config"))
                {
                    FileInfo file = new FileInfo(projectPath + "\\CRMDeveloperExtensions.config") { IsReadOnly = false };
                }

                doc.Save(projectPath + "\\CRMDeveloperExtensions.config");

                item.AllowPublish = true;
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

            string connString = ConnPane.SelectedConnection.ConnectionString;
            string projectName = ConnPane.SelectedProject.Name;
            DownloadReport(reportId, folder, connString, projectName);
        }

        private void DownloadReport(Guid reportId, string folder, string connString, string projectName)
        {
            _dte.StatusBar.Text = "Downloading file...";
            _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationSync);

            try
            {
                CrmServiceClient client = SharedConnection.GetCurrentConnection(connString, WindowType, _dte);
                Entity report = client.Retrieve("report", reportId,
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
            string projectName = ConnPane.SelectedProject.Name;
            Project project = GetProjectByName(projectName);
            if (project == null) return;

            string connString = ConnPane.SelectedConnection.ConnectionString;
            if (connString == null) return;
            CrmServiceClient client = SharedConnection.GetCurrentConnection(connString, WindowType, _dte);

            LockMessage.Content = "Deploying...";
            LockOverlay.Visibility = Visibility.Visible;

            bool success;
            //Check if < CRM 2011 UR12 (ExecuteMutliple)
            Version version = Version.Parse(ConnPane.SelectedConnection.Version);
            if (version.Major == 5 && version.Revision < 3200)
                success = await System.Threading.Tasks.Task.Run(() => UpdateAndPublishSingle(items, project, client));
            else
                success = await System.Threading.Tasks.Task.Run(() => UpdateAndPublishMultiple(items, project, client));

            LockOverlay.Visibility = Visibility.Hidden;

            if (success) return;

            MessageBox.Show("Error Updating Reports. See the Output Window for additional details.");
            _dte.StatusBar.Clear();
        }

        private bool UpdateAndPublishMultiple(List<ReportItem> items, Project project, CrmServiceClient client)
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
                _dte.StatusBar.Text = "Updating report(s)...";
                _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationDeploy);

                ExecuteMultipleResponse emResponse = (ExecuteMultipleResponse)client.Execute(emRequest);

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

        private bool UpdateAndPublishSingle(List<ReportItem> items, Project project, CrmServiceClient client)
        {
            //CRM 2011 < UR12
            _dte.StatusBar.Text = "Updating report(s)...";
            _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationDeploy);

            try
            {
                foreach (var reportItem in items)
                {
                    Entity report = new Entity("report") { Id = reportItem.ReportId };

                    string filePath = Path.GetDirectoryName(project.FullName) +
                                      reportItem.BoundFile.Replace("/", "\\");
                    if (!File.Exists(filePath)) continue;

                    report["bodytext"] = File.ReadAllText(filePath);

                    UpdateRequest request = new UpdateRequest { Target = report };
                    client.Execute(request);
                    _logger.WriteToOutputWindow("Uploaded Report", Logger.MessageType.Info);
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

        private void ConnPane_OnConnectionSelected(object sender, ConnectionSelectedEventArgs e)
        {
            if (e.SelectedConnection != null)
            {
                if (e.ConnectionAdded)
                {
                    Customizations.IsEnabled = true;
                    Solutions.IsEnabled = true;
                    Reports.IsEnabled = true;
                    AddReport.IsEnabled = true;
                }
                else
                {
                    Customizations.IsEnabled = false;
                    Solutions.IsEnabled = false;
                    Reports.IsEnabled = false;
                    AddReport.IsEnabled = false;
                }
            }
            else
            {
                Customizations.IsEnabled = false;
                Solutions.IsEnabled = false;
                Reports.IsEnabled = false;
                AddReport.IsEnabled = false;
            }

            ReportGrid.ItemsSource = null;
            ShowManaged.IsEnabled = false;
            Publish.IsEnabled = false;
            ReportGrid.IsEnabled = false;
        }

        private void ConnPane_OnProjectChanged(object sender, EventArgs e)
        {
            ReportGrid.ItemsSource = null;
            ShowManaged.IsEnabled = false;
            Publish.IsEnabled = false;
            Customizations.IsEnabled = false;
            Solutions.IsEnabled = false;
            Reports.IsEnabled = false;
            AddReport.IsEnabled = false;
            ReportGrid.IsEnabled = false;
        }

        private async void ConnPane_OnConnectionModified(object sender, ConnectionModifiedEventArgs e)
        {
            bool gotReports = await GetReports(e.ModifiedConnection.ConnectionString);

            if (!gotReports)
            {
                Customizations.IsEnabled = false;
                Solutions.IsEnabled = false;
                Reports.IsEnabled = false;
                return;
            }

            Customizations.IsEnabled = true;
            Solutions.IsEnabled = true;
            Reports.IsEnabled = true;
        }

        private void ConnPane_OnConnectionDeleted(object sender, EventArgs e)
        {
            ReportGrid.ItemsSource = null;
            ShowManaged.IsEnabled = false;
            Publish.IsEnabled = false;
            Customizations.IsEnabled = false;
            Solutions.IsEnabled = false;
            Reports.IsEnabled = false;
            AddReport.IsEnabled = false;
            ReportGrid.IsEnabled = false;
        }

        private void OpenReport_OnClick(object sender, RoutedEventArgs e)
        {
            Button button = (Button)sender;
            string page = "reportproperty.aspx";
            if (button.Name == "RunReport")
                page = "viewer/viewer.aspx";

            Guid reportId = new Guid(((Button)sender).CommandParameter.ToString());

            SharedWindow.OpenCrmPage("crmreports/" + page + "?id=%7b" + reportId + "%7d",
                ConnPane.SelectedConnection, _dte);
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
            NewReport newReport = new NewReport(ConnPane.SelectedConnection, ConnPane.SelectedProject, GetProjectFiles(ConnPane.SelectedProject.Name));
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
                ProjectFiles = GetProjectFiles(ConnPane.SelectedProject.Name)
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
            info.ShowDialog();
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

            string[] items = GetRdlFiles(ConnPane.SelectedProject.ProjectItems);

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
            SharedWindow.OpenCrmPage("tools/solution/edit.aspx?id=%7bfd140aaf-4df4-11dd-bd17-0019b9312238%7d",
                ConnPane.SelectedConnection, _dte);
        }

        private void Solutions_OnClick(object sender, RoutedEventArgs e)
        {
            SharedWindow.OpenCrmPage("tools/Solution/home_solution.aspx?etc=7100&sitemappath=Settings|Customizations|nav_solution",
                ConnPane.SelectedConnection, _dte);
        }

        private void Reports_OnClick(object sender, RoutedEventArgs e)
        {
            SharedWindow.OpenCrmPage("main.aspx?area=nav_reports&etc=9100&page=CS&pageType=EntityList&web=false",
                ConnPane.SelectedConnection, _dte);
        }
    }
}