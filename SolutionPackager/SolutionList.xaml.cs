using CommonResources;
using EnvDTE;
using EnvDTE80;
using InfoWindow;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using OutputLogger;
using SolutionPackager.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Xml;
using DialogResult = System.Windows.Forms.DialogResult;
using MessageBox = System.Windows.MessageBox;
using Window = EnvDTE.Window;

namespace SolutionPackager
{
    public partial class SolutionList
    {
        private readonly DTE _dte;
        private readonly DTE2 _dte2;
        private readonly Logger _logger;
        private static OrganizationService _orgService;

        public SolutionList()
        {
            InitializeComponent();

            _logger = new Logger();

            _dte = Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(DTE)) as DTE;
            if (_dte == null)
                return;

            _dte2 = Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(DTE)) as DTE2;
            if (_dte2 == null)
                return;

            var solution = _dte.Solution;
            if (solution == null)
                return;

            var events = _dte.Events;
            var solutionEvents = events.SolutionEvents;
            solutionEvents.BeforeClosing += SolutionBeforeClosing;
            solutionEvents.ProjectRemoved += SolutionProjectRemoved;

            SetDownloadManagedEnabled(false);
            DownloadManaged.IsChecked = false;
        }

        private void SetDownloadManagedEnabled(bool enabled)
        {
            var props = _dte.Properties["CRM Developer Extensions", "Solution Packager"];
            bool saveSolutionFiles = (bool)props.Item("SaveSolutionFiles").Value;
            if (!saveSolutionFiles)
            {
                DownloadManaged.IsEnabled = false;
                return;
            }

            DownloadManaged.IsEnabled = enabled;
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
                }
            }
            else
            {
                Customizations.IsEnabled = false;
                Solutions.IsEnabled = false;
            }

            Package.IsEnabled = false;
            Unpackage.IsEnabled = false;
            SetDownloadManagedEnabled(false);
            DownloadManaged.IsChecked = false;
        }

        private bool ConfigFileExists(Project project)
        {
            var path = Path.GetDirectoryName(project.FullName);
            return File.Exists(path + "/CRMDeveloperExtensions.config");
        }

        private void ConnPane_OnConnectionAdded(object sender, ConnectionAddedEventArgs e)
        {
            GetSolutions(e.AddedConnection.ConnectionString);
        }

        private void ConnPane_OnConnectionDeleted(object sender, EventArgs e)
        {
            Customizations.IsEnabled = false;
            Solutions.IsEnabled = false;
            Package.IsEnabled = false;
            Unpackage.IsEnabled = false;
            SolutionToPackage.ItemsSource = null;
            SetDownloadManagedEnabled(false);
            DownloadManaged.IsChecked = false;
        }

        private void ConnPane_OnConnected(object sender, ConnectEventArgs e)
        {
            GetSolutions(e.ConnectionString);

            Customizations.IsEnabled = true;
            Solutions.IsEnabled = true;
            SetDownloadManagedEnabled(true);
        }

        private void ConnPane_OnProjectChanged(object sender, EventArgs e)
        {
            Customizations.IsEnabled = false;
            Solutions.IsEnabled = false;
            Package.IsEnabled = false;
            Unpackage.IsEnabled = false;
            SolutionToPackage.ItemsSource = null;
            SetDownloadManagedEnabled(false);
            DownloadManaged.IsChecked = false;
        }

        private void ConnPane_OnConnectionModified(object sender, ConnectionModifiedEventArgs e)
        {
            GetSolutions(e.ModifiedConnection.ConnectionString);
        }

        private void Info_OnClick(object sender, RoutedEventArgs e)
        {
            Info info = new Info();
            info.ShowDialog();
        }

        private void Customizations_OnClick(object sender, RoutedEventArgs e)
        {
            OpenCrmPage("tools/solution/edit.aspx?id=%7bfd140aaf-4df4-11dd-bd17-0019b9312238%7d");
        }

        private void Solutions_OnClick(object sender, RoutedEventArgs e)
        {
            OpenCrmPage("tools/Solution/home_solution.aspx?etc=7100&sitemappath=Settings|Customizations|nav_solution");
        }

        private void OpenCrmPage(string url)
        {
            if (ConnPane.SelectedConnection == null) return;
            string connString = ConnPane.SelectedConnection.ConnectionString;
            if (string.IsNullOrEmpty(connString)) return;

            string[] connParts = connString.Split(';');
            string urlPart = connParts.FirstOrDefault(s => s.ToUpper().StartsWith("URL="));
            if (!string.IsNullOrEmpty(urlPart))
            {
                string[] urlParts = urlPart.Split('=');
                string baseUrl = (urlParts[1].EndsWith("/")) ? urlParts[1] : urlParts[1] + "/";

                var props = _dte.Properties["CRM Developer Extensions", "General"];
                bool useDefaultWebBrowser = (bool)props.Item("UseDefaultWebBrowser").Value;

                if (useDefaultWebBrowser) //User's default browser
                    System.Diagnostics.Process.Start(baseUrl + url);
                else //Internal VS browser
                    _dte.ItemOperations.Navigate(baseUrl + url);
            }
        }

        private async void GetSolutions(string connString)
        {
            _dte.StatusBar.Text = "Connecting to CRM and getting solutions...";
            _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationSync);
            LockMessage.Content = "Working...";
            LockOverlay.Visibility = Visibility.Visible;

            EntityCollection results = await Task.Run(() => GetSolutionsFromCrm(connString));
            if (results == null)
            {
                _dte.StatusBar.Clear();
                LockOverlay.Visibility = Visibility.Hidden;
                MessageBox.Show("Error Retrieving Solutions. See the Output Window for additional details.");
                return;
            }

            _logger.WriteToOutputWindow("Retrieved Solutions From CRM", Logger.MessageType.Info);

            ObservableCollection<CrmSolution> solutions = new ObservableCollection<CrmSolution>();

            CrmSolution emptyItem = new CrmSolution
            {
                SolutionId = Guid.Empty,
                Name = String.Empty
            };
            solutions.Add(emptyItem);

            foreach (Entity entity in results.Entities)
            {
                CrmSolution solution = new CrmSolution
                {
                    SolutionId = entity.Id,
                    Name = entity.GetAttributeValue<string>("friendlyname"),
                    Prefix = entity.GetAttributeValue<AliasedValue>("publisher.customizationprefix").Value.ToString(),
                    UniqueName = entity.GetAttributeValue<string>("uniquename"),
                    Version = Version.Parse(entity.GetAttributeValue<string>("version"))
                };

                solutions.Add(solution);
            }

            //Empty on top
            var i =
                solutions.IndexOf(
                    solutions.FirstOrDefault(s => s.SolutionId == new Guid("00000000-0000-0000-0000-000000000000")));

            var item = solutions[i];
            solutions.RemoveAt(i);
            solutions.Insert(0, item);

            //Default second
            i =
                solutions.IndexOf(
                    solutions.FirstOrDefault(s => s.SolutionId == new Guid("FD140AAF-4DF4-11DD-BD17-0019B9312238")));
            item = solutions[i];
            solutions.RemoveAt(i);
            solutions.Insert(1, item);

            solutions = HandleMappings(solutions);
            SolutionToPackage.ItemsSource = solutions;
            if (solutions.Count(s => !string.IsNullOrEmpty(s.BoundProject)) > 0)
            {
                SolutionToPackage.SelectedItem = solutions.First(s => !string.IsNullOrEmpty(s.BoundProject));
                SolutionToPackage.IsEnabled = false;
            }
            else
                SolutionToPackage.IsEnabled = true;

            CrmSolution crmSolution = solutions.FirstOrDefault(s => !string.IsNullOrEmpty(s.BoundProject));
            DownloadManaged.IsChecked = crmSolution != null && solutions.First(s => !string.IsNullOrEmpty(s.BoundProject)).DownloadManagedSolution;

            _dte.StatusBar.Clear();
            _dte.StatusBar.Animate(false, vsStatusAnimation.vsStatusAnimationSync);
            LockOverlay.Visibility = Visibility.Hidden;
        }

        private ObservableCollection<CrmSolution> HandleMappings(ObservableCollection<CrmSolution> sItems)
        {
            try
            {
                string projectName = ConnPane.SelectedProject.Name;
                Project project = GetProjectByName(projectName);
                if (project == null)
                    return new ObservableCollection<CrmSolution>();

                var path = Path.GetDirectoryName(project.FullName);
                if (!ConfigFileExists(project))
                {
                    _logger.WriteToOutputWindow("Error Updating Mapping In Config File: Missing CRMDeveloperExtensions.config File", Logger.MessageType.Error);
                    return new ObservableCollection<CrmSolution>();
                }

                XmlDocument doc = new XmlDocument();
                doc.Load(path + "\\CRMDeveloperExtensions.config");

                List<string> projectNames = new List<string>();
                foreach (Project p in ConnPane.Projects)
                {
                    projectNames.Add(p.Name.ToUpper());
                }

                XmlNodeList solutionNodes = doc.GetElementsByTagName("Solution");
                List<XmlNode> nodesToRemove = new List<XmlNode>();

                foreach (CrmSolution sItem in sItems)
                {
                    foreach (XmlNode solutionNode in solutionNodes)
                    {
                        XmlNode orgIdNode = solutionNode["OrgId"];
                        if (orgIdNode == null) continue;
                        if (orgIdNode.InnerText.ToUpper() != ConnPane.SelectedConnection.OrgId.ToUpper()) continue;

                        XmlNode solutionId = solutionNode["SolutionId"];
                        if (solutionId == null) continue;
                        if (solutionId.InnerText.ToUpper() != sItem.SolutionId.ToString().ToUpper()) continue;

                        XmlNode projectNameNode = solutionNode["ProjectName"];
                        if (projectNameNode == null) continue;

                        if (!projectNames.Contains(projectNameNode.InnerText.ToUpper()))
                            //Remove mappings for projects that might have been deleted from the solution
                            nodesToRemove.Add(projectNameNode);
                        else
                            sItem.BoundProject = projectNameNode.InnerText;

                        XmlNode downloadManagedNode = solutionNode["DownloadManaged"];
                        bool downloadManaged = false;
                        bool hasDownloadManaged = downloadManagedNode != null && bool.TryParse(downloadManagedNode.InnerText, out downloadManaged);
                        sItem.DownloadManagedSolution = hasDownloadManaged && downloadManaged;
                    }
                }

                //Remove mappings for solutions that might have been deleted from CRM
                solutionNodes = doc.GetElementsByTagName("Solution");
                foreach (XmlNode solution in solutionNodes)
                {
                    XmlNode orgIdNode = solution["OrgId"];
                    if (orgIdNode == null) continue;
                    if (orgIdNode.InnerText.ToUpper() != ConnPane.SelectedConnection.OrgId.ToUpper()) continue;

                    XmlNode solutionId = solution["SolutionId"];
                    if (solutionId == null) continue;

                    var count = sItems.Count(s => s.SolutionId.ToString().ToUpper() == solutionId.InnerText.ToUpper());
                    if (count == 0)
                        nodesToRemove.Add(solution);
                }

                //Remove the invalid mappings
                if (nodesToRemove.Count <= 0)
                    return sItems;

                XmlNode projects = nodesToRemove[0].ParentNode;
                foreach (XmlNode xmlNode in nodesToRemove)
                {
                    if (projects != null && projects.ParentNode != null)
                        projects.RemoveChild(xmlNode);
                }
                doc.Save(path + "\\CRMDeveloperExtensions.config");
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Updating Mappings In Config File: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
            }

            return sItems;
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

        private EntityCollection GetSolutionsFromCrm(string connString)
        {
            try
            {
                CrmConnection connection = CrmConnection.Parse(connString);

                using (_orgService = new OrganizationService(connection))
                {
                    QueryExpression query = new QueryExpression
                    {
                        EntityName = "solution",
                        ColumnSet = new ColumnSet("friendlyname", "solutionid", "uniquename", "version"),
                        Criteria = new FilterExpression
                        {
                            Conditions =
                            {
                                new ConditionExpression
                                {
                                    AttributeName = "ismanaged",
                                    Operator = ConditionOperator.Equal,
                                    Values = {false}
                                },
                                new ConditionExpression
                                {
                                    AttributeName = "isvisible",
                                    Operator = ConditionOperator.Equal,
                                    Values = {true}
                                }
                            }
                        },
                        LinkEntities =
                        {
                            new LinkEntity
                            {
                                LinkFromEntityName = "solution",
                                LinkFromAttributeName = "publisherid",
                                LinkToEntityName = "publisher",
                                LinkToAttributeName = "publisherid",
                                Columns = new ColumnSet("customizationprefix"),
                                EntityAlias = "publisher"
                            }
                        },
                        Orders =
                        {
                            new OrderExpression
                            {
                                AttributeName = "friendlyname",
                                OrderType = OrderType.Ascending
                            }
                        }
                    };

                    return _orgService.RetrieveMultiple(query);
                }
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                _logger.WriteToOutputWindow("Error Retrieving Solutions From CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, Logger.MessageType.Error);
                return null;
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Retrieving Solutions From CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
                return null;
            }
        }

        private void SolutionBeforeClosing()
        {
            //Close the tool window - forces having to reopen for a new solution
            foreach (Window window in _dte.Windows)
            {
                if (window.Caption != SolutionPackager.Resources.ResourceManager.GetString("ToolWindowTitle")) continue;

                ResetForm();
                _logger.DeleteOutputWindow();
                window.Close();
                return;
            }
        }

        private void ResetForm()
        {
            Package.IsEnabled = false;
            Unpackage.IsEnabled = false;
            Customizations.IsEnabled = false;
            Solutions.IsEnabled = false;
            SetDownloadManagedEnabled(false);
            DownloadManaged.IsChecked = false;
        }

        private void SolutionProjectRemoved(Project project)
        {
            if (ConnPane.SelectedProject == null)
                return;
            if (ConnPane.SelectedProject.FullName != project.FullName)
                return;

            SolutionToPackage.ItemsSource = null;
            Package.IsEnabled = false;
            Unpackage.IsEnabled = false;
            Customizations.IsEnabled = false;
            Solutions.IsEnabled = false;
            SetDownloadManagedEnabled(false);
            DownloadManaged.IsChecked = false;
        }

        private async void Unpackage_OnClick(object sender, RoutedEventArgs e)
        {
            SolutionToPackage.IsEnabled = false;

            _dte.StatusBar.Text = "Connecting to CRM and getting unmanaged solution...";
            _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationSync);
            LockMessage.Content = "Working...";
            LockOverlay.Visibility = Visibility.Visible;

            if (SolutionToPackage.SelectedItem == null)
                return;

            CrmSolution selectedSolution = (CrmSolution)SolutionToPackage.SelectedItem;
            bool? downloadManaged = DownloadManaged.IsChecked;

            string path = await Task.Run(() => GetSolutionFromCrm(ConnPane.SelectedConnection.ConnectionString, selectedSolution, false));
            if (string.IsNullOrEmpty(path))
            {
                _dte.StatusBar.Clear();
                LockOverlay.Visibility = Visibility.Hidden;
                MessageBox.Show("Error Retrieving Unmanaged Solution. See the Output Window for additional details.");
                return;
            }

            _logger.WriteToOutputWindow("Retrieved Unmanaged Solution From CRM", Logger.MessageType.Info);
            _dte.StatusBar.Text = "Extracting solution...";

            Project project = ConnPane.SelectedProject;
            var path1 = path;
            bool solutionChange = await Task.Run(() => ExtractPackage(path1, selectedSolution, project, downloadManaged));

            if (downloadManaged == true)
            {
                _dte.StatusBar.Text = "Connecting to CRM and getting managed solution...";
                path =
                    await
                        Task.Run(
                            () =>
                                GetSolutionFromCrm(ConnPane.SelectedConnection.ConnectionString, selectedSolution, true));

                if (string.IsNullOrEmpty(path))
                {
                    _dte.StatusBar.Clear();
                    LockOverlay.Visibility = Visibility.Hidden;
                    MessageBox.Show("Error Retrieving Managed Solution. See the Output Window for additional details.");
                    return;
                }

                _logger.WriteToOutputWindow("Retrieved Managed Solution From CRM", Logger.MessageType.Info);
                StoreSolutionFile(path, project, solutionChange);
            }

            _dte.StatusBar.Clear();
            _dte.StatusBar.Animate(false, vsStatusAnimation.vsStatusAnimationSync);
            LockOverlay.Visibility = Visibility.Hidden;
        }

        private async Task<bool> ExtractPackage(string path, CrmSolution selectedSolution, Project project, bool? downloadManaged)
        {
            //https://msdn.microsoft.com/en-us/library/jj602987.aspx#arguments
            try
            {
                CommandWindow cw = _dte2.ToolWindows.CommandWindow;

                var props = _dte.Properties["CRM Developer Extensions", "Solution Packager"];
                string spPath = (string)props.Item("SolutionPackagerPath").Value;

                if (string.IsNullOrEmpty(spPath))
                {
                    MessageBox.Show("Set SDK bin folder path under Tools -> Options -> CRM Developer Extensions");
                    return false;
                }

                if (!spPath.EndsWith("\\"))
                    spPath += "\\";

                string toolPath = @"""" + spPath + "SolutionPackager.exe" + @"""";

                string tempDirectory = Path.GetDirectoryName(path);
                if (Directory.Exists(tempDirectory + "\\" + Path.GetFileNameWithoutExtension(path)))
                    Directory.Delete(tempDirectory + "\\" + Path.GetFileNameWithoutExtension(path), true);
                DirectoryInfo extractedFolder =
                    Directory.CreateDirectory(tempDirectory + "\\" + Path.GetFileNameWithoutExtension(path));

                string command = toolPath + " /action: Extract";
                command += " /zipfile:" + "\"" + path + "\"";
                command += " /folder: " + "\"" + extractedFolder.FullName + "\"";
                command += " /clobber";

                cw.SendInput("shell " + command, true);

                //TODO: Adjust for larger solutions?
                System.Threading.Thread.Sleep(1000);

                bool solutionFileDelete = RemoveDeletedItems(extractedFolder.FullName, ConnPane.SelectedProject.ProjectItems);
                bool solutionFileAddChange = ProcessDownloadedSolution(extractedFolder, Path.GetDirectoryName(ConnPane.SelectedProject.FullName),
                    ConnPane.SelectedProject.ProjectItems);

                Directory.Delete(extractedFolder.FullName, true);

                //Solution change or file not present
                bool solutionChange = solutionFileDelete || solutionFileAddChange;
                StoreSolutionFile(path, project, solutionChange);

                //Download Managed
                if (downloadManaged != true)
                    return false;

                path = await Task.Run(() => GetSolutionFromCrm(ConnPane.SelectedConnection.ConnectionString, selectedSolution, true));

                if (string.IsNullOrEmpty(path))
                {
                    _dte.StatusBar.Clear();
                    LockOverlay.Visibility = Visibility.Hidden;
                    MessageBox.Show("Error Retrieving Solution. See the Output Window for additional details.");
                    return false;
                }

                return solutionChange;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error launching Solution Packager: " + Environment.NewLine + Environment.NewLine + ex.Message);
                return false;
            }
        }

        private void StoreSolutionFile(string path, Project project, bool solutionChange)
        {
            try
            {
                var props = _dte.Properties["CRM Developer Extensions", "Solution Packager"];
                bool saveSolutionFiles = (bool)props.Item("SaveSolutionFiles").Value;

                if (!saveSolutionFiles)
                    return;

                string projectPath = Path.GetDirectoryName(project.FullName);
                if (!Directory.Exists(projectPath + "\\" + "_Solutions"))
                    project.ProjectItems.AddFolder("_Solutions");

                string filename = Path.GetFileName(path);
                if (File.Exists(projectPath + "\\_Solutions\\" + filename))
                    File.Delete(projectPath + "\\" + "_Solutions\\" + filename);

                if (!solutionChange && File.Exists(projectPath + "\\_Solutions\\" + filename))
                    return;

                File.Move(path, projectPath + "\\" + "_Solutions\\" + filename);

                project.ProjectItems.AddFromFile(projectPath + "\\" + "_Solutions\\" + filename);
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error storing solution file to project: " + path +
                    Environment.NewLine + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
            }
        }

        private bool ProcessDownloadedSolution(DirectoryInfo extractedFolder, string baseFolder, ProjectItems projectItems)
        {
            bool itemChanged = false;

            //Handle file adds
            foreach (FileInfo file in extractedFolder.GetFiles())
            {
                if (File.Exists(baseFolder + "\\" + file.Name))
                {
                    if (FileEquals(baseFolder + "\\" + file.Name, file.FullName))
                        continue;
                }

                File.Copy(file.FullName, baseFolder + "\\" + file.Name, true);
                projectItems.AddFromFile(baseFolder + "\\" + file.Name);
                itemChanged = true;
            }

            //Handle folder adds
            foreach (DirectoryInfo folder in extractedFolder.GetDirectories())
            {
                if (!Directory.Exists(baseFolder + "\\" + folder.Name))
                    Directory.CreateDirectory(baseFolder + "\\" + folder.Name);

                var newProjectItems = projectItems;
                bool subItemChanged = ProcessDownloadedSolution(folder, baseFolder + "\\" + folder.Name, newProjectItems);
                if (subItemChanged)
                    itemChanged = true;
            }

            return itemChanged;
        }

        private static bool RemoveDeletedItems(string extractedFolder, ProjectItems projectItems)
        {
            bool itemChanged = false;

            //Handle file & folder deletes
            foreach (ProjectItem projectItem in projectItems)
            {
                string name = projectItem.FileNames[0];
                switch (projectItem.Kind.ToUpper())
                {
                    case "{6BB5F8EE-4483-11D3-8BCF-00C04F8EC28C}":
                        name = Path.GetFileName(name);
                        if (File.Exists(extractedFolder + "\\" + name))
                            continue;

                        projectItem.Delete();
                        itemChanged = true;
                        break;
                    case "{6BB5F8EF-4483-11D3-8BCF-00C04F8EC28C}":
                        name = new DirectoryInfo(name).Name;
                        if (name == "_Solutions")
                            continue;

                        if (!Directory.Exists(extractedFolder + "\\" + name))
                        {
                            projectItem.Delete();
                            itemChanged = true;
                        }
                        else
                        {
                            if (projectItem.ProjectItems.Count <= 0)
                                continue;

                            bool subItemChanged = RemoveDeletedItems(extractedFolder + "\\" + name,
                                projectItem.ProjectItems);
                            if (subItemChanged)
                                itemChanged = true;
                        }
                        break;
                }
            }

            return itemChanged;
        }

        private static bool FileEquals(string path1, string path2)
        {
            FileInfo first = new FileInfo(path1);
            FileInfo second = new FileInfo(path2);

            if (first.Length != second.Length)
                return false;

            int iterations = (int)Math.Ceiling((double)first.Length / sizeof(Int64));

            using (FileStream fs1 = first.OpenRead())
            using (FileStream fs2 = second.OpenRead())
            {
                byte[] one = new byte[sizeof(Int64)];
                byte[] two = new byte[sizeof(Int64)];

                for (int i = 0; i < iterations; i++)
                {
                    fs1.Read(one, 0, sizeof(Int64));
                    fs2.Read(two, 0, sizeof(Int64));

                    if (BitConverter.ToInt64(one, 0) != BitConverter.ToInt64(two, 0))
                        return false;
                }
            }

            return true;
        }

        private async void Package_OnClick(object sender, RoutedEventArgs e)
        {
            CrmSolution selectedSolution = (CrmSolution)SolutionToPackage.SelectedItem;
            if (selectedSolution == null || selectedSolution.SolutionId == Guid.Empty)
                return;

            string solutionXmlPath = Path.GetDirectoryName(ConnPane.SelectedProject.FullName) + "\\Other\\Solution.xml";
            if (!File.Exists(solutionXmlPath))
            {
                MessageBox.Show("Solution.xml does not exist at: " +
                                Path.GetDirectoryName(ConnPane.SelectedProject.FullName) + "\\Other");
                return;
            }

            XmlDocument doc = new XmlDocument();
            doc.Load(solutionXmlPath);

            XmlNodeList versionNodes = doc.GetElementsByTagName("Version");
            if (versionNodes.Count != 1)
            {
                MessageBox.Show("Invalid Solutions.xml: could not locate Versions node");
                return;
            }

            Version version;
            bool validVersion = Version.TryParse(versionNodes[0].InnerText, out version);
            if (!validVersion)
            {
                MessageBox.Show("Invalid Solutions.xml: invalid version");
                return;
            }

            var selectedProject = ConnPane.SelectedProject;
            if (selectedProject == null) return;

            string savePath;
            var props = _dte.Properties["CRM Developer Extensions", "Solution Packager"];
            bool saveSolutionFiles = (bool)props.Item("SaveSolutionFiles").Value;
            if (saveSolutionFiles)
                savePath = Path.GetDirectoryName(selectedProject.FullName) + "\\_Solutions";
            else
            {
                FolderBrowserDialog folderDialog = new FolderBrowserDialog();
                DialogResult result = folderDialog.ShowDialog();
                if (result == DialogResult.OK)
                    savePath = folderDialog.SelectedPath;
                else
                    return;
            }

            if (File.Exists(savePath + "\\" + selectedSolution.UniqueName + "_" +
                            FormatVersionString(version) + ".zip"))
            {
                MessageBoxResult result = MessageBox.Show("Overwrite solution version: " + version + "?", "Overwrite?", MessageBoxButton.YesNo);
                if (result != MessageBoxResult.Yes)
                    return;
            }

            await Task.Run(() => CreatePackage(selectedSolution, version, savePath, selectedProject));
        }

        private void CreatePackage(CrmSolution selectedSolution, Version version, string savePath, Project project)
        {
            try
            {
                //https://msdn.microsoft.com/en-us/library/jj602987.aspx#arguments

                CommandWindow cw = _dte2.ToolWindows.CommandWindow;

                var props = _dte.Properties["CRM Developer Extensions", "Solution Packager"];
                string spPath = (string)props.Item("SolutionPackagerPath").Value;

                if (string.IsNullOrEmpty(spPath))
                {
                    MessageBox.Show("Set SDK bin folder path under Tools -> Options -> CRM Developer Extensions");
                    return;
                }

                if (!spPath.EndsWith("\\"))
                    spPath += "\\";

                string toolPath = @"""" + spPath + "SolutionPackager.exe" + @"""";
                string filename = selectedSolution.UniqueName + "_" +
                                  FormatVersionString(version) + ".zip";

                string command = toolPath + " /action: Pack";
                command += " /zipfile:" + "\"" + savePath + "\\" + filename + "\"";
                command += " /folder: " + "\"" + Path.GetDirectoryName(ConnPane.SelectedProject.FullName) + "\"";

                cw.SendInput("shell " + command, true);

                AddNewSolutionToProject(savePath, project, filename);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error launching Solution Packager: " + Environment.NewLine + Environment.NewLine + ex.Message);
            }
        }

        private void AddNewSolutionToProject(string savePath, Project project, string filename)
        {
            var props = _dte.Properties["CRM Developer Extensions", "Solution Packager"];
            bool saveSolutionFiles = (bool)props.Item("SaveSolutionFiles").Value;

            if (!saveSolutionFiles)
                return;

            //TODO: Adjust for larger solutions?
            System.Threading.Thread.Sleep(2000);

            project.ProjectItems.AddFromFile(savePath + "\\" + filename);
        }

        private string GetSolutionFromCrm(string connString, CrmSolution selectedSolution, bool managed)
        {
            try
            {
                CrmConnection connection = CrmConnection.Parse(connString);

                using (_orgService = new OrganizationService(connection))
                {
                    ExportSolutionRequest request = new ExportSolutionRequest
                    {
                        Managed = managed,
                        SolutionName = selectedSolution.UniqueName
                    };

                    ExportSolutionResponse response = (ExportSolutionResponse)_orgService.Execute(request);

                    var tempFolder = Path.GetTempPath();
                    string fileName = Path.GetFileName(selectedSolution.UniqueName + "_" +
                        FormatVersionString(selectedSolution.Version) + ((managed) ? "_managed" : String.Empty) + ".zip");
                    var tempFile = Path.Combine(tempFolder, fileName);
                    if (File.Exists(tempFile))
                        File.Delete(tempFile);
                    File.WriteAllBytes(tempFile, response.ExportSolutionFile);

                    return tempFile;
                }
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                _logger.WriteToOutputWindow("Error Retrieving Solution From CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, Logger.MessageType.Error);
                return null;
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Retrieving Solution From CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
                return null;
            }
        }

        private string FormatVersionString(Version version)
        {
            string result = version.ToString().Replace(".", "_");

            return result;
        }

        private void SolutionToPackage_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedProject = ConnPane.SelectedProject;
            if (selectedProject == null) return;

            CrmSolution solution = (CrmSolution)SolutionToPackage.SelectedItem;
            if (solution == null || solution.SolutionId == Guid.Empty)
            {
                Package.IsEnabled = false;
                Unpackage.IsEnabled = false;
                SetDownloadManagedEnabled(false);
                DownloadManaged.IsChecked = false;
                return;
            }

            solution.BoundProject = selectedProject.Name;
            AddOrUpdateMapping(solution);

            Package.IsEnabled = solution.SolutionId != Guid.Empty;
            Unpackage.IsEnabled = solution.SolutionId != Guid.Empty;
            SetDownloadManagedEnabled(solution.SolutionId != Guid.Empty);
            DownloadManaged.IsChecked = solution.DownloadManagedSolution;
        }

        private void AddOrUpdateMapping(CrmSolution solution)
        {
            try
            {
                var projectPath = Path.GetDirectoryName(ConnPane.SelectedProject.FullName);
                if (!ConfigFileExists(ConnPane.SelectedProject))
                {
                    _logger.WriteToOutputWindow("Error Updating Mappings In Config File: Missing CRMDeveloperExtensions.config File", Logger.MessageType.Error);
                    return;
                }

                XmlDocument doc = new XmlDocument();
                doc.Load(projectPath + "\\CRMDeveloperExtensions.config");

                //Update or delete existing mapping
                XmlNodeList solutionNodes = doc.GetElementsByTagName("Solution");
                if (solutionNodes.Count > 0)
                {
                    foreach (XmlNode node in solutionNodes)
                    {
                        XmlNode orgId = node["OrgId"];
                        if (orgId != null && orgId.InnerText.ToUpper() != ConnPane.SelectedConnection.OrgId.ToUpper()) continue;

                        XmlNode projectNameNode = node["ProjectName"];
                        if (projectNameNode != null && projectNameNode.InnerText.ToUpper() != solution.BoundProject.ToUpper())
                            continue;

                        if (string.IsNullOrEmpty(solution.BoundProject) || solution.SolutionId == Guid.Empty)
                        {
                            //Delete
                            var parentNode = node.ParentNode;
                            if (parentNode != null)
                                parentNode.RemoveChild(node);
                        }
                        else
                        {
                            //Update
                            XmlNode solutionIdNode = node["SolutionId"];
                            if (solutionIdNode != null)
                                solutionIdNode.InnerText = solution.SolutionId.ToString();
                            XmlNode downloadManagedNode = node["DownloadManaged"];
                            if (downloadManagedNode != null)
                                downloadManagedNode.InnerText = (DownloadManaged.IsChecked == true) ? "true" : "false";
                        }

                        doc.Save(projectPath + "\\CRMDeveloperExtensions.config");
                        return;
                    }
                }

                //Create new mapping
                XmlNodeList projects = doc.GetElementsByTagName("Solutions");
                if (projects.Count > 0)
                {
                    XmlNode solutionNode = doc.CreateElement("Solution");
                    XmlNode org = doc.CreateElement("OrgId");
                    org.InnerText = ConnPane.SelectedConnection.OrgId;
                    solutionNode.AppendChild(org);
                    XmlNode projectNameNode2 = doc.CreateElement("ProjectName");
                    projectNameNode2.InnerText = solution.BoundProject;
                    solutionNode.AppendChild(projectNameNode2);
                    XmlNode solutionId = doc.CreateElement("SolutionId");
                    solutionId.InnerText = solution.SolutionId.ToString();
                    solutionNode.AppendChild(solutionId);
                    XmlNode downloadManaged = doc.CreateElement("DownloadManaged");
                    downloadManaged.InnerText = (DownloadManaged.IsChecked == true) ? "true" : "false";
                    solutionNode.AppendChild(downloadManaged);
                    projects[0].AppendChild(solutionNode);

                    doc.Save(projectPath + "\\CRMDeveloperExtensions.config");
                }
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Updating Mappings In Config File: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
            }
        }

        private void DownloadManaged_Checked(object sender, RoutedEventArgs e)
        {
            var selectedProject = ConnPane.SelectedProject;
            if (selectedProject == null) return;

            CrmSolution solution = (CrmSolution)SolutionToPackage.SelectedItem;
            if (solution != null && solution.SolutionId != Guid.Empty)
                AddOrUpdateMapping(solution);
        }
    }
}