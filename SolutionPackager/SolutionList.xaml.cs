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
using System.Windows;
using System.Windows.Controls;
using System.Xml;
using CommonResources;
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
            solutionEvents.BeforeClosing += BeforeSolutionClosing;
            solutionEvents.BeforeClosing += SolutionBeforeClosing;
            solutionEvents.ProjectRemoved += SolutionProjectRemoved;
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
        }

        private void ConnPane_OnConnected(object sender, ConnectEventArgs e)
        {
            GetSolutions(e.ConnectionString);

            Customizations.IsEnabled = true;
            Solutions.IsEnabled = true;
            Package.IsEnabled = true;
            Unpackage.IsEnabled = true;
        }

        private void ConnPane_OnProjectChanged(object sender, EventArgs e)
        {
            Customizations.IsEnabled = false;
            Solutions.IsEnabled = false;
            Package.IsEnabled = false;
            Unpackage.IsEnabled = false;
            SolutionToPackage.ItemsSource = null;
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

                var props = _dte.Properties["CRM Developer Extensions", "Settings"];
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

            EntityCollection results = await System.Threading.Tasks.Task.Run(() => GetSolutionsFromCrm(connString));
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
                SolutionToPackage.SelectedItem = solutions.First(s => !string.IsNullOrEmpty(s.BoundProject));
            SolutionToPackage.IsEnabled = true;

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

        private void BeforeSolutionClosing()
        {
            ResetForm();
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
        }

        private void SolutionProjectRemoved(Project project)
        {
            if (ConnPane.SelectedProject != null)
            {
                if (ConnPane.SelectedProject.FullName == project.FullName)
                {
                    SolutionToPackage.ItemsSource = null;
                    Package.IsEnabled = false;
                    Unpackage.IsEnabled = false;
                    Customizations.IsEnabled = false;
                    Solutions.IsEnabled = false;
                }
            }
        }

        private async void Unpackage_OnClick(object sender, RoutedEventArgs e)
        {
            _dte.StatusBar.Text = "Connecting to CRM and getting solution...";
            _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationSync);
            LockMessage.Content = "Working...";
            LockOverlay.Visibility = Visibility.Visible;

            if (SolutionToPackage.SelectedItem == null)
                return;

            CrmSolution selectedSolution = (CrmSolution)SolutionToPackage.SelectedItem;

            string path = await System.Threading.Tasks.Task.Run(() => GetSolutionFromCrm(ConnPane.SelectedConnection.ConnectionString, selectedSolution));
            if (string.IsNullOrEmpty(path))
            {
                _dte.StatusBar.Clear();
                LockOverlay.Visibility = Visibility.Hidden;
                MessageBox.Show("Error Retrieving Solution. See the Output Window for additional details.");
                return;
            }

            _logger.WriteToOutputWindow("Retrieved Solution From CRM", Logger.MessageType.Info);
            _dte.StatusBar.Text = "Extracting solution...";

            await System.Threading.Tasks.Task.Run(() => ExtractPackage(path));

            _dte.StatusBar.Clear();
            _dte.StatusBar.Animate(false, vsStatusAnimation.vsStatusAnimationSync);
            LockOverlay.Visibility = Visibility.Hidden;
        }

        private void ExtractPackage(string path)
        {
            CommandWindow cw = _dte2.ToolWindows.CommandWindow;

            //TODO: Make user setting
            string toolPath = @"""C:\Users\jason.lattimer\Documents\SDK\CRM SDK\2015 SDK 7.1.0\SDK\Bin\SolutionPackager.exe""";

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

            RemoveDeletedItems(extractedFolder.FullName, ConnPane.SelectedProject.ProjectItems);
            ProcessDownloadedSolution(extractedFolder, Path.GetDirectoryName(ConnPane.SelectedProject.FullName),
                ConnPane.SelectedProject.ProjectItems);

            Directory.Delete(extractedFolder.FullName, true);
        }

        private void ProcessDownloadedSolution(DirectoryInfo extractedFolder, string baseFolder, ProjectItems projectItems)
        {
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
            }

            //Handle folder adds
            foreach (DirectoryInfo folder in extractedFolder.GetDirectories())
            {
                if (!Directory.Exists(baseFolder + "\\" + folder.Name))
                    Directory.CreateDirectory(baseFolder + "\\" + folder.Name);

                var newProjectItems = projectItems;

                ProcessDownloadedSolution(folder, baseFolder + "\\" + folder.Name, newProjectItems);
            }
        }

        private static void RemoveDeletedItems(string extractedFolder, ProjectItems projectItems)
        {
            //Handle file & folder deletes
            foreach (ProjectItem projectItem in projectItems)
            {
                string name = projectItem.FileNames[0];
                if (projectItem.Kind.ToUpper() == "{6BB5F8EE-4483-11D3-8BCF-00C04F8EC28C}") //File
                {
                    name = Path.GetFileName(name);
                    if (!File.Exists(extractedFolder + "\\" + name))
                        projectItem.Delete();
                }
                else if (projectItem.Kind.ToUpper() == "{6BB5F8EF-4483-11D3-8BCF-00C04F8EC28C}") //Folder
                {
                    name = new DirectoryInfo(name).Name;
                    if (!Directory.Exists(extractedFolder + "\\" + name))
                        projectItem.Delete();
                    else
                    {
                        if (projectItem.ProjectItems.Count > 0)
                            RemoveDeletedItems(extractedFolder + "\\" + name, projectItem.ProjectItems);
                    }
                }
            }
        }

        private static bool FileEquals(string path1, string path2)
        {
            byte[] file1 = File.ReadAllBytes(path1);
            byte[] file2 = File.ReadAllBytes(path2);
            if (file1.Length != file2.Length)
                return false;

            for (int i = 0; i < file1.Length; i++)
            {
                if (file1[i] != file2[i])
                    return false;
            }
            return true;
        }

        private void Package_OnClick(object sender, RoutedEventArgs e)
        {
            CommandWindow cw = _dte2.ToolWindows.CommandWindow;

            //TODO: Make user setting
            string toolPath = @"""C:\Users\jason.lattimer\Documents\SDK\CRM SDK\2015 SDK 7.1.0\SDK\Bin\SolutionPackager.exe""";

            CrmSolution selectedSolution = (CrmSolution)SolutionToPackage.SelectedItem;

            string command = toolPath + " /action: Pack";
            command += " /zipfile:" + "\"" + Path.GetTempPath() + selectedSolution.UniqueName + "_" +
                        FormatVersionString(selectedSolution.Version) + ".zip" + "\"";
            command += " /folder: " + "\"" + Path.GetDirectoryName(ConnPane.SelectedProject.FullName) + "\"";

            cw.SendInput("shell " + command, true);
        }

        private string GetSolutionFromCrm(string connString, CrmSolution selectedSolution)
        {
            try
            {
                CrmConnection connection = CrmConnection.Parse(connString);

                using (_orgService = new OrganizationService(connection))
                {
                    ExportSolutionRequest request = new ExportSolutionRequest
                    {
                        Managed = false,
                        SolutionName = selectedSolution.UniqueName
                    };

                    ExportSolutionResponse response = (ExportSolutionResponse)_orgService.Execute(request);

                    var tempFolder = Path.GetTempPath();
                    string fileName = Path.GetFileName(selectedSolution.UniqueName + "_" +
                        FormatVersionString(selectedSolution.Version) + ".zip");
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
            if (solution == null)
            {
                Package.IsEnabled = false;
                Unpackage.IsEnabled = false;
                return;
            }

            solution.BoundProject = selectedProject.Name;
            AddOrUpdateMapping(solution);

            Package.IsEnabled = solution.SolutionId != Guid.Empty;
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
                    projects[0].AppendChild(solutionNode);

                    doc.Save(projectPath + "\\CRMDeveloperExtensions.config");
                }
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Updating Mappings In Config File: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
            }
        }
    }
}