using CommonResources.Models;
using CrmConnectionWindow;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using OutputLogger;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Xml;
using Window = EnvDTE.Window;

namespace CommonResources
{
    public partial class ConnectionPane
    {
        private readonly DTE _dte;
        private readonly Solution _solution;
        private readonly Logger _logger;
        private bool _connectionAdded;

        const string SolutionFolder = "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}";

        public event EventHandler ProjectChanged;
        public event EventHandler<ConnectionSelectedEventArgs> ConnectionSelected;
        public event EventHandler<ConnectionModifiedEventArgs> ConnectionModified;
        public event EventHandler<ConnectionAddedEventArgs> ConnectionAdded;
        public event EventHandler ConnectionDeleted;
        public event EventHandler<ConnectEventArgs> Connected;
        public event EventHandler ConnectionStarted;

        public Project SelectedProject { get; private set; }
        public CrmConn SelectedConnection { get; private set; }
        public Projects Projects { get; private set; }
        public string SourceWindow { get; set; }

        public ConnectionPane()
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
            solutionEvents.BeforeClosing += BeforeSolutionClosing;
            solutionEvents.ProjectAdded += SolutionProjectAdded;
            solutionEvents.Opened += SolutionEventsOnOpened;
            solutionEvents.ProjectRemoved += SolutionProjectRemoved;
            solutionEvents.ProjectRenamed += SolutionProjectRenamed;
        }

        private void SolutionEventsOnOpened()
        {
            Projects = _dte.Solution.Projects;
        }

        private void SolutionProjectRenamed(Project project, string oldName)
        {
            string name = Path.GetFileNameWithoutExtension(oldName);
            foreach (ComboBoxItem comboBoxItem in ProjectsDdl.Items)
            {
                if (string.IsNullOrEmpty(comboBoxItem.Content.ToString())) continue;
                if (name != null && comboBoxItem.Content.ToString().ToUpper() != name.ToUpper()) continue;

                comboBoxItem.Content = project.Name;
            }

            Projects = _dte.Solution.Projects;
        }

        private void SolutionProjectRemoved(Project project)
        {
            foreach (ComboBoxItem comboBoxItem in ProjectsDdl.Items)
            {
                if (string.IsNullOrEmpty(comboBoxItem.Content.ToString())) continue;
                if (comboBoxItem.Content.ToString().ToUpper() != project.Name.ToUpper()) continue;

                ProjectsDdl.Items.Remove(comboBoxItem);
                break;
            }

            if (SelectedProject != null)
            {
                if (SelectedProject.FullName == project.FullName)
                {
                    Connections.ItemsSource = null;
                    Connections.Items.Clear();
                    Connections.IsEnabled = false;
                    AddConnection.IsEnabled = false;
                }
            }

            Projects = _dte.Solution.Projects;
        }

        private void BeforeSolutionClosing()
        {
            ResetForm();
            Expander.IsExpanded = true;
        }

        private void WindowEventsOnWindowActivated(Window gotFocus, Window lostFocus)
        {
            if (Projects == null)
                Projects = _dte.Solution.Projects;

            //No solution loaded
            if (_solution.Count == 0)
            {
                ResetForm();
                return;
            }

            //WindowEventsOnWindowActivated in this project can be called when activating another window
            //so we don't want to contine further unless our window is active
            if (!gotFocus.Caption.StartsWith("CRM DevEx")) return;

            ProjectsDdl.IsEnabled = true;
            AddConnection.IsEnabled = true;
            Connections.IsEnabled = true;

            foreach (var project in GetProjects())
            {
                SolutionProjectAdded(project);
            }
        }

        private IEnumerable<Project> GetProjects()
        {
            var list = new List<Project>();
            var item = _dte.Solution.Projects.GetEnumerator();

            while (item.MoveNext())
            {
                var project = item.Current as Project;

                if (project == null) continue;

                if (IsUnitTestProject(project)) continue;

                if (project.Kind.ToUpper() == SolutionFolder)
                    list.AddRange(GetFolderProjects(project));
                else
                    list.Add(project);
            }

            return list;
        }

        private static bool IsUnitTestProject(Project project)
        {
            bool isUnitTestProject = false;

            if (string.IsNullOrEmpty(project.FullName)) return true;

            var settingsPath = Path.GetDirectoryName(project.FullName);
            if (File.Exists(settingsPath + "\\Properties\\settings.settings"))
            {
                XmlDocument settingsDoc = new XmlDocument();
                settingsDoc.Load(settingsPath + "\\Properties\\settings.settings");

                XmlNodeList settings = settingsDoc.GetElementsByTagName("Settings");
                if (settings.Count > 0)
                {
                    XmlNodeList appSettings = settings[0].ChildNodes;
                    foreach (XmlNode node in appSettings)
                    {
                        if (node.Attributes == null || node.Attributes["Name"] == null) continue;
                        if (node.Attributes["Name"].Value != "CRMTestType") continue;

                        XmlNode value = node.FirstChild;
                        if (string.IsNullOrEmpty(value.InnerText)) continue;

                        isUnitTestProject = true;
                        break;
                    }
                }
            }

            return isUnitTestProject;
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

        private void SolutionProjectAdded(Project project)
        {
            //Don't want to include the VS Miscellaneous Files Project - which appears occasionally and during a diff operation
            if (project.Name.ToUpper() == "MISCELLANEOUS FILES")
                return;

            // Exclude solution folders 
            if (project.Kind != null && project.Kind.ToUpper() == SolutionFolder)
                return;

            bool addProject = true;
            foreach (ComboBoxItem projectItem in ProjectsDdl.Items)
            {
                if (projectItem.Content.ToString().ToUpper() != project.Name.ToUpper())
                    continue;

                addProject = false;
                break;
            }

            if (addProject)
            {
                ComboBoxItem item = new ComboBoxItem { Content = project.Name, Tag = project };
                ProjectsDdl.Items.Add(item);
            }

            if (ProjectsDdl.SelectedIndex == -1)
                ProjectsDdl.SelectedIndex = 0;

            Projects = _dte.Solution.Projects;
        }

        private void ResetForm()
        {
            Connections.ItemsSource = null;
            Connections.Items.Clear();
            Connections.IsEnabled = false;
            ProjectsDdl.ItemsSource = null;
            ProjectsDdl.Items.Clear();
            ProjectsDdl.IsEnabled = false;
            AddConnection.IsEnabled = false;
        }

        private void Projects_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //No solution loaded
            if (_solution.Count == 0) return;

            ComboBoxItem item = (ComboBoxItem)ProjectsDdl.SelectedItem;
            if (item == null) return;
            if (string.IsNullOrEmpty(item.Content.ToString()))
            {
                SelectedProject = null;
                OnProjectChanged();
                return;
            }

            SelectedProject = (Project)((ComboBoxItem)ProjectsDdl.SelectedItem).Tag;
            GetConnections();

            OnProjectChanged();
        }

        private void GetConnections()
        {
            Connections.ItemsSource = null;

            var path = Path.GetDirectoryName(SelectedProject.FullName);
            XmlDocument doc = new XmlDocument();

            if (!SharedConfigFile.ConfigFileExists(SelectedProject))
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

        private string DecodeString(string value)
        {
            byte[] data = Convert.FromBase64String(value);
            return Encoding.UTF8.GetString(data);
        }

        private void Connections_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SelectedConnection = (CrmConn)Connections.SelectedItem;
            if (SelectedConnection != null)
            {
                Connect.IsEnabled = !string.IsNullOrEmpty(SelectedConnection.Name);
                Delete.IsEnabled = !string.IsNullOrEmpty(SelectedConnection.Name);
                ModifyConnection.IsEnabled = !string.IsNullOrEmpty(SelectedConnection.Name);

                if (_connectionAdded)
                    _connectionAdded = false;
            }
            else
            {
                Connect.IsEnabled = false;
                Delete.IsEnabled = false;
                ModifyConnection.IsEnabled = false;
            }

            OnConnectionSelected(new ConnectionSelectedEventArgs
            {
                SelectedConnection = SelectedConnection,
                ConnectionAdded = _connectionAdded
            });

            SharedGlobals.SetGlobal("SelectedConnection", SelectedConnection, _dte);
        }

        private void AddConnection_Click(object sender, RoutedEventArgs e)
        {
            var connection = new Connection(null, null, SourceWindow, _dte);

            bool? result = connection.ShowDialog();

            if (!result.HasValue || !result.Value) return;

            var configExists = SharedConfigFile.ConfigFileExists(SelectedProject);
            if (!configExists)
                CreateConfigFile(SelectedProject);

            Expander.IsExpanded = false;

            AddOrUpdateConnection(SelectedProject, connection.ConnectionName, connection.ConnectionString, connection.OrgId, connection.Version, true);

            GetConnections();

            foreach (CrmConn conn in Connections.Items)
            {
                if (conn.Name != connection.ConnectionName) continue;

                Connections.SelectedItem = conn;
                OnConnectionAdded(new ConnectionAddedEventArgs
                {
                    AddedConnection = conn
                });

                break;
            }
        }

        private void AddOrUpdateConnection(Project vsProject, string connectionName, string connString, string orgId, string versionNum, bool showPrompt)
        {
            try
            {
                var path = Path.GetDirectoryName(vsProject.FullName);
                if (!SharedConfigFile.ConfigFileExists(vsProject))
                {
                    _logger.WriteToOutputWindow("Error Adding Or Updating Connection: Missing CRMDeveloperExtensions.config File", Logger.MessageType.Error);
                    return;
                }

                FileInfo file = new FileInfo(path + "\\CRMDeveloperExtensions.config");

                XmlDocument doc = new XmlDocument();
                doc.Load(path + "\\CRMDeveloperExtensions.config");

                //Check if connection already exists for project
                XmlNodeList connectionNodes = doc.GetElementsByTagName("Connection");
                if (connectionNodes.Count > 0)
                {
                    foreach (XmlNode node in connectionNodes)
                    {
                        XmlNode nameNode = node["Name"];
                        if (nameNode != null && nameNode.InnerText != connectionName) continue;

                        if (showPrompt)
                        {
                            MessageBoxResult result = MessageBox.Show("Update Connection?", "Connection Already Added",
                                MessageBoxButton.YesNo);

                            //Update existing connection
                            if (result != MessageBoxResult.Yes)
                                return;
                        }

                        bool changed = false;
                        if (nameNode != null)
                        {
                            string oldConnectionName = nameNode.InnerText;
                            if (oldConnectionName != connectionName)
                            {
                                nameNode.InnerText = connectionName;
                                changed = true;
                            }
                        }
                        XmlNode versionNode = node["Version"];
                        if (versionNode != null)
                        {
                            string oldVersionNum = versionNode.InnerText;
                            if (oldVersionNum != versionNum)
                            {
                                versionNode.InnerText = versionNum;
                                changed = true;
                            }
                        }
                        XmlNode connectionStringNode = node["ConnectionString"];
                        if (connectionStringNode != null)
                        {
                            string oldConnectionString = connectionStringNode.InnerText;
                            string encodedConnectionString = EncodeString(connString);
                            if (oldConnectionString != encodedConnectionString)
                            {
                                connectionStringNode.InnerText = encodedConnectionString;
                                changed = true;
                            }
                        }

                        if (!changed) return;

                        if (SharedConfigFile.IsConfigReadOnly(path + "\\CRMDeveloperExtensions.config"))
                            file.IsReadOnly = false;

                        doc.Save(path + "\\CRMDeveloperExtensions.config");

                        return;
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
                connections[0].AppendChild(connection);

                _connectionAdded = true;

                if (SharedConfigFile.IsConfigReadOnly(path + "\\CRMDeveloperExtensions.config"))
                    file.IsReadOnly = false;

                doc.Save(path + "\\CRMDeveloperExtensions.config");
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Adding Or Updating Connection: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
            }
        }

        private string EncodeString(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
        }

        private void CreateConfigFile(Project vsProject)
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                XmlElement root = doc.CreateElement(SourceWindow);
                doc.AppendChild(root);
                XmlElement connections = doc.CreateElement("Connections");
                root.AppendChild(connections);

                //Keeeping the different root element names as not not introduce a breaking change
                //and figuring there should also only be 1 deployer type in use per VS project 
                switch (SourceWindow)
                {
                    case "PluginDeployer":
                        XmlElement assemblies = doc.CreateElement("Assemblies");
                        root.AppendChild(assemblies);
                        break;
                    case "WebResourceDeployer":
                    case "ReportDeployer":
                        XmlElement files = doc.CreateElement("Files");
                        root.AppendChild(files);
                        break;
                    case "SolutionPackager":
                        XmlElement solution = doc.CreateElement("Solutions");
                        root.AppendChild(solution);
                        break;
                }

                var path = Path.GetDirectoryName(vsProject.FullName);
                doc.Save(path + "/CRMDeveloperExtensions.config");
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Creating Config File: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
            }
        }

        private void ModifyConnection_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedConnection == null) return;
            if (string.IsNullOrEmpty(SelectedConnection.ConnectionString)) return;

            string name = SelectedConnection.Name;
            Connection connection = new Connection(name, SelectedConnection.ConnectionString, SourceWindow, _dte);
            bool? result = connection.ShowDialog();

            if (!result.HasValue || !result.Value) return;

            var configExists = SharedConfigFile.ConfigFileExists(SelectedProject);
            if (!configExists)
                CreateConfigFile(SelectedProject);

            Expander.IsExpanded = false;

            AddOrUpdateConnection(SelectedProject, connection.ConnectionName, connection.ConnectionString, connection.OrgId, connection.Version, false);

            //Keep to refresh the connections in the list
            GetConnections();

            foreach (CrmConn conn in Connections.Items)
            {
                if (conn.Name != connection.ConnectionName) continue;

                Connections.SelectedItem = conn;
                OnConnectionModified(new ConnectionModifiedEventArgs
                {
                    ModifiedConnection = conn
                });
                break;
            }
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (SelectedConnection == null) return;
                if (string.IsNullOrEmpty(SelectedConnection.ConnectionString)) return;

                var path = Path.GetDirectoryName(SelectedProject.FullName);
                if (!SharedConfigFile.ConfigFileExists(SelectedProject))
                {
                    _logger.WriteToOutputWindow("Error Deleting Connection: Missing CRMDeveloperExtensions.config File", Logger.MessageType.Error);
                    return;
                }

                MessageBoxResult result = MessageBox.Show("Are you sure?" + Environment.NewLine + Environment.NewLine +
                    "This will delete the connection information and all associated mappings.", "Delete Connection", MessageBoxButton.YesNo);
                if (result != MessageBoxResult.Yes) return;

                if (!SharedConfigFile.ConfigFileExists(SelectedProject)) return;

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
                    if (orgId.InnerText.ToUpper() != SelectedConnection.OrgId.ToUpper()) continue;

                    nodesToRemove.Add(connection);
                }

                foreach (XmlNode xmlNode in nodesToRemove)
                {
                    if (xmlNode.ParentNode != null)
                        xmlNode.ParentNode.RemoveChild(xmlNode);
                }

                if (SharedConfigFile.IsConfigReadOnly(path + "\\CRMDeveloperExtensions.config"))
                {
                    FileInfo file = new FileInfo(path + "\\CRMDeveloperExtensions.config") { IsReadOnly = false };
                }

                doc.Save(path + "\\CRMDeveloperExtensions.config");

                //Delete related Files
                doc.Load(path + "\\CRMDeveloperExtensions.config");

                string type;
                switch (SourceWindow)
                {
                    case "PluginDeployer":
                        type = "Assembly";
                        break;
                    case "SolutionPackager":
                        type = "Solution";
                        break;
                    default: //WebResourceDeployer or ReportDeployer
                        type = "File";
                        break;
                }

                XmlNodeList mapping = doc.GetElementsByTagName(type);
                if (mapping.Count > 0)
                {
                    nodesToRemove = new List<XmlNode>();
                    foreach (XmlNode fileNode in mapping)
                    {
                        XmlNode orgId = fileNode["OrgId"];
                        if (orgId == null) continue;
                        if (orgId.InnerText.ToUpper() != SelectedConnection.OrgId.ToUpper()) continue;

                        nodesToRemove.Add(fileNode);
                    }

                    foreach (XmlNode xmlNode in nodesToRemove)
                    {
                        if (xmlNode.ParentNode != null)
                            xmlNode.ParentNode.RemoveChild(xmlNode);
                    }
                    doc.Save(path + "\\CRMDeveloperExtensions.config");
                }

                GetConnections();

                OnConnectionDeleted();
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Deleting Connection: Missing CRMDeveloperExtensions.config File: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
            }
        }

        private void Connect_Click(object sender, RoutedEventArgs e)
        {
            OnConnectionStarted();

            if (SelectedConnection == null) return;

            string connString = SelectedConnection.ConnectionString;
            if (string.IsNullOrEmpty(connString)) return;

            bool isValid = ValidateXrmToolingConnString(connString);
            if (!isValid) return;

            Expander.IsExpanded = false;

            OnConnected(new ConnectEventArgs
            {
                ConnectionString = connString
            });

            AddOrUpdateConnection(SelectedProject, SelectedConnection.Name, SelectedConnection.ConnectionString, SelectedConnection.OrgId, SelectedConnection.Version, false);
        }

        private static bool ValidateXrmToolingConnString(string connString)
        {
            if (connString.ToUpper().Contains("AUTHTYPE="))
                return true;

            MessageBox.Show("You are using an old connection string which does not work with the new Xrm.Tooling connection" + Environment.NewLine + Environment.NewLine +
                "Your options are:" + Environment.NewLine + Environment.NewLine +
                "1. Create a new connection & remap items" + Environment.NewLine +
                "2. Open the CRMDeveloperExtensions.config at the project root and replace the Base64 encoded connection string with " +
                "one created from a new connection or a Base64 encoded version from: https://msdn.microsoft.com/en-us/library/mt608573.aspx",
                "Upgrade Connection String");

            return false;
        }

        protected virtual void OnProjectChanged()
        {
            var handler = ProjectChanged;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        protected virtual void OnConnectionSelected(ConnectionSelectedEventArgs e)
        {
            var handler = ConnectionSelected;
            if (handler != null) handler(this, e);
        }

        protected virtual void OnConnected(ConnectEventArgs e)
        {
            var handler = Connected;
            if (handler != null) handler(this, e);
        }

        protected virtual void OnConnectionAdded(ConnectionAddedEventArgs e)
        {
            var handler = ConnectionAdded;
            if (handler != null) handler(this, e);
        }

        protected virtual void OnConnectionDeleted()
        {
            var handler = ConnectionDeleted;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        protected virtual void OnConnectionModified(ConnectionModifiedEventArgs e)
        {
            var handler = ConnectionModified;
            if (handler != null) handler(this, e);
        }

        protected virtual void OnConnectionStarted()
        {
            var handler = ConnectionStarted;
            if (handler != null) handler(this, EventArgs.Empty);
        }
    }

    public class ConnectEventArgs
    {
        public string ConnectionString { get; set; }
    }

    public class ConnectionSelectedEventArgs
    {
        public CrmConn SelectedConnection { get; set; }

        public bool ConnectionAdded { get; set; }
    }

    public class ConnectionAddedEventArgs
    {
        public CrmConn AddedConnection { get; set; }
    }

    public class ConnectionModifiedEventArgs
    {
        public CrmConn ModifiedConnection { get; set; }
    }
}
