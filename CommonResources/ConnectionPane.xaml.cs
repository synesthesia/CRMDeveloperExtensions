using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Xml;
using CommonResources.Models;
using CrmConnectionWindow;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using OutputLogger;
using Window = EnvDTE.Window;

namespace CommonResources
{
    /// <summary>
    /// Interaction logic for ConnectionPane.xaml
    /// </summary>
    public partial class ConnectionPane
    {
        private readonly DTE _dte;
        private readonly Solution _solution;
        private readonly Logger _logger;
        private bool _connectionAdded;

        public event EventHandler ProjectChanged;
        public event EventHandler<ConnectionSelectedEventArgs> ConnectionSelected;
        public event EventHandler<ConnectionModifiedEventArgs> ConnectionModified;
        public event EventHandler<ConnectionAddedEventArgs> ConnectionAdded;
        public event EventHandler ConnectionDeleted;
        public event EventHandler<ConnectEventArgs> Connected;

        
        public Project SelectedProject { get; private set; }
        public CrmConn SelectedConnection { get; private set; }
        public Projects Projects { get; private set; }

         

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

            windowEvents.WindowActivated += WindowEventsOnWindowActivated;
            var solutionEvents = events.SolutionEvents;
            solutionEvents.BeforeClosing += BeforeSolutionClosing;
            solutionEvents.ProjectAdded += SolutionProjectAdded;
            solutionEvents.ProjectRemoved += SolutionProjectRemoved;
            solutionEvents.ProjectRenamed += SolutionProjectRenamed;
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
            // TODO: Can this be replaced with LINQ?
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

        private void ConnectionPane_OnUnloaded(object sender, RoutedEventArgs e)
        {
            ResetForm();
        }

        private void BeforeSolutionClosing()
        {
            ResetForm();
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
            
            // TODO: What is the purpose of this line?
            //if (gotFocus.Caption != SolutionPackager.Resources.ResourceManager.GetString("ToolWindowTitle")) return;

            ProjectsDdl.IsEnabled = true;
            AddConnection.IsEnabled = true;
            Connections.IsEnabled = true;

            foreach (Project project in Projects)
            {
                SolutionProjectAdded(project);
            }
        }

        private void SolutionProjectAdded(Project project)
        {
            //Don't want to include the VS Miscellaneous Files Project - which appears occasionally and during a diff operation
            if (project.Name.ToUpper() == "MISCELLANEOUS FILES")
                return;

            bool addProject = true;
            foreach (ComboBoxItem projectItem in ProjectsDdl.Items)
            {
                if (projectItem.Content.ToString().ToUpper() != project.Name.ToUpper()) continue;

                addProject = false;
                break;
            }

            if (addProject)
            {
                ComboBoxItem item = new ComboBoxItem() { Content = project.Name, Tag = project };
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

            if (!ConfigFileExists(SelectedProject))
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

        private bool ConfigFileExists(Project project)
        {
            var path = Path.GetDirectoryName(project.FullName);
            return File.Exists(path + "/CRMDeveloperExtensions.config");
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
                {
                    _connectionAdded = false;
                }
                else
                {
                    // TODO: Is it really necessary to update the selected connection here?
                    UpdateSelectedConnection(false);
                }
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
        }

        private void AddConnection_Click(object sender, RoutedEventArgs e)
        {
            Connection connection = new Connection(null, null);
            bool? result = connection.ShowDialog();

            if (!result.HasValue || !result.Value) return;

            var configExists = ConfigFileExists(SelectedProject);
            if (!configExists)
                CreateConfigFile(SelectedProject);

            Expander.IsExpanded = false;

            bool change = AddOrUpdateConnection(SelectedProject, connection.ConnectionName, connection.ConnectionString, connection.OrgId, connection.Version, true);
            if (!change) return;

            GetConnections();

            // TODO: Can this be done using LINQ?
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

                //Check if connection already exists for project
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

        private string EncodeString(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
        }

        private void CreateConfigFile(Project vsProject)
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                XmlElement pluginDeployer = doc.CreateElement("PluginDeployer");
                XmlElement connections = doc.CreateElement("Connections");
                XmlElement projects = doc.CreateElement("Assemblies");
                pluginDeployer.AppendChild(connections);
                pluginDeployer.AppendChild(projects);
                doc.AppendChild(pluginDeployer);

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
            Connection connection = new Connection(name, SelectedConnection.ConnectionString);
            bool? result = connection.ShowDialog();

            if (!result.HasValue || !result.Value) return;

            var configExists = ConfigFileExists(SelectedProject);
            if (!configExists)
                CreateConfigFile(SelectedProject);

            Expander.IsExpanded = false;

            AddOrUpdateConnection(SelectedProject, connection.ConnectionName, connection.ConnectionString, connection.OrgId, connection.Version, false);

            GetConnections();
            // TODO: Can this be replaced with LINQ?
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
                MessageBoxResult result = MessageBox.Show("Are you sure?" + Environment.NewLine + Environment.NewLine +
                    "This will delete the connection information and all associated mappings.", "Delete Connection", MessageBoxButton.YesNo);
                if (result != MessageBoxResult.Yes) return;

                if (SelectedConnection == null) return;
                if (string.IsNullOrEmpty(SelectedConnection.ConnectionString)) return;

                var path = Path.GetDirectoryName(SelectedProject.FullName);
                if (!ConfigFileExists(SelectedProject))
                {
                    _logger.WriteToOutputWindow("Error Deleting Connection: Missing CRMDeveloperExtensions.config File", Logger.MessageType.Error);
                    return;
                }

                if (!ConfigFileExists(SelectedProject)) return;

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
                        if (orgId.InnerText.ToUpper() != SelectedConnection.OrgId.ToUpper()) continue;

                        nodesToRemove.Add(file);
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
            if (SelectedConnection == null) return;

            string connString = SelectedConnection.ConnectionString;
            if (string.IsNullOrEmpty(connString)) return;

            UpdateSelectedConnection(true);

            Expander.IsExpanded = false;

            OnConnected(new ConnectEventArgs
            {
                ConnectionString = connString
            });
        }

        private void UpdateSelectedConnection(bool makeSelected)
        {
            try
            {
                var path = Path.GetDirectoryName(SelectedProject.FullName);
                if (!ConfigFileExists(SelectedProject))
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
                            selected.InnerText = name.InnerText != SelectedConnection.Name ? "False" : "True";
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
