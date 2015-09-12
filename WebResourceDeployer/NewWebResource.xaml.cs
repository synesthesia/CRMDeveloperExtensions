using EnvDTE;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using OutputLogger;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using WebResourceDeployer.Models;

namespace WebResourceDeployer
{
    public partial class NewWebResource
    {
        private static OrganizationService _orgService;
        private readonly CrmConn _connection;
        private readonly Project _project;
        private readonly Logger _logger;

        public Guid NewId;
        public int NewType;
        public string NewName;
        public string NewDisplayName;
        public string NewBoundFile;
        public Guid NewSolutionId;

        public NewWebResource(CrmConn connection, Project project, ObservableCollection<ComboBoxItem> projectFiles)
        {
            InitializeComponent();

            _logger = new Logger();

            _connection = connection;
            _project = project;

            bool result = GetSolutions();

            if (!result)
            {
                MessageBox.Show("Error Retrieving Solutions From CRM. See the Output Window for additional details.");
                DialogResult = false;
                Close();
            }

            Files.ItemsSource = projectFiles;
        }

        private bool GetSolutions()
        {
            try
            {
                List<CrmSolution> solutions = new List<CrmSolution>();

                CrmConnection connection = CrmConnection.Parse(_connection.ConnectionString);
                using (_orgService = new OrganizationService(connection))
                {
                    QueryExpression query = new QueryExpression
                    {
                        EntityName = "solution",
                        ColumnSet = new ColumnSet("friendlyname", "solutionid", "uniquename"),
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

                    EntityCollection results = _orgService.RetrieveMultiple(query);


                    foreach (Entity entity in results.Entities)
                    {
                        CrmSolution solution = new CrmSolution
                        {
                            SolutionId = entity.Id,
                            Name = entity.GetAttributeValue<string>("friendlyname"),
                            Prefix = entity.GetAttributeValue<AliasedValue>("publisher.customizationprefix").Value.ToString(),
                            UniqueName = entity.GetAttributeValue<string>("uniquename")
                        };

                        solutions.Add(solution);
                    }
                }

                //Default on top
                var i = solutions.FindIndex(s => s.SolutionId == new Guid("FD140AAF-4DF4-11DD-BD17-0019B9312238"));
                var item = solutions[i];
                solutions.RemoveAt(i);
                solutions.Insert(0, item);

                Solutions.ItemsSource = solutions;

                return true;
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                _logger.WriteToOutputWindow("Error Retrieving Solutions From CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, Logger.MessageType.Error);
                return false;
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Retrieving Solutions From CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
                return false;
            }
        }

        private async void Create_Click(object sender, RoutedEventArgs e)
        {
            if (Solutions.SelectedItem == null || Files.SelectedItem == null ||
                string.IsNullOrEmpty(Name.Text) || Type.SelectedItem == null)
                return;

            string relativePath = ((ComboBoxItem)Files.SelectedItem).Content.ToString();
            string filePath = Path.GetDirectoryName(_project.FullName) + relativePath.Replace("/", "\\");
            if (!File.Exists(filePath))
            {
                _logger.WriteToOutputWindow("Missing File: " + filePath, Logger.MessageType.Error);
                MessageBox.Show("File does not exist");
                return;
            }

            CrmSolution solution = (CrmSolution)Solutions.SelectedItem;
            int type = Convert.ToInt32(((ComboBoxItem)Type.SelectedItem).Tag.ToString());
            string prefix = Prefix.Text;
            string name = Name.Text;
            string displayName = DisplayName.Text;

            LockOverlay.Visibility = Visibility.Visible;

            bool result = await System.Threading.Tasks.Task.Run(() => CreateWebResource(solution.SolutionId, solution.UniqueName, type, prefix, name, displayName, filePath, relativePath));

            LockOverlay.Visibility = Visibility.Hidden;

            if (!result)
            {
                MessageBox.Show("Error Creating Web Resource. See the Output Window for additional details.");
                return;
            }

            DialogResult = true;
            Close();
        }

        private bool CreateWebResource(Guid solutionId, string uniqueName, int type, string prefix, string name, string displayName, string filePath, string relativePath)
        {
            try
            {
                CrmConnection connection = CrmConnection.Parse(_connection.ConnectionString);

                using (_orgService = new OrganizationService(connection))
                {
                    Entity webResource = new Entity("webresource");
                    webResource["name"] = prefix + name;
                    webResource["webresourcetype"] = new OptionSetValue(type);
                    if (!string.IsNullOrEmpty(displayName))
                        webResource["displayname"] = displayName;

                    if (type == 8)
                        webResource["silverlightversion"] = "4.0";

                    string extension = Path.GetExtension(filePath);
                    string content = extension != null && (extension.ToUpper() != ".TS")
                        ? File.ReadAllText(filePath)
                        : File.ReadAllText(Path.ChangeExtension(filePath, ".js"));
                    webResource["content"] = EncodeString(content);

                    Guid id = _orgService.Create(webResource);

                    _logger.WriteToOutputWindow("New Web Resource Created: " + id, Logger.MessageType.Info);

                    //Add to the choosen solution (not default)
                    if (solutionId != new Guid("FD140AAF-4DF4-11DD-BD17-0019B9312238"))
                    {
                        AddSolutionComponentRequest scRequest = new AddSolutionComponentRequest
                        {
                            ComponentType = 61,
                            SolutionUniqueName = uniqueName,
                            ComponentId = id
                        };
                        AddSolutionComponentResponse response =
                            (AddSolutionComponentResponse)_orgService.Execute(scRequest);

                        _logger.WriteToOutputWindow("New Web Resource Added To Solution: " + response.id, Logger.MessageType.Info);
                    }

                    NewId = id;
                    NewType = type;
                    NewName = prefix + name;
                    if (!string.IsNullOrEmpty(displayName))
                        NewDisplayName = displayName;
                    NewBoundFile = relativePath;
                    NewSolutionId = solutionId;

                    return true;
                }
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                _logger.WriteToOutputWindow("Error Creating Web Resource: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, Logger.MessageType.Error);
                return false;
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Creating Web Resource: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
                return false;
            }
        }

        private void Solutions_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Solutions.SelectedItem != null)
            {
                CrmSolution solution = (CrmSolution)Solutions.SelectedItem;
                SolutionsLabel.Foreground = Brushes.Black;
                Prefix.Text = solution.Prefix + "_";
                Files.IsEnabled = true;
                Name.IsEnabled = true;
                DisplayName.IsEnabled = true;
                Type.IsEnabled = true;
            }
            else
            {
                SolutionsLabel.Foreground = Brushes.Red;
                Prefix.Text = "new_";
                Files.IsEnabled = false;
                Name.IsEnabled = false;
                DisplayName.IsEnabled = false;
                Type.IsEnabled = false;
            }

            Name.Text = null;
            DisplayName.Text = null;
            Type.SelectedIndex = -1;
            Files.SelectedIndex = -1;
        }

        private void Files_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Files.SelectedItem != null)
            {
                FilesLabel.Foreground = Brushes.Black;
                string fileName = ((ComboBoxItem)Files.SelectedItem).Content.ToString();

                if (fileName.Count(s => s == '/') == 1) //not nested in a folder
                {
                    fileName = fileName.Replace("/", String.Empty);
                    DisplayName.Text = fileName;
                    if (fileName.ToUpper().StartsWith(Prefix.Text.ToUpper()))
                    {
                        fileName = fileName.Substring(Prefix.Text.Length, fileName.Length - Prefix.Text.Length);
                    }
                }
                else
                    DisplayName.Text = fileName;

                string ex = Path.GetExtension(fileName);

                if (ex.ToUpper() != ".TS")
                    Name.Text = fileName;
                else
                {
                    DisplayName.Text = DisplayName.Text.Substring(0, DisplayName.Text.Length - 3) + ".js";
                    Name.Text = fileName.Substring(0, fileName.Length - 3) + ".js"; ;
                }

                if (string.IsNullOrEmpty(ex))
                {
                    Type.SelectedValue = null;
                    return;
                }

                switch (ex.ToUpper())
                {
                    case ".HTM":
                    case ".HTML":
                        Type.SelectedValue = "Webpage (HTML)";
                        break;
                    case ".CSS":
                        Type.SelectedValue = "Style Sheet (CSS)";
                        break;
                    case ".JS":
                    case ".TS":
                        Type.SelectedValue = "Script (JScript)";
                        break;
                    case ".XML":
                        Type.SelectedValue = "Data (XML)";
                        break;
                    case ".PNG":
                        Type.SelectedValue = "PNG format";
                        break;
                    case ".JPG":
                        Type.SelectedValue = "JPG format";
                        break;
                    case ".GIF":
                        Type.SelectedValue = "GIF format";
                        break;
                    case ".XAP":
                        Type.SelectedValue = "Silverlight (XAP)";
                        break;
                    case ".XSL":
                    case ".XSLT":
                        Type.SelectedValue = "Style Sheet (XSL)";
                        break;
                    case ".ICO":
                        Type.SelectedValue = "ICO format";
                        break;
                }
            }
            else
            {
                FilesLabel.Foreground = Brushes.Red;
                Name.Text = null;
                DisplayName.Text = null;
            }
        }

        private void Name_TextChanged(object sender, TextChangedEventArgs e)
        {
            NameLabel.Foreground = string.IsNullOrEmpty(Name.Text) ? Brushes.Red : Brushes.Black;
        }

        private void Type_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TypeLabel.Foreground = Type.SelectedItem != null ? Brushes.Black : Brushes.Red;
        }

        private string EncodeString(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
        }
    }
}
