using EnvDTE;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using OutputLogger;
using ReportDeployer.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.ServiceModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using CommonResources.Models;
using Microsoft.Xrm.Tooling.Connector;

namespace ReportDeployer
{
    public partial class NewReport
    {
        private readonly CrmConn _connection;
        private readonly Project _project;
        private readonly Logger _logger;

        public Guid NewId;
        public string NewName;
        public string NewBoudndFile;

        public NewReport(CrmConn connection, Project project, ObservableCollection<ComboBoxItem> projectFiles)
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

                var client = new CrmServiceClient(_connection.ConnectionString);
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
                    Orders =
                        {
                            new OrderExpression
                            {
                                AttributeName = "friendlyname",
                                OrderType = OrderType.Ascending
                            }
                        }
                };

                EntityCollection results = client.RetrieveMultiple(query);

                foreach (Entity entity in results.Entities)
                {
                    CrmSolution solution = new CrmSolution
                    {
                        SolutionId = entity.Id,
                        Name = entity.GetAttributeValue<string>("friendlyname"),
                        UniqueName = entity.GetAttributeValue<string>("uniquename")
                    };

                    solutions.Add(solution);
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
            if (Solutions.SelectedItem == null || Files.SelectedItem == null || string.IsNullOrEmpty(Name.Text))
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
            string name = Name.Text;
            int viewableIndex = Viewable.SelectedIndex;

            LockOverlay.Visibility = Visibility.Visible;

            bool result = await System.Threading.Tasks.Task.Run(() => CreateReport(solution.SolutionId, solution.UniqueName, name, filePath, relativePath, viewableIndex));

            LockOverlay.Visibility = Visibility.Hidden;

            if (!result)
            {
                MessageBox.Show("Error Creating Report. See the Output Window for additional details.");
                return;
            }

            DialogResult = true;
            Close();
        }

        private bool CreateReport(Guid solutionId, string uniqueName, string name, string filePath, string relativePath, int viewableIndex)
        {
            try
            {
                var client = new CrmServiceClient(_connection.ConnectionString);
                Entity report = new Entity("report");
                report["name"] = name;
                report["bodytext"] = File.ReadAllText(filePath);
                report["reporttypecode"] = new OptionSetValue(1); //ReportingServicesReport
                report["filename"] = Path.GetFileName(filePath);
                report["languagecode"] = 1033; //TODO: handle multiple 
                report["ispersonal"] = (viewableIndex == 0);

                Guid id = client.Create(report);

                _logger.WriteToOutputWindow("Report Created: " + id, Logger.MessageType.Info);

                //Add to the choosen solution (not default)
                if (solutionId != new Guid("FD140AAF-4DF4-11DD-BD17-0019B9312238"))
                {
                    AddSolutionComponentRequest scRequest = new AddSolutionComponentRequest
                    {
                        ComponentType = 31,
                        SolutionUniqueName = uniqueName,
                        ComponentId = id
                    };
                    AddSolutionComponentResponse response =
                        (AddSolutionComponentResponse)client.Execute(scRequest);

                    _logger.WriteToOutputWindow("New Report Added To Solution: " + response.id, Logger.MessageType.Info);
                }

                NewId = id;
                NewName = name;
                NewBoudndFile = relativePath;

                return true;
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                _logger.WriteToOutputWindow("Error Creating Report: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, Logger.MessageType.Error);
                return false;
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Creating Report: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
                return false;
            }
        }

        private void Solutions_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Solutions.SelectedItem != null)
            {
                SolutionsLabel.Foreground = Brushes.Black;
                Files.IsEnabled = true;
                Name.IsEnabled = true;
                if (Solutions.SelectedIndex != 0)
                {
                    Viewable.SelectedIndex = 1; //Organization
                    Viewable.IsEnabled = false;
                }
                else
                    Viewable.IsEnabled = true;
            }
            else
            {
                SolutionsLabel.Foreground = Brushes.Red;
                Files.IsEnabled = false;
                Name.IsEnabled = false;
                Viewable.IsEnabled = true;
            }

            Name.Text = null;
            Files.SelectedIndex = -1;
        }

        private void Files_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Files.SelectedItem != null)
            {
                FilesLabel.Foreground = Brushes.Black;
                string fileName = ((ComboBoxItem)Files.SelectedItem).Content.ToString();
                fileName = fileName.Replace("/", String.Empty);
                Name.Text = fileName;
            }
            else
            {
                FilesLabel.Foreground = Brushes.Red;
                Name.Text = null;
            }
        }

        private void Name_TextChanged(object sender, TextChangedEventArgs e)
        {
            NameLabel.Foreground = string.IsNullOrEmpty(Name.Text) ? Brushes.Red : Brushes.Black;
        }
    }
}
