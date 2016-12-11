using EnvDTE;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using OutputLogger;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using CommonResources;
using CommonResources.Models;
using WebResourceDeployer.Models;
using Microsoft.Xrm.Tooling.Connector;

namespace WebResourceDeployer
{
    public partial class NewWebResource
    {
        private readonly CrmConn _connection;
        private readonly Logger _logger;
        private readonly DTE _dte;
        public Guid NewId;
        public int NewType;
        public string NewName;
        public string NewDisplayName;
        public string NewBoundFile;
        public Guid NewSolutionId;
        private const string WindowType = "WebResourceDeployer";

        public NewWebResource(CrmConn connection, ObservableCollection<ComboBoxItem> projectFiles, Guid selectedSolutionId, DTE dte)
        {
            InitializeComponent();

            _logger = new Logger();
            _dte = dte;
            _connection = connection;

            bool result = GetSolutions(selectedSolutionId);

            if (!result)
            {
                MessageBox.Show("Error Retrieving Solutions From CRM. See the Output Window for additional details.");
                DialogResult = false;
                Close();
            }

            Files.ItemsSource = projectFiles;
        }

        private bool GetSolutions(Guid selectedSolutionId)
        {
            try
            {
                List<CrmSolution> solutions = new List<CrmSolution>();

                CrmServiceClient client = SharedConnection.GetCurrentConnection(_connection.ConnectionString, WindowType, _dte);
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

                EntityCollection results = client.RetrieveMultiple(query);


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

                //Default on top
                var i = solutions.FindIndex(s => s.SolutionId == new Guid("FD140AAF-4DF4-11DD-BD17-0019B9312238"));
                var item = solutions[i];
                solutions.RemoveAt(i);
                solutions.Insert(0, item);

                if (selectedSolutionId != Guid.Empty)
                {
                    var sel = solutions.FindIndex(s => s.SolutionId == selectedSolutionId);
                    if (sel != -1)
                        Solutions.SelectedIndex = sel;
                }

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

            ProjectItem projectItem = (ProjectItem)((ComboBoxItem)Files.SelectedItem).Tag;
            string relativePath = ((ComboBoxItem)Files.SelectedItem).Content.ToString();
            string filePath = projectItem.Properties.Item("FullPath").Value.ToString();
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

            bool isNameValid = ValidateName();
            if (!isNameValid)
                return;

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
                CrmServiceClient client = SharedConnection.GetCurrentConnection(_connection.ConnectionString, WindowType, _dte);
                Entity webResource = new Entity("webresource");
                webResource["name"] = prefix + name;
                webResource["webresourcetype"] = new OptionSetValue(type);
                if (!string.IsNullOrEmpty(displayName))
                    webResource["displayname"] = displayName;

                if (type == 8)
                    webResource["silverlightversion"] = "4.0";

                string extension = Path.GetExtension(filePath);

                List<string> imageExs = new List<string>() { ".ICO", ".PNG", ".GIF", ".JPG" };
                string content;
                //TypeScript
                if (extension != null && (extension.ToUpper() == ".TS"))
                {
                    content = File.ReadAllText(Path.ChangeExtension(filePath, ".js"));
                    webResource["content"] = EncodeString(content);
                }
                //Images
                else if (extension != null && imageExs.Any(s => extension.ToUpper().EndsWith(s)))
                {
                    content = EncodedImage(filePath, extension);
                    webResource["content"] = content;
                }
                //Everything else
                else
                {
                    content = File.ReadAllText(filePath);
                    webResource["content"] = EncodeString(content);
                }

                Guid id = client.Create(webResource);

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
                        (AddSolutionComponentResponse)client.Execute(scRequest);

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

        private static string EncodedImage(string filePath, string extension)
        {
            string encodedImage;

            if (extension.ToUpper() == ".ICO")
            {
                System.Drawing.Icon icon = System.Drawing.Icon.ExtractAssociatedIcon(filePath);

                using (MemoryStream ms = new MemoryStream())
                {
                    if (icon != null) icon.Save(ms);
                    byte[] imageBytes = ms.ToArray();
                    encodedImage = Convert.ToBase64String(imageBytes);
                }

                return encodedImage;
            }

            System.Drawing.Image image = System.Drawing.Image.FromFile(filePath, true);

            ImageFormat format = null;
            switch (extension.ToUpper())
            {
                case ".GIF":
                    format = ImageFormat.Gif;
                    break;
                case ".JPG":
                    format = ImageFormat.Jpeg;
                    break;
                case ".PNG":
                    format = ImageFormat.Png;
                    break;
            }

            if (format == null)
                return null;

            using (MemoryStream ms = new MemoryStream())
            {
                image.Save(ms, format);
                byte[] imageBytes = ms.ToArray();
                encodedImage = Convert.ToBase64String(imageBytes);
            }
            return encodedImage;
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
                    Name.Text = fileName.Substring(0, fileName.Length - 3) + ".js";
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

        private bool ValidateName()
        {
            string error =
                "Web resource names may only include letters, numbers, periods, and nonconsecutive forward slash characters.";

            if (string.IsNullOrEmpty(Name.Text))
                return true;

            string name = Name.Text.Trim();

            Regex r = new Regex("^[a-zA-Z0-9_.\\/]*$");
            if (!r.IsMatch(name))
            {
                MessageBox.Show(error);
                return false;
            }

            if (name.Contains("//"))
            {
                MessageBox.Show(error);
                return false;
            }

            return true;
        }
    }
}
