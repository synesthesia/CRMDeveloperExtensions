using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace CRMDeveloperExtensions
{
    class ItemTemplateCommands
    {
        private readonly DTE _dte;
        private string _projectType;
        private string _testType;
        private readonly bool _showWrMenuItems;
        private readonly bool _showPdMenuItems;

        public ItemTemplateCommands(DTE dte)
        {
            _dte = dte;

            var props = _dte.Properties["CRM Developer Extensions", "Web Resource Deployer"];
            _showWrMenuItems = (bool)props.Item("EnableCrmWrContextTemplates").Value;
            props = _dte.Properties["CRM Developer Extensions", "Plug-in Deployer"];
            _showPdMenuItems = (bool)props.Item("EnableCrmPdContextTemplates").Value;
        }

        public void MenuItem1_BeforeQueryStatus(object sender, EventArgs e)
        {
            OleMenuCommand menuCommand = sender as OleMenuCommand;
            if (menuCommand == null) return;

            //Hide menu item if option is disabled
            if (!_showPdMenuItems)
            {
                menuCommand.Visible = false;
                return;
            }

            //Determine if the Project -> Add Item should be displayed and with what text
            GetCrmProject();

            if (_projectType == "PLUGIN" || _projectType == "WORKFLOW")
                menuCommand.Visible = true;
            else
            {
                menuCommand.Visible = false;
                return;
            }

            if (_projectType == "PLUGIN")
            {
                if (string.IsNullOrEmpty(_testType))
                {
                    menuCommand.Text = "CRM Plug-in Class...";
                    return;
                }

                if (_testType == "UNIT")
                    menuCommand.Text = "CRM Plug-in Unit Test...";
                if (_testType == "NUNIT")
                    menuCommand.Text = "CRM Plug-in NUnit Test...";

                return;
            }

            if (_projectType == "WORKFLOW")
            {
                if (string.IsNullOrEmpty(_testType))
                {
                    menuCommand.Text = "CRM Workflow Class...";
                    return;
                }

                if (_testType == "UNIT")
                    menuCommand.Text = "CRM Workflow Unit Test...";
                if (_testType == "NUNIT")
                    menuCommand.Text = "CRM Workflow NUnit Test...";
            }
        }

        public void MenuItem2_BeforeQueryStatus(object sender, EventArgs e)
        {
            OleMenuCommand menuCommand = sender as OleMenuCommand;
            if (menuCommand == null) return;

            //Hide menu item if option is disabled
            if (!_showWrMenuItems)
            {
                menuCommand.Visible = false;
                return;
            }

            //Determine if the Project -> Add Item should be displayed and with what text
            GetCrmProject();

            menuCommand.Visible = _projectType == "WEBRESOURCE";
        }

        public void MenuItem3_BeforeQueryStatus(object sender, EventArgs e)
        {
            OleMenuCommand menuCommand = sender as OleMenuCommand;
            if (menuCommand == null) return;

            //Hide menu item if option is disabled
            if (!_showWrMenuItems)
            {
                menuCommand.Visible = false;
                return;
            }

            //Determine if the Project -> Add Item should be displayed and with what text
            if (string.IsNullOrEmpty(_projectType))
                GetCrmProject();

            menuCommand.Visible = _projectType == "WEBRESOURCE";
        }

        private void GetCrmProject()
        {
            _projectType = null;

            Array activeSolutionProjects = (Array)_dte.ActiveSolutionProjects;
            if (activeSolutionProjects == null || activeSolutionProjects.Length <= 0) return;

            var project = (Project)((Array)(_dte.ActiveSolutionProjects)).GetValue(0);

            var path = Path.GetDirectoryName(project.FullName);
            if (!File.Exists(path + "\\Properties\\settings.settings")) return;

            XmlDocument doc = new XmlDocument();
            doc.Load(path + "\\Properties\\settings.settings");

            XmlNodeList settings = doc.GetElementsByTagName("Settings");
            if (settings.Count == 0) return;

            XmlNodeList appSettings = settings[0].ChildNodes;
            foreach (XmlNode node in appSettings)
            {
                if (node.Attributes != null && node.Attributes["Name"] != null)
                {
                    if (node.Attributes["Name"].Value == "CRMProjectType")
                    {
                        XmlNode value = node.FirstChild;
                        _projectType = value.InnerText;
                    }
                }

                if (node.Attributes != null && node.Attributes["Name"] != null)
                {
                    if (node.Attributes["Name"].Value == "CRMTestType")
                    {
                        XmlNode value = node.FirstChild;
                        _testType = value.InnerText;
                    }
                }
            }
        }

        public void MenuItem1Callback(object sender, EventArgs e)
        {
            bool isUnit = (_testType == "UNIT" || _testType == "NUNIT");
            bool isNunit = _testType == "NUNIT";

            ProjectItem selectedProjectItem = _dte.SelectedItems.Item(1).ProjectItem;
            Solution2 solution = (Solution2)_dte.Application.Solution;
            Project project = (Project)((Array)(_dte.ActiveSolutionProjects)).GetValue(0);

            //Get the exsting class file names so it won't get duplicated when creating a new item
            List<string> fileNames = new List<string>();
            string projectPath = (selectedProjectItem == null)
                ? Path.GetDirectoryName(project.FullName)
                : Path.GetDirectoryName(selectedProjectItem.FileNames[1]);

            if (projectPath != null)
            {
                foreach (string fullname in Directory.GetFiles(projectPath))
                {
                    string filename = Path.GetFileName(fullname);
                    if (!string.IsNullOrEmpty(filename) && filename.ToUpper().EndsWith(".CS"))
                        fileNames.Add(filename);
                }
            }

            //Display the form prompting for the class file name and unit test type
            ClassNamer namer = new ClassNamer(_projectType, isUnit, isNunit, null, fileNames);
            bool? fileNamed = namer.ShowDialog();
            if (!fileNamed.HasValue || !fileNamed.Value) return;
            string testStyle = namer.TestType;
            string templateName = String.Empty;

            //Based on the results, load the proper template in the project
            if (_projectType == "PLUGIN")
            {
                if (!isUnit)
                    templateName = "Plugin Class.csharp.zip";
                else
                {
                    if (!isNunit)
                    {
                        switch (testStyle)
                        {
                            case "Unit":
                                templateName = "Plugin Unit Test.csharp.zip";
                                break;
                            case "Integration":
                                templateName = "Plugin Int Test.csharp.zip";
                                break;
                            default:
                                return;
                        }
                    }
                    else
                    {
                        switch (testStyle)
                        {
                            case "Unit":
                                templateName = "Plugin NUnit Test.csharp.zip";
                                break;
                            case "Integration":
                                templateName = "Plugin NUnit Int Test.csharp.zip";
                                break;
                            default:
                                return;
                        }
                    }
                }
            }

            if (_projectType == "WORKFLOW")
            {
                if (!isUnit)
                    templateName = "Workflow Class.csharp.zip";
                else
                {
                    if (!isNunit)
                    {
                        switch (testStyle)
                        {
                            case "Unit":
                                templateName = "Workflow Unit Test.csharp.zip";
                                break;
                            case "Integration":
                                templateName = "Workflow Int Test.csharp.zip";
                                break;
                            default:
                                return;
                        }
                    }
                    else
                    {
                        switch (testStyle)
                        {
                            case "Unit":
                                templateName = "Workflow NUnit Test.csharp.zip";
                                break;
                            case "Integration":
                                templateName = "Workflow NUnit Int Test.csharp.zip";
                                break;
                            default:
                                return;
                        }
                    }
                }
            }

            if (string.IsNullOrEmpty(templateName)) return;

            var item = solution.GetProjectItemTemplate(templateName, "CSharp");
            _dte.StatusBar.Text = @"Adding class from template...";

            if (selectedProjectItem == null)
                project.ProjectItems.AddFromTemplate(item, namer.FileName);
            else
                selectedProjectItem.ProjectItems.AddFromTemplate(item, namer.FileName);
        }

        public void MenuItem2Callback(object sender, EventArgs e)
        {
            ProjectItem selectedProjectItem = _dte.SelectedItems.Item(1).ProjectItem;
            Solution2 solution = (Solution2)_dte.Application.Solution;
            Project project = (Project)((Array)(_dte.ActiveSolutionProjects)).GetValue(0);

            //Get the exsting file names so it won't get duplicated when creating a new item
            List<string> fileNames = new List<string>();
            string projectPath = (selectedProjectItem == null)
                ? Path.GetDirectoryName(project.FullName)
                : Path.GetDirectoryName(selectedProjectItem.FileNames[1]);

            if (projectPath != null)
            {
                foreach (string fullname in Directory.GetFiles(projectPath))
                {
                    string filename = Path.GetFileName(fullname);
                    if (!string.IsNullOrEmpty(filename) && (filename.ToUpper().EndsWith(".HTML") || filename.ToUpper().EndsWith(".HTM")))
                        fileNames.Add(filename);
                }
            }

            //Display the form prompting for the file name
            ClassNamer namer = new ClassNamer(_projectType, false, false, "HTML", fileNames);
            bool? fileNamed = namer.ShowDialog();
            if (!fileNamed.HasValue || !fileNamed.Value) return;
            string templateName = "HTML Web.csharp.zip";

            var item = solution.GetProjectItemTemplate(templateName, "CSharp");
            _dte.StatusBar.Text = @"Adding file from template...";

            if (selectedProjectItem == null)
                project.ProjectItems.AddFromTemplate(item, namer.FileName);
            else
                selectedProjectItem.ProjectItems.AddFromTemplate(item, namer.FileName);
        }

        public void MenuItem3Callback(object sender, EventArgs e)
        {
            ProjectItem selectedProjectItem = _dte.SelectedItems.Item(1).ProjectItem;
            Solution2 solution = (Solution2)_dte.Application.Solution;
            Project project = (Project)((Array)(_dte.ActiveSolutionProjects)).GetValue(0);

            //Get the exsting file names so it won't get duplicated when creating a new item
            List<string> fileNames = new List<string>();

            string projectPath = (selectedProjectItem == null)
                ? Path.GetDirectoryName(project.FullName)
                : Path.GetDirectoryName(selectedProjectItem.FileNames[1]);

            if (projectPath != null)
            {
                foreach (string fullname in Directory.GetFiles(projectPath))
                {
                    string filename = Path.GetFileName(fullname);
                    if (!string.IsNullOrEmpty(filename) && filename.ToUpper().EndsWith(".JS"))
                        fileNames.Add(filename);
                }
            }

            //Display the form prompting for the file name
            ClassNamer namer = new ClassNamer(_projectType, false, false, "JS", fileNames);
            bool? fileNamed = namer.ShowDialog();
            if (!fileNamed.HasValue || !fileNamed.Value) return;
            string templateName = "JavaScript Web.csharp.zip";

            var item = solution.GetProjectItemTemplate(templateName, "CSharp");
            _dte.StatusBar.Text = @"Adding file from template...";

            if (selectedProjectItem == null)
                project.ProjectItems.AddFromTemplate(item, namer.FileName);
            else
                selectedProjectItem.ProjectItems.AddFromTemplate(item, namer.FileName);
        }
    }
}
