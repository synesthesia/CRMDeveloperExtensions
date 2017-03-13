﻿using EnvDTE;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.TemplateWizard;
using NuGet.VisualStudio;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Xml;
using System.Xml.Linq;
using VSLangProj;

namespace TemplateWizards
{
    public class VersionWizard : IWizard
    {
        [DllImport("mscoree.dll")]
        internal static extern int StrongNameFreeBuffer(IntPtr pbMemory);
        [DllImport("mscoree.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        internal static extern int StrongNameKeyGen(IntPtr wszKeyContainer, uint dwFlags, out IntPtr keyBlob, out uint keyBlobSize);
        [DllImport("mscoree.dll", CharSet = CharSet.Unicode)]
        internal static extern int StrongNameErrorInfo();

        private DTE _dte;
        private string _sdkVersion;
        private string _project;
        private string _crmProjectType = "Plug-in";
        private string _needsClient = "False";
        private string _isUnitTest = "False";
        private string _isUnitTestItem = "False";
        private bool _isNunit;
        private string _destDirectory;

        public void RunStarted(object automationObject, Dictionary<string, string> replacementsDictionary, WizardRunKind runKind, object[] customParams)
        {
            try
            {
                _dte = (DTE)automationObject;

                if (replacementsDictionary.ContainsKey("$destinationdirectory$"))
                    _destDirectory = replacementsDictionary["$destinationdirectory$"];

                if (replacementsDictionary.ContainsKey("$wizarddata$"))
                {
                    string wizardData = replacementsDictionary["$wizarddata$"];
                    ReadWizardData(wizardData);
                }

                //Find default SDK version
                var props = _dte.Properties["CRM Developer Extensions", "General"];
                string defaultSdkVersion = props.Item("DefaultCrmSdkVersion").Value.ToString();

                //TypeScript projects
                if (_crmProjectType == "TypeScript")
                {
                    Version version = Version.Parse(_dte.Version);
                    if (version.Major < 12)
                        throw new WizardBackoutException("TypeScript projects require Visual Studio 2013 or later");

                    TypeScriptPicker tsPicker = new TypeScriptPicker(defaultSdkVersion);
                    tsPicker.ShowDialog();

                    _sdkVersion = tsPicker.Version;

                    return;
                }

                //If UnitTest Item - load the assembly & class names from the referenced project into the picker
                if (_isUnitTestItem == "True")
                {
                    List<string> referencedProjects = new List<string>();
                    Array activeSolutionProjects = (Array)_dte.ActiveSolutionProjects;
                    if (activeSolutionProjects != null && activeSolutionProjects.Length > 0)
                    {
                        VSProject vsproject = (VSProject) ((Project)activeSolutionProjects.GetValue(0)).Object;
                        referencedProjects.AddRange(from Reference reference in vsproject.References
                                                    where reference.SourceProject != null
                                                    select reference.Name);

                        List<ComboBoxItem> classItems = GetSourceProjectItems(referencedProjects);
                        var testClassPickerform = new TestClassPicker(classItems);
                        testClassPickerform.ShowDialog();

                        replacementsDictionary.Add("$fullclassname$", testClassPickerform.FullClassname);
                        replacementsDictionary.Add("$assemblyname$", testClassPickerform.AssemblyName);
                        _sdkVersion = GetSdkVersion(vsproject.Project);
                        if (_sdkVersion == "CRM 2016 (8.0.X)" || _sdkVersion == "CRM 2016 (8.1.X)" || _sdkVersion == "CRM 2016 (8.2.X)")
                            replacementsDictionary.Add("$useXrmToolingClientUsing$", "1");

                        else
                            replacementsDictionary.Add("$useXrmToolingClientUsing$", "0");
                    }

                    return;
                }

                //If UnitTest Project - load the projects into the picker
                List<ComboBoxItem> projectItems = new List<ComboBoxItem>();
                if (_isUnitTest == "True")
                    projectItems = GetSourceProjects();

                //Display the form prompting for the SDK version and/or project to unit test against
                if (_crmProjectType == "Plug-in" || _crmProjectType == "Workflow" || _crmProjectType == "Console")
                {
                    var form = new SdkProjectPicker((_isUnitTest == "True"), projectItems, defaultSdkVersion);
                    form.ShowDialog();

                    _sdkVersion = form.Version;
                    _project = form.Project;
                    _isNunit = form.Nunit;
                }

                if (_sdkVersion == null)
                {
                    Array activeSolutionProjects = (Array)_dte.ActiveSolutionProjects;
                    if (activeSolutionProjects != null && activeSolutionProjects.Length > 0)
                        _sdkVersion = GetSdkVersion(((Project)activeSolutionProjects.GetValue(0)));
                }
                if (_sdkVersion == "CRM 2016 (8.0.X)" || _sdkVersion == "CRM 2016 (8.1.X)" || _sdkVersion == "CRM 2016 (8.2.X)")
                    replacementsDictionary.Add("$useXrmToolingClientUsing$", "1");
                else
                    replacementsDictionary.Add("$useXrmToolingClientUsing$", "0");

                //If UnitTest Project - set the reference to the project being tested
                if (_isUnitTest == "True")
                {
                    if (string.IsNullOrEmpty(_project))
                    {
                        replacementsDictionary.Add("$referenceproject$", "False");
                        return;
                    }

                    Projects projects = _dte.Solution.Projects;
                    Solution solution = _dte.Solution;
                    foreach (Project project in projects)
                    {
                        if (project.Name != _project) continue;

                        string path = string.Empty;
                        string projectPath = Path.GetDirectoryName(project.FullName);
                        string solutionPath = Path.GetDirectoryName(solution.FullName);
                        if (!string.IsNullOrEmpty(projectPath) && !string.IsNullOrEmpty(solutionPath))
                        {
                            if (projectPath.StartsWith(solutionPath))
                                path = "..\\" + project.UniqueName;
                            else
                                path = project.FullName;
                        }

                        replacementsDictionary.Add("$referenceproject$", "True");
                        replacementsDictionary.Add("$projectPath$", path);
                        replacementsDictionary.Add("$projectId$", project.Kind);
                        replacementsDictionary.Add("$projectName$", project.Name);
                        break;
                    }
                }
                else
                    replacementsDictionary.Add("$referenceproject$", "False");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Processing Template: " + ex.Message);
            }
        }

        public void BeforeOpeningFile(ProjectItem projectItem)
        {
        }

        public void ProjectFinishedGenerating(Project project)
        {
            var componentModel = (IComponentModel)Package.GetGlobalService(typeof(SComponentModel));
            if (componentModel == null) return;

            var installer = componentModel.GetService<IVsPackageInstaller>();

            switch (_crmProjectType)
            {
                case "Console":
                case "Plug-in":
                case "Workflow":
                    HandleCrmAssemblyProjects(project, installer);
                    break;
                case "TypeScript":
                    HandleTypeScriptProject(project, installer);
                    break;
                case "Package":
                    HandleSolutionPackagerProject(project);
                    break;
            }
        }

        private void HandleSolutionPackagerProject(Project project)
        {
            foreach (SolutionConfiguration solutionConfiguration in _dte.Solution.SolutionBuild.SolutionConfigurations)
            {
                foreach (SolutionContext solutionContext in solutionConfiguration.SolutionContexts)
                {
                    solutionContext.ShouldBuild = false;
                }
            }

            //Delete bin & obj folders
            Directory.Delete(Path.GetDirectoryName(project.FullName) + "//bin", true);
            Directory.Delete(Path.GetDirectoryName(project.FullName) + "//obj", true);
        }

        private void HandleTypeScriptProject(Project project, IVsPackageInstaller installer)
        {
            try
            {
                InstallPackage(installer, project, "xrm.TypeScript.DefinitelyTyped", null);

                //Remove the files that don't match the selected version
                ProjectItem scripts = project.ProjectItems.Item("Scripts");
                ProjectItem typings = scripts.ProjectItems.Item("typings");
                ProjectItem xrm = typings.ProjectItems.Item("xrm");

                ProjectItem xrm6 = xrm.ProjectItems.Item("xrm-6.d.ts");
                ProjectItem xrm70 = xrm.ProjectItems.Item("xrm-7.0.d.ts");
                ProjectItem xrm71 = xrm.ProjectItems.Item("xrm-7.1.d.ts");
                ProjectItem xrm80 = xrm.ProjectItems.Item("xrm.d.ts");
                ProjectItem para = xrm.ProjectItems.Item("parature.d.ts");
                string filename;

                switch (_sdkVersion)
                {
                    case "CRM 2013 (6.0.X)":
                        filename = xrm70.FileNames[0];
                        xrm70.Remove();
                        File.Delete(filename);

                        filename = xrm71.FileNames[0];
                        xrm71.Remove();
                        File.Delete(filename);

                        filename = para.FileNames[0];
                        para.Remove();
                        File.Delete(filename);
                        break;
                    case "CRM 2015 (7.0.X)":
                        filename = xrm6.FileNames[0];
                        xrm6.Remove();
                        File.Delete(filename);

                        filename = xrm71.FileNames[0];
                        xrm71.Remove();
                        File.Delete(filename);

                        filename = para.FileNames[0];
                        para.Remove();
                        File.Delete(filename);
                        break;
                    case "CRM 2015 (7.1.X)":
                        filename = xrm6.FileNames[0];
                        xrm6.Remove();
                        File.Delete(filename);

                        filename = xrm70.FileNames[0];
                        xrm70.Remove();
                        File.Delete(filename);

                        filename = xrm80.FileNames[0];
                        xrm80.Remove();
                        File.Delete(filename);
                        break;
                    case "CRM 2016 (8.0.X)":
                        filename = xrm6.FileNames[0];
                        xrm6.Remove();
                        File.Delete(filename);

                        filename = xrm70.FileNames[0];
                        xrm70.Remove();
                        File.Delete(filename);

                        filename = xrm71.FileNames[0];
                        xrm71.Remove();
                        File.Delete(filename);
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Processing Template: " + ex.Message);
            }
        }

        private void HandleCrmAssemblyProjects(Project project, IVsPackageInstaller installer)
        {
            try
            {
                //Install the proper NuGet packages based on the CRM SDK version and project types
                switch (_sdkVersion)
                {
                    case "CRM 2011 (5.0.X)":
                        InstallPackage(installer, project, "Microsoft.CrmSdk.CoreAssemblies", "5.0.18");
                        if (_crmProjectType == "Workflow")
                            InstallPackage(installer, project, "Microsoft.CrmSdk.Workflow", "5.0.18");
                        if (_needsClient == "True")
                            InstallPackage(installer, project, "Microsoft.CrmSdk.Extensions", "5.0.18");
                        break;
                    case "CRM 2013 (6.0.X)":
                        InstallPackage(installer, project, "Microsoft.CrmSdk.CoreAssemblies", "6.0.4");
                        if (_crmProjectType == "Workflow")
                            InstallPackage(installer, project, "Microsoft.CrmSdk.Workflow", "6.0.4");
                        if (_needsClient == "True")
                            InstallPackage(installer, project, "Microsoft.CrmSdk.Extensions", "6.0.4.1");
                        break;
                    case "CRM 2013 (6.1.X)":
                        InstallPackage(installer, project, "Microsoft.CrmSdk.CoreAssemblies", "6.1.1");
                        if (_crmProjectType == "Workflow")
                            InstallPackage(installer, project, "Microsoft.CrmSdk.Workflow", "6.1.1");
                        if (_needsClient == "True")
                            InstallPackage(installer, project, "Microsoft.CrmSdk.Extensions", "6.0.4.1");
                        break;
                    case "CRM 2015 (7.0.X)":
                        project.DTE.SuppressUI = true;
                        project.Properties.Item("TargetFrameworkMoniker").Value = ".NETFramework,Version=v4.5.2";
                        project = (Project)((Array)(_dte.ActiveSolutionProjects)).GetValue(0);
                        InstallPackage(installer, project, "Microsoft.CrmSdk.CoreAssemblies", "7.0.1");
                        if (_crmProjectType == "Workflow")
                            InstallPackage(installer, project, "Microsoft.CrmSdk.Workflow", "7.0.1");
                        if (_needsClient == "True")
                            InstallPackage(installer, project, "Microsoft.CrmSdk.Extensions", "7.0.0.1");
                        break;
                    case "CRM 2015 (7.1.X)":
                        project.DTE.SuppressUI = true;
                        project.Properties.Item("TargetFrameworkMoniker").Value = ".NETFramework,Version=v4.5.2";
                        project = (Project)((Array)(_dte.ActiveSolutionProjects)).GetValue(0);
                        InstallPackage(installer, project, "Microsoft.CrmSdk.CoreAssemblies", "7.1.1");
                        if (_crmProjectType == "Workflow")
                            InstallPackage(installer, project, "Microsoft.CrmSdk.Workflow", "7.1.1");
                        if (_needsClient == "True")
                            InstallPackage(installer, project, "Microsoft.CrmSdk.Extensions", null);
                        break;
                    case "CRM 2016 (8.0.X)":
                        project.DTE.SuppressUI = true;
                        project.Properties.Item("TargetFrameworkMoniker").Value = ".NETFramework,Version=v4.5.2";
                        project = (Project)((Array)(_dte.ActiveSolutionProjects)).GetValue(0);
                        InstallPackage(installer, project, "Microsoft.CrmSdk.CoreAssemblies", "8.0.2.1");
                        if (_crmProjectType == "Workflow")
                            InstallPackage(installer, project, "Microsoft.CrmSdk.Workflow", "8.0.2.1");
                        if (_needsClient == "True")
                            InstallPackage(installer, project, "Microsoft.CrmSdk.XrmTooling.CoreAssembly", "8.0.2");
                        break;
                    case "CRM 2016 (8.1.X)":
                        project.DTE.SuppressUI = true;
                        project.Properties.Item("TargetFrameworkMoniker").Value = ".NETFramework,Version=v4.5.2";
                        project = (Project)((Array)(_dte.ActiveSolutionProjects)).GetValue(0);
                        InstallPackage(installer, project, "Microsoft.CrmSdk.CoreAssemblies", "8.1.0.2");
                        if (_crmProjectType == "Workflow")
                            InstallPackage(installer, project, "Microsoft.CrmSdk.Workflow", "8.1.0.2");
                        if (_needsClient == "True")
                            InstallPackage(installer, project, "Microsoft.CrmSdk.XrmTooling.CoreAssembly", null);
                        break;
                    case "CRM 2016 (8.2.X)":
                        project.DTE.SuppressUI = true;
                        project.Properties.Item("TargetFrameworkMoniker").Value = ".NETFramework,Version=v4.5.2";
                        project = (Project)((Array)(_dte.ActiveSolutionProjects)).GetValue(0);
                        InstallPackage(installer, project, "Microsoft.CrmSdk.CoreAssemblies", null);
                        if (_crmProjectType == "Workflow")
                            InstallPackage(installer, project, "Microsoft.CrmSdk.Workflow", null);
                        if (_needsClient == "True")
                            InstallPackage(installer, project, "Microsoft.CrmSdk.XrmTooling.CoreAssembly", null);
                        break;
                }

                ExcludeSdkBinFolder(project);

                if (_crmProjectType == "Plug-in" || _crmProjectType == "Workflow")
                    GenerateNewKey(project);

                if (_isUnitTest != "True") return;

                InstallPackage(installer, project, "Moq", "4.5.28");
                if (_isNunit)
                {
                    InstallPackage(installer, project, "NUnitTestAdapter.WithFramework", "2.0.0");
                    AddSetting(project, "CRMTestType", "NUNIT");
                }
                else
                    AddSetting(project, "CRMTestType", "UNIT");

                ExcludePerformaceFolder(project);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Processing Template: " + ex.Message);
            }
        }

        /// <summary>
        /// Adds the setting value to the project settings.
        /// </summary>
        /// <param name="project">The target project.</param>
        /// <param name="setting">The target setting.</param>
        /// <param name="value">The value to set.</param>
        private static void AddSetting(Project project, string setting, string value)
        {
            try
            {
                var path = Path.GetDirectoryName(project.FullName);
                if (!File.Exists(path + "\\Properties\\settings.settings")) return;

                XmlDocument doc = new XmlDocument();
                doc.Load(path + "\\Properties\\settings.settings");

                XmlNodeList settings = doc.GetElementsByTagName("Settings");
                if (settings.Count == 0) return;

                XmlNodeList appSettings = settings[0].ChildNodes;
                foreach (XmlNode node in appSettings)
                {
                    if (node.Attributes == null || node.Attributes["Name"].Value != setting) continue;

                    XmlNode valueNode = node.FirstChild;
                    valueNode.InnerText = value;
                    doc.Save(path + "\\Properties\\settings.settings");
                    break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Processing Template: " + ex.Message);
            }
        }

        /// <summary>
        /// Installs the specified NuGet package and displays a message in the status bar.
        /// </summary>
        /// <param name="installer">The VS package installer.</param>
        /// <param name="project">The target project.</param>
        /// <param name="package">The NuGet package name.</param>
        /// <param name="version">The NuGet package version.</param>
        private void InstallPackage(IVsPackageInstaller installer, Project project, string package, string version)
        {
            try
            {
                string nuGetSource = "https://www.nuget.org/api/v2/";
                _dte.StatusBar.Text = @"Installing " + package + "...";
                if (!string.IsNullOrEmpty(version))
                    installer.InstallPackage(nuGetSource, project, package, version, false);
                else
                    installer.InstallPackage(nuGetSource, project, package, (Version)null, false);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Processing Template: Error Installing NuGet Package: " + ex.Message);
            }
        }

        public void ProjectItemFinishedGenerating(ProjectItem projectItem)
        {

        }

        public void RunFinished()
        {
        }

        public bool ShouldAddProjectItem(string filePath)
        {
            return true;
        }

        /// <summary>
        /// Generates a new strong key.
        /// </summary>
        /// <param name="project">The project.</param>
        private void GenerateNewKey(Project project)
        {
            try
            {
                if (string.IsNullOrEmpty(_destDirectory))
                    return;

                //Generate new key
                _dte.StatusBar.Text = "Generating key...";

                string keyFilePath = Path.Combine(_destDirectory, "MyKey.snk");
                IntPtr buffer = IntPtr.Zero;

                try
                {
                    uint buffSize;
                    if (0 != StrongNameKeyGen(IntPtr.Zero, 0, out buffer, out buffSize))
                        Marshal.ThrowExceptionForHR(StrongNameErrorInfo());
                    if (buffer == IntPtr.Zero)
                        throw new InvalidOperationException("StrongNameKeyGen Failed");

                    var keyBuffer = new byte[buffSize];
                    Marshal.Copy(buffer, keyBuffer, 0, (int)buffSize);
                    File.WriteAllBytes(keyFilePath, keyBuffer);
                }
                finally
                {
                    StrongNameFreeBuffer(buffer);
                }

                var props = _dte.Properties["CRM Developer Extensions", "General"];
                string defaultKeyFileName = (string) props.Item("DefaultProjectKeyFileName").Value;

                foreach (ProjectItem item in project.ProjectItems)
                {
                    if (item.Name.ToUpper() != "MYKEY.SNK") continue;

                    item.Name = defaultKeyFileName + ".snk";
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Processing Template: Error Generating Key: " + ex.Message);
            }
        }

        /// <summary>
        /// Reads the wizard data that is stored in the _project.vstemplate.xml file.
        /// </summary>
        /// <param name="wizardData">The wizard data.</param>
        private void ReadWizardData(string wizardData)
        {
            XmlReaderSettings settings = new XmlReaderSettings { ConformanceLevel = ConformanceLevel.Fragment };
            using (XmlReader reader = XmlReader.Create(new StringReader(wizardData), settings))
            {
                while (reader.Read())
                {
                    if (reader.Name == "CRMProjectType")
                    {
                        XElement el = XNode.ReadFrom(reader) as XElement;
                        if (el != null)
                            _crmProjectType = el.Value;
                    }
                    if (reader.Name == "NeedsClient")
                    {
                        XElement el = XNode.ReadFrom(reader) as XElement;
                        if (el != null)
                            _needsClient = el.Value;
                    }
                    if (reader.Name == "IsUnitTest")
                    {
                        XElement el = XNode.ReadFrom(reader) as XElement;
                        if (el != null)
                            _isUnitTest = el.Value;
                    }
                    if (reader.Name == "IsUnitTestItem")
                    {
                        XElement el = XNode.ReadFrom(reader) as XElement;
                        if (el != null)
                            _isUnitTestItem = el.Value;
                    }
                }
            }
        }

        /// <summary>
        /// Gets the available assemblies & classes of which unit test items can be created against.
        /// </summary>
        /// <param name="referencedProjects">The solution projects which are referenced by the unit test project.</param>
        /// <returns>List of ComboboxItems containing the project assembly & class info.</returns>
        private static List<ComboBoxItem> GetSourceProjectItems(ICollection<string> referencedProjects)
        {
            List<ComboBoxItem> classItems = new List<ComboBoxItem>();
            var dte = Package.GetGlobalService(typeof(DTE)) as DTE;
            if (dte == null)
                return classItems;
            Projects projects = dte.Solution.Projects;

            ComboBoxItem itemEmpty = new ComboBoxItem { Content = string.Empty, Tag = string.Empty };
            classItems.Add(itemEmpty);

            foreach (Project project in projects)
            {
                if (!referencedProjects.Contains(project.Name)) continue;

                string assemblyName = project.Properties.Item("AssemblyName").Value.ToString();

                ProjectItems activeProjectItems = project.ProjectItems;
                if (activeProjectItems.Count <= 0) continue;

                foreach (ProjectItem projectItem in activeProjectItems)
                {
                    if (!projectItem.FileNames[0].EndsWith((".cs"))) continue;

                    foreach (CodeElement element in projectItem.FileCodeModel.CodeElements)
                    {
                        if (element.Kind != vsCMElement.vsCMElementNamespace) continue;

                        foreach (CodeElement childElement in element.Children)
                        {
                            if (childElement.Kind != vsCMElement.vsCMElementClass) continue;

                            ComboBoxItem item = new ComboBoxItem { Content = childElement.FullName, Tag = assemblyName };
                            classItems.Add(item);
                        }
                    }
                }
            }

            return classItems;
        }

        /// <summary>
        /// Gets the available solution projects of which unit test projects can be created against.
        /// </summary>
        /// <returns>List of ComboboxItems containing the project info.</returns>
        private static List<ComboBoxItem> GetSourceProjects()
        {
            List<ComboBoxItem> projectItems = new List<ComboBoxItem>();

            try
            {
                var dte = Package.GetGlobalService(typeof(DTE)) as DTE;
                if (dte == null)
                    return projectItems;
                Projects projects = dte.Solution.Projects;

                ComboBoxItem itemEmpty = new ComboBoxItem { Content = string.Empty, Tag = string.Empty };
                projectItems.Add(itemEmpty);

                foreach (Project project in projects)
                {
                    //Don't add existing unit test projects to the list of projects to unit test against
                    bool isUnitTestProject = false;

                    if (string.IsNullOrEmpty(project.FullName)) continue;

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

                    if (isUnitTestProject) continue;

                    var sdkVersion = GetSdkVersion(project);

                    ComboBoxItem item = new ComboBoxItem { Content = project.Name, Tag = sdkVersion };
                    projectItems.Add(item);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Processing Template: Error Reading Projects: " + ex.Message);
            }

            return projectItems;
        }

        private static string GetSdkVersion(Project project)
        {
            /*Examine the NuGet packages.config file to determine the CRM SDK version of the target project
            Looking at the assembly version doesn't work the underlying assembly version doesn't always change
            if no updates were done between versions. Example Microsoft.CrmSdk.Extensions assembly versions did not
            change between 7.0.x and 7.1.x*/
            string sdkVersion = String.Empty;
            var packagePath = Path.GetDirectoryName(project.FullName);
            if (File.Exists(packagePath + "\\packages.config"))
            {
                XmlDocument packageDoc = new XmlDocument();
                packageDoc.Load(packagePath + "\\packages.config");

                XmlNodeList packages = packageDoc.GetElementsByTagName("packages");
                if (packages.Count > 0)
                {
                    XmlNodeList packageDetails = packages[0].ChildNodes;
                    foreach (XmlNode package in packageDetails)
                    {
                        if (package.Attributes == null || package.Attributes["id"] == null) continue;
                        if (package.Attributes["id"].Value != "Microsoft.CrmSdk.CoreAssemblies") continue;
                        if (package.Attributes["version"] == null) continue;

                        switch (package.Attributes["version"].Value)
                        {
                            case "5.0.18":
                                return "CRM 2011 (5.0.X)";
                            case "6.0.4":
                                return "CRM 2013 (6.0.X)";
                            case "6.1.1":
                                return "CRM 2013 (6.1.X)";
                            case "7.0.1":
                                return "CRM 2015 (7.0.X)";
                            case "7.1.1":
                                return "CRM 2015 (7.1.X)";
                            case "8.0.2.1":
                                return "CRM 2016 (8.0.X)";
                            case "8.1.0.2":
                                return "CRM 2016 (8.1.X)";
                            default:
                                return "CRM 2016 (8.2.X)";
                        }
                    }
                }
            }
            return sdkVersion;
        }

        /// <summary>
        /// Excludes the bin folder which gets added to the project when installing the v5/6 Microsoft.CrmSdk.CoreAssemblies NuGet package.
        /// </summary>
        /// <param name="project">The target project.</param>
        private static void ExcludeSdkBinFolder(Project project)
        {
            for (int i = 1; i <= project.ProjectItems.Count; i++)
            {
                string itemType = project.ProjectItems.Item(i).Kind;
                if (itemType.ToUpper() == "{6BB5F8EF-4483-11D3-8BCF-00C04F8EC28C}") //GUID_ItemType_PhysicalFolder
                {
                    if (project.ProjectItems.Item(i).Name.ToUpper() == "BIN")
                        project.ProjectItems.Item(i).Remove();
                }
            }
        }

        /// <summary>
        /// Excludes the performance folder which appears in VS2015
        /// </summary>
        /// <param name="project">The target project.</param>
        private static void ExcludePerformaceFolder(Project project)
        {
            for (int i = 1; i <= project.ProjectItems.Count; i++)
            {
                string itemType = project.ProjectItems.Item(i).Kind;
                if (itemType.ToUpper() == "{6BB5F8EF-4483-11D3-8BCF-00C04F8EC28C}") //GUID_ItemType_PhysicalFolder
                {
                    if (project.ProjectItems.Item(i).Name.ToUpper() == "PERFORMANCE")
                        project.ProjectItems.Item(i).Remove();
                }
            }
        }
    }
}
