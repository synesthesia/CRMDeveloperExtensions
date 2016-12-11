using CommonResources;
using CommonResources.Models;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using OutputLogger;
using PluginDeployer.Models;
using System;
using System.ComponentModel.Design;
using System.IO;
using System.Runtime.InteropServices;
using System.ServiceModel;
using Window = EnvDTE.Window;

namespace PluginDeployer
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.GuidPluginDeployerPkgString)]
    [ProvideAutoLoad("f1536ef8-92ec-443c-9ed7-fdadf150da82")]
    public sealed class PluginDeployerPackage : Package
    {
        private DTE _dte;
        private Logger _logger;
        private const string WindowType = "PluginDeployer";

        protected override void Initialize()
        {
            base.Initialize();

            _logger = new Logger();

            _dte = GetGlobalService(typeof(DTE)) as DTE;
            if (_dte == null)
                return;

            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (mcs != null)
            {
                CommandID publishCommandId = new CommandID(GuidList.GuidProjectMenuCommandsCmdSet, (int)PkgCmdIdList.CmdidPluginDeployerPublish);
                OleMenuCommand publishMenuItem = new OleMenuCommand(PublishItemCallback, publishCommandId);
                publishMenuItem.BeforeQueryStatus += PublishItem_BeforeQueryStatus;
                publishMenuItem.Visible = false;
                mcs.AddCommand(publishMenuItem);
            }
        }

        private void PublishItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            OleMenuCommand menuCommand = sender as OleMenuCommand;
            if (menuCommand == null) return;

            bool windowOpen = false;
            foreach (Window window in _dte.Windows)
            {
                if (window.Caption != Resources.ResourceManager.GetString("ToolWindowTitle")) continue;

                windowOpen = window.Visible;
                break;
            }

            if (!windowOpen)
            {
                menuCommand.Visible = false;
                return;
            }

            if (_dte.SelectedItems.Count != 1)
            {
                menuCommand.Visible = false;
                return;
            }

            CrmConn selectedConnection = (CrmConn)SharedGlobals.GetGlobal("SelectedConnection", _dte);
            if (selectedConnection == null)
            {
                menuCommand.Visible = false;
                return;
            }

            if (SelectedAssemblyItem.Item == null)
            {
                menuCommand.Visible = false;
                return;
            }

            Guid assemblyId = SelectedAssemblyItem.Item.AssemblyId;
            menuCommand.Visible = assemblyId != Guid.Empty;
        }

        private void PublishItemCallback(object sender, EventArgs e)
        {
            if (_dte.SelectedItems.Count != 1) return;

            SelectedItem item = _dte.SelectedItems.Item(1);
            Project project = item.Project;

            CrmConn selectedConnection = (CrmConn)SharedGlobals.GetGlobal("SelectedConnection", _dte);
            if (selectedConnection == null) return;

            Guid assemblyId = SelectedAssemblyItem.Item.AssemblyId;
            if (assemblyId == Guid.Empty) return;

            CrmServiceClient client = SharedConnection.GetCurrentConnection(selectedConnection.ConnectionString, WindowType, _dte);

            UpdateAndPublishSingle(client, project);
        }

        private void UpdateAndPublishSingle(CrmServiceClient client, Project project)
        {
            try
            {
                _dte.StatusBar.Text = "Deploying assembly...";
                _dte.StatusBar.Animate(true, vsStatusAnimation.vsStatusAnimationDeploy);

                string outputFileName = project.Properties.Item("OutputFileName").Value.ToString();
                string path = GetOutputPath(project) + outputFileName;

                //Build the project
                SolutionBuild solutionBuild = _dte.Solution.SolutionBuild;
                solutionBuild.BuildProject(_dte.Solution.SolutionBuild.ActiveConfiguration.Name, project.UniqueName, true);

                if (solutionBuild.LastBuildInfo > 0)
                    return;

                //Make sure Major and Minor versions match
                Version assemblyVersion = Version.Parse(project.Properties.Item("AssemblyVersion").Value.ToString());
                if (SelectedAssemblyItem.Item.Version.Major != assemblyVersion.Major ||
                    SelectedAssemblyItem.Item.Version.Minor != assemblyVersion.Minor)
                {
                    _logger.WriteToOutputWindow("Error Updating Assembly In CRM: Changes To Major & Minor Versions Require Redeployment", Logger.MessageType.Error);
                    return;
                }

                //Make sure assembly names match
                string assemblyName = project.Properties.Item("AssemblyName").Value.ToString();
                if (assemblyName.ToUpper() != SelectedAssemblyItem.Item.Name.ToUpper())
                {
                    _logger.WriteToOutputWindow("Error Updating Assembly In CRM: Changes To Assembly Name Require Redeployment", Logger.MessageType.Error);
                    return;
                }

                //Update CRM
                Entity crmAssembly = new Entity("pluginassembly") { Id = SelectedAssemblyItem.Item.AssemblyId };
                crmAssembly["content"] = Convert.ToBase64String(File.ReadAllBytes(path));

                client.Update(crmAssembly);

                //Update assembly name and version numbers
                SelectedAssemblyItem.Item.Version = assemblyVersion;
                SelectedAssemblyItem.Item.Name = project.Properties.Item("AssemblyName").Value.ToString();
                SelectedAssemblyItem.Item.DisplayName = SelectedAssemblyItem.Item.Name + " (" + assemblyVersion + ")";
                SelectedAssemblyItem.Item.DisplayName += (SelectedAssemblyItem.Item.IsWorkflowActivity) ? " [Workflow]" : " [Plug-in]";
            }
            catch (FaultException<OrganizationServiceFault> crmEx)
            {
                _logger.WriteToOutputWindow("Error Deploying Report To CRM: " + crmEx.Message + Environment.NewLine + crmEx.StackTrace, Logger.MessageType.Error);
            }
            catch (Exception ex)
            {
                _logger.WriteToOutputWindow("Error Deploying Report To CRM: " + ex.Message + Environment.NewLine + ex.StackTrace, Logger.MessageType.Error);
            }

            _dte.StatusBar.Clear();
            _dte.StatusBar.Animate(false, vsStatusAnimation.vsStatusAnimationDeploy);
        }

        private string GetOutputPath(Project project)
        {
            ConfigurationManager configurationManager = project.ConfigurationManager;
            if (configurationManager == null) return null;

            Configuration activeConfiguration = configurationManager.ActiveConfiguration;
            string outputPath = activeConfiguration.Properties.Item("OutputPath").Value.ToString();
            string absoluteOutputPath = String.Empty;
            string projectFolder;

            if (outputPath.StartsWith(Path.DirectorySeparatorChar.ToString() + Path.DirectorySeparatorChar))
            {
                absoluteOutputPath = outputPath;
            }
            else if (outputPath.Length >= 2 && outputPath[0] == Path.VolumeSeparatorChar)
            {
                absoluteOutputPath = outputPath;
            }
            else if (outputPath.IndexOf("..\\", StringComparison.Ordinal) != -1)
            {
                projectFolder = Path.GetDirectoryName(project.FullName);

                while (outputPath.StartsWith("..\\"))
                {
                    outputPath = outputPath.Substring(3);
                    projectFolder = Path.GetDirectoryName(projectFolder);
                }

                if (projectFolder != null) absoluteOutputPath = Path.Combine(projectFolder, outputPath);
            }
            else
            {
                projectFolder = Path.GetDirectoryName(project.FullName);
                if (projectFolder != null)
                    absoluteOutputPath = Path.Combine(projectFolder, outputPath);
            }

            return absoluteOutputPath;
        }
    }
}
