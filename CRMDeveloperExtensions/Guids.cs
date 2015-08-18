using System;

namespace CRMDeveloperExtensions
{
    static class GuidList
    {
        public const string GuidCrmDeveloperExtensionsPkgString = "EDBA1509-9962-4FDB-B52D-1B5CA2154DD2";

        //Item Templates
        public const string GuidItMenuCommandsCmdSetString = "393DA428-4DEC-489D-9BC7-586DD3DEAE24";
        public static readonly Guid GuidItMenuCommandsCmdSet = new Guid(GuidItMenuCommandsCmdSetString);

        //Web Resource Deployer
        //public const string GuidWebResourceDeployerCmdSetString = "86128E89-8A56-4A17-9029-A75379A89B9F";
        public const string GuidWrdToolWindowPersistanceString = "96AA3696-8674-484F-A95E-08355D14A7FB";
        //public static readonly Guid GuidWebResourceDeployerCmdSet = new Guid(GuidWebResourceDeployerCmdSetString);

        //Report Deployer
        //public const string GuidReportDeployerCmdSetString = "4752ABFB-0B1B-41AF-ADBD-47231C81F353";
        public const string GuidReportToolWindowPersistanceString = "F9FE1738-5BBA-4234-B1A0-5FF31833020B";
        //public static readonly Guid GuidReportDeployerCmdSet = new Guid(GuidReportDeployerCmdSetString);

        //Plug-in Deployer
        //public const string GuidPluginDeployerCmdSetString = "CF3C6DB1-0999-42BF-BE3A-B9F4E6542F1E";
        public const string GuidPlugunToolWindowPersistanceString = "F4C786FD-06CB-458E-94E1-D5D55B9FA9D7";
        //public static readonly Guid GuidPluginDeployerCmdSet = new Guid(GuidPluginDeployerCmdSetString);


        //CRM Developer Extensions Project Menu
        public const string GuidCrmDevExCmdSetString = "95CD7B0B-0592-4683-B42C-A79A41380FFE";
        public static readonly Guid GuidCrmDevExCmdSet = new Guid(GuidCrmDevExCmdSetString);
    };
}