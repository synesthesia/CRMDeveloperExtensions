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
        public const string GuidWrdToolWindowPersistanceString = "96AA3696-8674-484F-A95E-08355D14A7FB";

        //Report Deployer
        public const string GuidReportToolWindowPersistanceString = "F9FE1738-5BBA-4234-B1A0-5FF31833020B";

        //Plug-in Deployer
        public const string GuidPluginToolWindowPersistanceString = "F4C786FD-06CB-458E-94E1-D5D55B9FA9D7";

        //Solution Pacjager
        public const string GuidPackgerToolWindowPersistanceString = "5EBF470B-CD52-4A25-BCB2-4F37B176CE54";

        //CRM Developer Extensions Project Menu
        public const string GuidCrmDevExCmdSetString = "95CD7B0B-0592-4683-B42C-A79A41380FFE";
        public static readonly Guid GuidCrmDevExCmdSet = new Guid(GuidCrmDevExCmdSetString);
    };
}