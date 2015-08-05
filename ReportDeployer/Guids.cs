// Guids.cs
// MUST match guids.h
using System;

namespace ReportDeployer
{
    static class GuidList
    {
        public const string GuidReportDeployerPkgString = "078a8cfc-ba76-4bf3-9600-36d1b36cba2e";
        public const string GuidReportDeployerCmdSetString = "4752abfb-0b1b-41af-adbd-47231c81f353";
        public const string GuidToolWindowPersistanceString = "f9fe1738-5bba-4234-b1a0-5ff31833020b";
        public static readonly Guid GuidReportDeployerCmdSet = new Guid(GuidReportDeployerCmdSetString);

        public const string GuidMenuCommandsPkgString = "15213CBD-E72C-452F-8CD9-42D61665C4DC";
        public const string GuidMenuCommandsCmdSetString = "ADBAE3F9-2CD2-4EB8-ABF3-07A6C84D24DA";
        public static readonly Guid GuidItemMenuCommandsCmdSet = new Guid(GuidMenuCommandsCmdSetString);
    };
}