using System;

namespace ReportDeployer
{
    static class GuidList
    {
        public const string GuidReportDeployerPkgString = "078A8CFC-BA76-4BF3-9600-36D1B36CBA2E";

        public const string GuidMenuCommandsPkgString = "15213CBD-E72C-452F-8CD9-42D61665C4DC";
        public const string GuidMenuCommandsCmdSetString = "ADBAE3F9-2CD2-4EB8-ABF3-07A6C84D24DA";
        public static readonly Guid GuidItemMenuCommandsCmdSet = new Guid(GuidMenuCommandsCmdSetString);
    };
}