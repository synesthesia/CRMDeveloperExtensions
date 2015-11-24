using System;

namespace SolutionPackager
{
    static class GuidList
    {
        public const string GuidSolutionPackagerPkgString = "3449A2B2-F2F1-4026-A531-7980C4BE9B6F";
        public const string GuidSolutionPackagerCmdSetString = "EB1218C1-3AC1-47A2-90A8-545DB3F340CA";
        public const string GuidToolWindowPersistanceString = "5EBF470B-CD52-4A25-BCB2-4F37B176CE54";

        public static readonly Guid GuidSolutionPackagerCmdSet = new Guid(GuidSolutionPackagerCmdSetString);
    };
}