using System;

namespace PluginDeployer
{
    static class GuidList
    {
        public const string GuidPluginDeployerPkgString = "684C769A-9DBE-4784-9386-8F4B169DF3FF";

        public const string GuidMenuCommandsPkgString = "CF3C6DB1-0999-42BF-BE3A-B9F4E6542F1E";
        public const string GuidMenuCommandsCmdSetString = "F4C786FD-06CB-458E-94E1-D5D55B9FA9D7";
        public static readonly Guid GuidProjectMenuCommandsCmdSet = new Guid(GuidMenuCommandsCmdSetString);
    };
}