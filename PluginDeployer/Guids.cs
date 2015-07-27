// Guids.cs
// MUST match guids.h
using System;

namespace PluginDeployer
{
    static class GuidList
    {
        public const string GuidPluginDeployerPkgString = "684c769a-9dbe-4784-9386-8f4b169df3ff";
        public const string GuidPluginDeployerCmdSetString = "cf3c6db1-0999-42bf-be3a-b9f4e6542f1e";
        public const string GuidToolWindowPersistanceString = "f4c786fd-06cb-458e-94e1-d5d55b9fa9d7";
        public static readonly Guid GuidPluginDeployerCmdSet = new Guid(GuidPluginDeployerCmdSetString);
    };
}