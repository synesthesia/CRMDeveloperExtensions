using System;

namespace WebResourceDeployer
{
    static class GuidList
    {
        public const string GuidWebResourceDeployerPkgString = "E42F2686-2A9B-40C1-8D13-9AB90D3B951D";

        public const string GuidMenuCommandsPkgString = "67F5E05C-872F-4637-A050-F0C66252CEC2";
        public const string GuidMenuCommandsCmdSetString = "44EC1EE2-36CB-4D17-B951-896254C73D35";
        public static readonly Guid GuidItemMenuCommandsCmdSet = new Guid(GuidMenuCommandsCmdSetString);

        public const string GuidEditorCommandsPkgString = "2723E9D7-8962-4F84-AA37-8AE4EC58F04E";
        public const string GuidEditorCommandsCmdSetString = "D3593638-DF55-428F-9570-4FAA63E7CB3A";
        public static readonly Guid GuidEditorCommandsCmdSet = new Guid(GuidEditorCommandsCmdSetString);
    };
}