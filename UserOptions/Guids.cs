using System;

namespace UserOptions
{
    static class GuidList
    {
        public const string GuidUserOptionsPkgString = "8F9D97A5-8FC7-461C-9A4F-0C4D2E03210F";
        public const string GuidUserOptionsCmdSetString = "044D3D3A-3D4C-42DB-A7B0-CE4D04A5B002";
        public static readonly Guid GuidUserOptionsCmdSet = new Guid(GuidUserOptionsCmdSetString);
    };
}