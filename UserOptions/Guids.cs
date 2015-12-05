using System;

namespace UserOptions
{
    internal static class GuidList
    {
        public const string GuidUserOptionsPkgString = "8F9D97A5-8FC7-461C-9A4F-0C4D2E03210F";
        public const string GuidSpUserOptionsPkgString = "1541F7B5-59B2-46D4-8AA4-EF97C2204BDB";
        public const string GuidPdUserOptionsPkgString = "AB16A8A3-737C-4D7C-BD5E-D227EC0BFD58";
        public const string GuidWrdUserOptionsPkgString = "2E765D6D-3A04-4D86-8956-ECEB5FB8B811";
        public const string GuidRdUserOptionsPkgString = "9436DAE2-E279-4203-BFA6-A7F964AF6C9F";
        public const string GuidUserOptionsCmdSetString = "044D3D3A-3D4C-42DB-A7B0-CE4D04A5B002";
        public static readonly Guid GuidUserOptionsCmdSet = new Guid(GuidUserOptionsCmdSetString);
    }
}