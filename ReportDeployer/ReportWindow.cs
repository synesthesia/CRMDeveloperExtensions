using Microsoft.VisualStudio.Shell;
using System.Runtime.InteropServices;

namespace ReportDeployer
{
    [Guid("f9fe1738-5bba-4234-b1a0-5ff31833020b")]
    public sealed class ReportWindow : ToolWindowPane
    {
        public ReportWindow() :
            base(null)
        {
            Caption = Resources.ToolWindowTitle;
            BitmapResourceID = 301;
            BitmapIndex = 1;
            Content = new ReportList();
        }
    }
}
