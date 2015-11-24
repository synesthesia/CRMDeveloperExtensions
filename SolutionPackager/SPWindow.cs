using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace SolutionPackager
{
    [Guid("5ebf470b-cd52-4a25-bcb2-4f37b176ce54")]
    public sealed class SPWindow : ToolWindowPane
    {
        public SPWindow() :
            base(null)
        {
            Caption = Resources.ToolWindowTitle;
            BitmapResourceID = 301;
            BitmapIndex = 1;
            Content = new SolutionList();
        }
    }
}
