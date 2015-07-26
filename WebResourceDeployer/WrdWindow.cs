using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;

namespace WebResourceDeployer
{
    [Guid("96aa3696-8674-484f-a95e-08355d14a7fb")]
    public sealed class WrdWindow : ToolWindowPane
    {
        public WrdWindow()
            : base(null)
        {
            Caption = Resources.ToolWindowTitle;
            BitmapResourceID = 301;
            BitmapIndex = 1;
            Content = new WebResourceList();
        }
    }
}
