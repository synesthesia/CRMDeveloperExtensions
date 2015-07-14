using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;

namespace OutputLogger
{
    public class Logger
    {
        private static Guid _windowsId = new Guid("C2A5E6D0-0EE6-498F-BE3E-FA96CAF3A928");
        private static IVsOutputWindowPane _customPane;
        static Logger()
        {
            CreateOutputWindow();
        }

        public enum MessageType
        {
            Error,
            Warning,
            Info
        }

        private static void CreateOutputWindow()
        {
            IVsOutputWindow outWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;

            string title = "CRM Developer Extensions";
            if (outWindow == null)
                return;

            outWindow.GetPane(ref _windowsId, out _customPane);
            if (_customPane != null) return; //Already exists

            outWindow.CreatePane(ref _windowsId, title, 1, 1);

            outWindow.GetPane(ref _windowsId, out _customPane);
        }

        public void DeleteOutputWindow()
        {
            IVsOutputWindow outWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;

            if (outWindow != null)
                outWindow.DeletePane(_windowsId);
        }

        public void WriteToOutputWindow(string message, MessageType type)
        {
            switch (type)
            {
                case MessageType.Error:
                    message = "Error: " + DateTime.Now + "  " + message;
                    break;
                case MessageType.Warning:
                    message = "Warning: " + DateTime.Now + "  " + message;
                    break;
                case MessageType.Info:
                    message = "Info: " + DateTime.Now + "  " + message;
                    break;
            }

            message = Environment.NewLine + message;

            _customPane.OutputString(message);
            _customPane.Activate();
        }
    }
}
