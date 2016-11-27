using CommonResources.Models;
using EnvDTE;
using System.Linq;
using Microsoft.Xrm.Tooling.Connector;

namespace CommonResources
{
    public static class SharedWindow
    {
        public static void OpenCrmPage(string url, CrmConn selectedConnection, DTE dte)
        {
            if (selectedConnection == null) return;
            string connString = selectedConnection.ConnectionString;
            if (string.IsNullOrEmpty(connString)) return;

            string[] connParts = connString.Split(';');
            string urlPart = connParts.FirstOrDefault(s => s.ToUpper().StartsWith("URL="));
            if (!string.IsNullOrEmpty(urlPart))
            {
                string[] urlParts = urlPart.Split('=');
                string baseUrl = (urlParts[1].EndsWith("/")) ? urlParts[1] : urlParts[1] + "/";

                var props = dte.Properties["CRM Developer Extensions", "General"];
                bool useDefaultWebBrowser = (bool)props.Item("UseDefaultWebBrowser").Value;

                if (useDefaultWebBrowser) //User's default browser
                    System.Diagnostics.Process.Start(baseUrl + url);
                else //Internal VS browser
                    dte.ItemOperations.Navigate(baseUrl + url);
            }
        }

        public static CrmServiceClient GetCachedConnection(string connName, string connString, DTE dte)
        {
            var currentWrClient = SharedGlobals.GetGlobal(connName, dte);
            if (currentWrClient == null)
                return new CrmServiceClient(connString);

            return (CrmServiceClient)currentWrClient;
        }
    }
}
