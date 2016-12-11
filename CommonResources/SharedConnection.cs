using EnvDTE;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Tooling.Connector;
using OutputLogger;
using System;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace CommonResources
{
    public class SharedConnection
    {
        public static void SetCurrentConnection(CrmServiceClient client, string type, DTE dte)
        {
            Globals globals = dte.Globals;
            globals["CurrentConnection" + type] = client;
        }

        public static void ClearCurrentConnection(string type, DTE dte)
        {
            Globals globals = dte.Globals;
            globals["CurrentConnection" + type] = null;
        }

        public static CrmServiceClient GetCurrentConnection(string connString, string type, DTE dte)
        {
            try
            {
                Globals globals = dte.Globals;
                if (!globals.VariableExists["CurrentConnection" + type])
                    return CreateConnection(connString, type, dte);

                CrmServiceClient client = (CrmServiceClient)globals["CurrentConnection" + type];
                return client ?? CreateConnection(connString, type, dte);
            }
            catch (Exception ex)
            {
                dte.StatusBar.Clear();
                dte.StatusBar.Animate(false, vsStatusAnimation.vsStatusAnimationSync);
                Logger logger = new Logger();
                logger.WriteToOutputWindow("Error connecting to CRM: " + ex.Message, Logger.MessageType.Error);
                return null;
            }
        }

        private static CrmServiceClient CreateConnection(string connString, string type, DTE dte)
        {
            CrmServiceClient client = new CrmServiceClient(connString);
            WhoAmIRequest wRequest = new WhoAmIRequest();
            WhoAmIResponse wResponse = (WhoAmIResponse)client.Execute(wRequest);
            Logger logger = new Logger();
            logger.WriteToOutputWindow("Connected To CRM Organization: " + wResponse.OrganizationId, Logger.MessageType.Info);
            logger.WriteToOutputWindow("Version: " + client.ConnectedOrgVersion, Logger.MessageType.Info);

            SetCurrentConnection(client, type, dte);

            return client;
        }

        public static void EnableLogging(DTE dte)
        {
            var props = dte.Properties["CRM Developer Extensions", "General"];
            bool enableLogging = (bool)props.Item("EnableXrmToolingLogging").Value;

            if (enableLogging)
            {
                TraceControlSettings.TraceLevel = SourceLevels.All;
                string logPath = (string)props.Item("XrmToolingLogPath").Value;
                string fileName = "CRMDevExXrmToolingLog" + Regex.Replace(DateTime.Now.ToShortDateString(), "[^0-9]", String.Empty) + ".log";
                TraceControlSettings.AddTraceListener(new TextWriterTraceListener(logPath + "\\" + fileName));
            }
        }
    }
}
