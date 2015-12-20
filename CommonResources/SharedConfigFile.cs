using System.IO;
using EnvDTE;

namespace CommonResources
{
    public static class SharedConfigFile
    {
        public static bool IsConfigReadOnly(string path)
        {
            FileInfo file = new FileInfo(path);
            return file.IsReadOnly;
        }

        public static bool ConfigFileExists(Project project)
        {
            var path = Path.GetDirectoryName(project.FullName);
            return File.Exists(path + "/CRMDeveloperExtensions.config");
        }
    }
}
