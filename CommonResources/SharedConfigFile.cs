using System.IO;

namespace CommonResources
{
    public static class SharedConfigFile
    {
        public static bool IsConfigReadOnly(string path)
        {
            FileInfo file = new FileInfo(path);
            return file.IsReadOnly;
        }
    }
}
