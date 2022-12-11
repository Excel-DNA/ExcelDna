using System.IO;

namespace ExcelDna.ManagedHost
{
    internal class TempDir
    {
        public TempDir(string topDirName)
        {
            path = Path.Combine(Path.GetTempPath(), topDirName, System.Guid.NewGuid().ToString());
            Directory.CreateDirectory(path);
        }

        public string GetPath()
        {
            return path;
        }

        private string path;
    }
}
