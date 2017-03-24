using System;
using System.IO;

namespace ExcelDna.AddIn.Tasks.Utils
{
    internal class ExcelDnaPhysicalFileSystem : IExcelDnaFileSystem
    {
        public bool DirectoryExists(string path)
        {
            return Directory.Exists(path);
        }

        public bool FileExists(string path)
        {
            return File.Exists(path);
        }

        public void CreateDirectory(string path)
        {
            Directory.CreateDirectory(path);
        }

        public void CopyFile(string sourceFileName, string destFileName, bool overwrite)
        {
            if (overwrite)
            {
                var fileInfo = new FileInfo(destFileName);
                if (fileInfo.Exists && fileInfo.IsReadOnly)
                {
                    fileInfo.IsReadOnly = false;
                }
            }

            File.Copy(sourceFileName, destFileName, overwrite);
        }

        public string GetRelativePath(string path, string workingDirectory = null)
        {
            workingDirectory = workingDirectory ?? Environment.CurrentDirectory;

            var result = string.Empty;
            int offset;

            if (path.StartsWith(workingDirectory))
            {
                return path.Substring(workingDirectory.Length + 1);
            }

            var baseDirs = workingDirectory.Split(new[] { ':', '\\', '/' });
            var fileDirs = path.Split(new[] { ':', '\\', '/' });

            if (baseDirs.Length <= 0 || fileDirs.Length <= 0 || baseDirs[0] != fileDirs[0])
            {
                return path;
            }

            for (offset = 1; offset < baseDirs.Length; offset++)
            {
                if (baseDirs[offset] != fileDirs[offset])
                {
                    break;
                }
            }

            for (var i = 0; i < baseDirs.Length - offset; i++)
            {
                result += "..\\";
            }

            for (var i = offset; i < fileDirs.Length - 1; i++)
            {
                result += fileDirs[i] + "\\";
            }

            result += fileDirs[fileDirs.Length - 1];

            return result;
        }
    }
}
