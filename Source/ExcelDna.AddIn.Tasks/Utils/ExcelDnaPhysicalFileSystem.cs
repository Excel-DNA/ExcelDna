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
            if (string.Equals(Path.GetFullPath(sourceFileName), Path.GetFullPath(destFileName), StringComparison.OrdinalIgnoreCase))
                return;

            if (overwrite)
            {
                var destFileInfo = new FileInfo(destFileName);
                if (destFileInfo.Exists && destFileInfo.IsReadOnly)
                {
                    destFileInfo.IsReadOnly = false;
                }
            }

            var outputFileMode = overwrite ? FileMode.Create : FileMode.CreateNew;

            using (var inputStream = new FileStream(sourceFileName, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var outputStream = new FileStream(destFileName, outputFileMode, FileAccess.Write, FileShare.None))
            {
                const int bufferSize = 16384; // 16 Kb
                var buffer = new byte[bufferSize];

                int bytesRead;

                do
                {
                    bytesRead = inputStream.Read(buffer, 0, bufferSize);
                    outputStream.Write(buffer, 0, bytesRead);
                } while (bytesRead == bufferSize);
            }
        }

        public void WriteFile(string sourceText, string destinationFileName)
        {
            File.WriteAllText(destinationFileName, sourceText);
        }

        public void DeleteFile(string sourceFileName)
        {
            var fileInfo = new FileInfo(sourceFileName);
            if (fileInfo.IsReadOnly)
            {
                fileInfo.IsReadOnly = false;
            }

            File.Delete(sourceFileName);
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
