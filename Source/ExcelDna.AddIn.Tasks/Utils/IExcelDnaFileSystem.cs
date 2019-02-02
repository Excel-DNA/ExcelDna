namespace ExcelDna.AddIn.Tasks.Utils
{
    internal interface IExcelDnaFileSystem
    {
        bool DirectoryExists(string path);
        bool FileExists(string path);

        void CreateDirectory(string path);

        void CopyFile(string sourceFileName, string destinationFileName, bool overwrite);
        void DeleteFile(string sourceFileName);

        string GetRelativePath(string path, string workingDirectory = null);
    }
}
