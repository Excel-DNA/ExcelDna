namespace ExcelDna.AddIn.Tasks.Utils
{
    public interface IExcelDnaFileSystem
    {
        bool DirectoryExists(string path);
        bool FileExists(string path);

        void CreateDirectory(string path);

        void CopyFile(string sourceFileName, string destinationFileName, bool overwrite);

        string GetRelativePath(string path, string workingDirectory = null);
    }
}