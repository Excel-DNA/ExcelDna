namespace ExcelDna.AddIn.Tasks.Utils
{
    internal interface IExcelDnaProject
    {
        bool TrySetDebuggerOptions(string projectName, string excelExePath, string excelAddInToDebug);
    }
}
