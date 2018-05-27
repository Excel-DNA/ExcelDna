using System;

namespace ExcelDna.AddIn.Tasks.Utils
{
    public interface IExcelDnaProject
    {
        bool TrySetDebuggerOptions(string projectName, string excelExePath, string excelAddInToDebug);
    }
}
