using ExcelDna.Integration;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace AOTPackPdb
{
    public class AddIn : IExcelAddIn
    {
        [DllImport("kernel32.dll")]
        static extern void DebugBreak();

        public void AutoOpen()
        {
            var thisAddInName = Path.GetFileName((string)XlCall.Excel(XlCall.xlGetName));
            var message = string.Format("Excel-DNA Add-In '{0}' loaded!", thisAddInName);

            DebugBreak();
        }

        public void AutoClose()
        {
        }
    }
}
