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

        [DllImport("MyNativeLibrary.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int MySum(int a1, int a2);

        public void AutoOpen()
        {
            var thisAddInName = Path.GetFileName((string)XlCall.Excel(XlCall.xlGetName));
            var message = string.Format("Excel-DNA Add-In '{0}' loaded!", thisAddInName);
            message += Environment.NewLine + MySum(40, 2);

            DebugBreak();
        }

        public void AutoClose()
        {
        }
    }
}
