using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace NET6PackNativeInclude
{
    public class AddIn : IExcelAddIn
    {
        [DllImport("MyNativeLibrary.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int MySum(int a1, int a2);

        public void AutoOpen()
        {
            var thisAddInName = Path.GetFileName((string)XlCall.Excel(XlCall.xlGetName));
            try
            {
                var message = string.Format("Excel-DNA Add-In '{0}' loaded!", thisAddInName);
                message += Environment.NewLine + MySum(40, 2);
                MessageBox.Show(message, thisAddInName, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), thisAddInName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void AutoClose()
        {
        }
    }
}
