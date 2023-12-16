using System.IO;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace NET6RollForward
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            var thisAddInName = Path.GetFileName((string)XlCall.Excel(XlCall.xlGetName));
            var message = string.Format("Excel-DNA Add-In '{0}' loaded!", thisAddInName);
            message += Environment.NewLine + $".NET version {Environment.Version}.";

            MessageBox.Show(message, thisAddInName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void AutoClose()
        {
        }
    }
}
