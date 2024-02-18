using System.IO;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace NET6CustomRuntimeConfiguration
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            var thisAddInName = Path.GetFileName((string)XlCall.Excel(XlCall.xlGetName));
            var message = string.Format("Excel-DNA Add-In '{0}' loaded!", thisAddInName);

            var webApplication = WebApplication.Create();
            message += Environment.NewLine;
            message += Environment.NewLine;
            message += $"WebApplication Environment: {webApplication.Environment.EnvironmentName}.";

            MessageBox.Show(message, thisAddInName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void AutoClose()
        {
        }
    }
}
