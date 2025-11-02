using System.IO;
using ExcelDna.Integration;

namespace AOTAddIn
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            var thisAddInName = Path.GetFileName((string)XlCall.Excel(XlCall.xlGetName));
            var message = string.Format("Excel-DNA Add-In '{0}'", thisAddInName);

            ExcelDnaUtil.DynamicApplication.Set("Caption", message);
        }

        public void AutoClose()
        {
        }
    }
}
