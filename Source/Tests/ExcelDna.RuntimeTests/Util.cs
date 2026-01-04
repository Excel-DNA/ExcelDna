using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
    public class Util
    {
        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void WindowHandle()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];

            functionRange.Formula = "=MyWindowHandle()";
            Assert.True(functionRange.Value.ToString().StartsWith("My WindowHandle is "));
        }
    }
}
