using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
#if DEBUG
    public class NativeAOT
    {
        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTestsAOT\bin\Debug\net8.0-windows\win-x64\publish\ExcelDna.AddIn.RuntimeTestsAOT")]
        public void Hello()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=NativeHello(\"world\")";
            Assert.Equal("Hello world!", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTestsAOT\bin\Debug\net8.0-windows\win-x64\publish\ExcelDna.AddIn.RuntimeTestsAOT")]
        public void Sum()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=NativeSum(3, 4)";
            Assert.Equal("7", functionRange.Value.ToString());
        }
    }
#endif
}
