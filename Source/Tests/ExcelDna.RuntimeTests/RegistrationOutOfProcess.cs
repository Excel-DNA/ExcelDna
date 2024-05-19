using ExcelDna.Testing;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
#if DEBUG
    [ExcelTestSettings(OutOfProcess = true)]
    [Collection("OutOfProcess")]
    public class RegistrationOutOfProcess
    {
        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn")]
        public void AsyncSleep()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=MyAsyncHello(\"world\", 0)";

                Automation.WaitFor(() => functionRange.Value?.ToString() == "Hello async world", 1000);

                Assert.Equal("Hello async world", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange.Formula = "=MyAsyncHello(\"world\", 200)";

                Automation.WaitFor(() => functionRange.Value?.ToString() == "Hello async world", 1000);

                Assert.Equal("Hello async world", functionRange.Value.ToString());
            }
        }
    }
#endif
}
