using ExcelDna.Testing;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
#if DEBUG
    [ExcelTestSettings(OutOfProcess = true)]
    [Collection("OutOfProcess")]
    public class RegistrationSampleOutOfProcess
    {
        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RegistrationSample\bin\Debug\net6.0-windows\ExcelDna.AddIn.RegistrationSample-AddIn")]
        public void AsyncSleep()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=dnaDelayedHello(\"world\", 0)";

                Automation.WaitFor(() => functionRange.Value?.ToString() == "Hello world!", 1000);

                Assert.Equal("Hello world!", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange.Formula = "=dnaDelayedHello(\"world\", 200)";

                Automation.WaitFor(() => functionRange.Value?.ToString() == "Hello world!", 1000);

                Assert.Equal("Hello world!", functionRange.Value.ToString());
            }
        }
    }
#endif
}
