using ExcelDna.Testing;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
#if DEBUG
    [ExcelTestSettings(OutOfProcess = true)]
    [Collection("OutOfProcess")]
    public class RegistrationSampleOutOfProcessFS
    {
        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExcelDna.AddIn.RegistrationSampleFS\bin\Debug\net6.0-windows\ExcelDna.AddIn.RegistrationSampleFS-AddIn")]
        public void AsyncSleep()
        {
            {
                Range functionRangeA = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A3"];
                functionRangeA.Value = "alice";

                Range functionRangeB = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B3"];
                functionRangeB.Value = "1000";

                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C3"];
                functionRange.Formula = "=dnaFsHelloAsync(A3,B3)";
                Assert.Equal(-2146826246, functionRange.Value); // #N/A

                Automation.WaitFor(() => functionRange.Value?.ToString() == "Hello alice", 3000);

                Assert.Equal("Hello alice", functionRange.Value.ToString());
            }
            {
                Range functionRangeA = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A4"];
                functionRangeA.Value = "bob";

                Range functionRangeB = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B4"];
                functionRangeB.Value = "2000";

                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C4"];
                functionRange.Formula = "=dnaFsHelloAsync(A4,B4)";
                Assert.Equal(-2146826246, functionRange.Value); // #N/A

                Automation.WaitFor(() => functionRange.Value?.ToString() == "Hello bob", 3000);

                Assert.Equal("Hello bob", functionRange.Value.ToString());
            }
        }
    }
#endif
}
