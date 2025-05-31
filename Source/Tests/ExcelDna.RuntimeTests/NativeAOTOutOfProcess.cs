using ExcelDna.Testing;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
    [ExcelTestSettings(OutOfProcess = true)]
    [Collection("OutOfProcess")]
    public class NativeAOTOutOfProcess
    {
        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsAOT)]
        public void AsyncTask()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=NativeAsyncTaskHello(\"world\", 200)";

                Automation.WaitFor(() => functionRange.Value?.ToString() == "Hello native async task world", 1000);

                Assert.Equal("Hello native async task world", functionRange.Value.ToString());
            }
        }
    }
}

