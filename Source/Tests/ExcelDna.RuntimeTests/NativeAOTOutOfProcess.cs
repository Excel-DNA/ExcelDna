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
            Runner.ExecuteWithRetryWhenExcelBusy(() =>
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=NativeAsyncTaskHello(\"world\", 200)";

                Automation.WaitFor(() => functionRange.Value?.ToString() == "Hello native async task world", 1000);

                Assert.Equal("Hello native async task world", functionRange.Value.ToString());
            });
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsAOT)]
        public void AsyncSleep()
        {
            Runner.ExecuteWithRetryWhenExcelBusy(() =>
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=NativeAsyncHello(\"world\", 0)";

                Automation.WaitFor(() => functionRange.Value?.ToString() == "Hello native async world", 1000);

                Assert.Equal("Hello native async world", functionRange.Value.ToString());
            });

            Runner.ExecuteWithRetryWhenExcelBusy(() =>
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange.Formula = "=NativeAsyncHello(\"world\", 200)";

                Automation.WaitFor(() => functionRange.Value?.ToString() == "Hello native async world", 2000);

                Assert.Equal("Hello native async world", functionRange.Value.ToString());
            });
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsAOT)]
        public void DynamicApplication()
        {
            Runner.ExecuteWithRetryWhenExcelBusy(() =>
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E1"];

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E2"];
                functionRange2.Formula = "=NativeApplicationAlignCellRight(\"E1\")";

                Assert.Equal(-4152, functionRange2.Value);
                Assert.Equal(-4152, (int)functionRange1.HorizontalAlignment);
            });
        }
    }
}

