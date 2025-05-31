using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
    public class NativeAOT
    {
        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsAOT)]
        public void Hello()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=NativeHello(\"world\")";
            Assert.Equal("Hello world!", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsAOT)]
        public void Sum()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=NativeSum(3, 4)";
            Assert.Equal("7", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsAOT)]
        public void AsyncTaskInstant()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=NativeAsyncTaskHello(\"world\", 0)";

            Assert.Equal("Hello native async task world", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTestsAOT)]
        public void DefaultAsyncReturnValue()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=NativeAsyncTaskHello(\"world\", 1000)";

            Assert.Equal(-2146826246, functionRange.Value); // #N/A
        }
    }
}
