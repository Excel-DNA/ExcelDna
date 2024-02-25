using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using ExcelDna.Testing;

namespace ExcelDna.AddIn.Tasks.IntegrationRuntimeTests
{
    [ExcelTestSettings(OutOfProcess = true)]
    public class NET6CustomRuntimeConfigurationTests
    {
        public NET6CustomRuntimeConfigurationTests()
        {
            TestTargetAddIn.BuildAndRegister("NET6CustomRuntimeConfiguration");
        }

        [ExcelFact(Workbook = "")]
        public void NotYetPacked()
        {
            Range functionRange = ((Worksheet)Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=SayHello(\"world\")";
            Assert.Equal("world WebApplication Environment: Production.", functionRange.Value.ToString());
        }
    }
}
