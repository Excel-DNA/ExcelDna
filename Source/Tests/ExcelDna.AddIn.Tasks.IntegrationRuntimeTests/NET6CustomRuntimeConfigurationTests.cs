using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using ExcelDna.Testing;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationRuntimeTests
{
    [ExcelTestSettings(OutOfProcess = true)]
    public class NET6CustomRuntimeConfigurationTests
    {
        [ExcelFact(Workbook = "")]
        public void NotYetPacked()
        {
            TestTargetAddIn.Build(projectName);
            Assert.True(TestTargetAddIn.Register(projectName, "NET6CustomRuntimeConfiguration-AddIn64.xll"), "Registration failed.");
            TestSayHello();
        }

        [ExcelFact(Workbook = "")]
        public void Packed()
        {
            TestTargetAddIn.Build(projectName);
            Assert.True(TestTargetAddIn.Register(projectName, Path.Combine("publish", "NET6CustomRuntimeConfiguration-AddIn64-packed.xll")), "Registration failed.");
            TestSayHello();
        }

        [ExcelFact(Workbook = "")]
        public void Unpacked()
        {
            TestTargetAddIn.Build(projectName, "/p:ExcelDnaUnpack=true");
            Assert.True(TestTargetAddIn.Register(projectName, Path.Combine("publish", "NET6CustomRuntimeConfiguration-AddIn64.xll")), "Registration failed.");
            TestSayHello();
        }

        private void TestSayHello()
        {
            Range functionRange = ((Worksheet)Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=SayHello(\"world\")";
            Assert.Equal("world WebApplication Environment: Production.", functionRange.Value.ToString());
        }

        private const string projectName = "NET6CustomRuntimeConfiguration";
    }
}
