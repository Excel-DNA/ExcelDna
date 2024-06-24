using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using ExcelDna.Testing;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationRuntimeTests
{
    public class NET6CustomRuntimeConfigurationTestsFixture
    {
        public NET6CustomRuntimeConfigurationTestsFixture()
        {
            TestTargetAddIn.Build(NET6CustomRuntimeConfigurationTests.ProjectName, NET6CustomRuntimeConfigurationTests.PackedId);
            TestTargetAddIn.Build(NET6CustomRuntimeConfigurationTests.ProjectName, NET6CustomRuntimeConfigurationTests.UnpackedId, "/p:ExcelDnaUnpack=true");
        }
    }

    [ExcelTestSettings(OutOfProcess = true)]
    public class NET6CustomRuntimeConfigurationTests : IClassFixture<NET6CustomRuntimeConfigurationTestsFixture>
    {
        [ExcelFact(Workbook = "")]
        public void NotYetPacked()
        {
            Assert.True(TestTargetAddIn.Register(ProjectName, PackedId, "NET6CustomRuntimeConfiguration-AddIn64.xll"), "Registration failed.");
            TestSayHello();
        }

        [ExcelFact(Workbook = "")]
        public void Packed()
        {
            Assert.True(TestTargetAddIn.Register(ProjectName, PackedId, Path.Combine("publish", "NET6CustomRuntimeConfiguration-AddIn64-packed.xll")), "Registration failed.");
            TestSayHello();
        }

        [ExcelFact(Workbook = "")]
        public void Unpacked()
        {
            Assert.True(TestTargetAddIn.Register(ProjectName, UnpackedId, Path.Combine("publish", "NET6CustomRuntimeConfiguration-AddIn64.xll")), "Registration failed.");
            TestSayHello();
        }

        private void TestSayHello()
        {
            Range functionRange = ((Worksheet)Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=SayHello(\"world\")";
            Assert.Equal("world WebApplication Environment: Production.", functionRange.Value.ToString());
        }

        internal const string ProjectName = "NET6CustomRuntimeConfiguration";
        internal const string PackedId = "Packed";
        internal const string UnpackedId = "Unpacked";
    }
}
