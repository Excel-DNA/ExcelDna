using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SDKExcelDnaCreate32BitAddInFalseTests : IntegrationTestBase
    {
        [Test]
        public void SDKExcelDnaCreate32BitAddInFalseTest()
        {
            const string projectBasePath = @"SDKExcelDnaCreate32BitAddInFalse\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKExcelDnaCreate32BitAddInFalse.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dna", "SDKExcelDnaCreate32BitAddInFalse-AddIn64.dna");
        }
    }
}
