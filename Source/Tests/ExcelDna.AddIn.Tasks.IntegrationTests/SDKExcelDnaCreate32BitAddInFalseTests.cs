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

            MsBuild(projectBasePath + "SDKExcelDnaCreate32BitAddInFalse.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dna", "SDKExcelDnaCreate32BitAddInFalse-AddIn64.dna");
        }

        [Test]
        public void No64BitAddInSuffix()
        {
            const string projectBasePath = @"SDKExcelDnaCreate32BitAddInFalse\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKExcelDnaCreate32BitAddInFalse.csproj /t:Restore,Build /p:ExcelDna64BitAddInSuffix=%none% /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dna", "SDKExcelDnaCreate32BitAddInFalse-AddIn.dna");
        }
    }
}
