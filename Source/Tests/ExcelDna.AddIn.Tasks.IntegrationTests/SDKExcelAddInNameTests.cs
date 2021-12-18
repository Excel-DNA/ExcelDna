using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SDKExcelAddInNameTests : IntegrationTestBase
    {
        [Test]
        public void SDKExcelAddInNameTest()
        {
            const string projectBasePath = @"SDKExcelAddInName\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKExcelAddInName.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertFileContains(Path.Combine(projectOutDir, "SDKExcelAddInName-AddIn64.dna"), "Name=\"MyName\"");
        }
    }
}
