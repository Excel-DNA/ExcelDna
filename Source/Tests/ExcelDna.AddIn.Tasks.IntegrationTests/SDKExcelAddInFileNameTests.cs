using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SDKExcelAddInFileNameTests : IntegrationTestBase
    {
        [Test]
        public void SDKExcelAddInFileNameTest()
        {
            const string projectBasePath = @"SDKExcelAddInFileName\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKExcelAddInFileName.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dna", "MyFile.dna", "MyFile64.dna");
        }
    }
}
