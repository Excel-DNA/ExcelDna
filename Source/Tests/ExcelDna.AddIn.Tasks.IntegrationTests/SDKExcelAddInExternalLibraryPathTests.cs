using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SDKExcelAddInExternalLibraryPathTests : IntegrationTestBase
    {
        [Test]
        public void SDKExcelAddInExternalLibraryPathTest()
        {
            const string projectBasePath = @"SDKExcelAddInExternalLibraryPath\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKExcelAddInExternalLibraryPath.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertFileContains(Path.Combine(projectOutDir, "SDKExcelAddInExternalLibraryPath-AddIn64.dna"), "Path=\"SDKExcelAddInName.dll\"");
        }
    }
}
