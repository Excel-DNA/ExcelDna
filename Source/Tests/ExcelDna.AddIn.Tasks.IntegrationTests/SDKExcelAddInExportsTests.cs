using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SDKExcelAddInExportsTests : IntegrationTestBase
    {
        [Test]
        public void ExternalLibrary()
        {
            const string projectBasePath = @"SDKExcelAddInExports\";
            const string projectOutDir = projectBasePath + @"bin\Release";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKExcelAddInExports.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertFileContains(Path.Combine(projectOutDir, "SDKExcelAddInExports-AddIn64.dna"), "ExternalLibrary Path=\"SDKExcelAddInExplicitExports.dll\"");
        }
    }
}

