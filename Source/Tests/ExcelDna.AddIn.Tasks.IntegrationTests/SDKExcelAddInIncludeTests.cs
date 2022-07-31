using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SDKExcelAddInIncludeTests : IntegrationTestBase
    {
        [Test]
        public void Property()
        {
            const string projectBasePath = @"SDKExcelAddInInclude\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKExcelAddInInclude.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertFileContains(Path.Combine(projectOutDir, "SDKExcelAddInInclude-AddIn64.dna"), "Path=\"SDKExcelAddInName.dll\"");
            AssertFileContains(Path.Combine(projectOutDir, "SDKExcelAddInInclude-AddIn64.dna"), "Path=\"SDKExcelAddInFileName.dll\"");
        }

        [Test]
        public void Wildcard()
        {
            const string projectBasePath = @"SDKExcelAddInIncludeWildcard\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKExcelAddInIncludeWildcard.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertFileContains(Path.Combine(projectOutDir, "SDKExcelAddInIncludeWildcard-AddIn64.dna"), "Path=\"SDKExcelAddInName.dll\"");
            AssertFileContains(Path.Combine(projectOutDir, "SDKExcelAddInIncludeWildcard-AddIn64.dna"), "Path=\"SDKExcelAddInFileName.dll\"");
        }
    }
}
