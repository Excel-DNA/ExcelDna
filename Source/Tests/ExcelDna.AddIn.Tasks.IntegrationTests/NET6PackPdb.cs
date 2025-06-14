using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class NET6PackPdbTests : IntegrationTestBase
    {
        [Test]
        public void NET6PackPdbTest()
        {
            const string projectBasePath = @"NET6PackPdb\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET6PackPdb.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertFileContains(Path.Combine(projectOutDir, "NET6PackPdb-AddIn64.dna"), "ExternalLibrary Path=\"NET6PackPdb.dll\" ExplicitExports=\"false\" LoadFromBytes=\"false\" Pack=\"true\" IncludePdb=\"true\"");
            AssertFileContains(Path.Combine(projectOutDir, "NET6PackPdb-AddIn64.dna"), "Reference Path=\"CallStackLibrary.dll\" Pack=\"true\" IncludePdb=\"true\"");
        }
    }
}
