using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class NET6DisableAssemblyContextUnloadTests : IntegrationTestBase
    {
        [Test]
        public void DisableUnload()
        {
            const string projectBasePath = @"NET6DisableAssemblyContextUnload\";
            const string projectOutDir = projectBasePath + @"bin\Release";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET6DisableAssemblyContextUnload.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertFileContains(Path.Combine(projectOutDir, "NET6DisableAssemblyContextUnload-AddIn64.dna"), "DisableAssemblyContextUnload=\"true\"");
        }
    }
}
