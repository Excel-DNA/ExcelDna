using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class NET6DnaComServerTests : IntegrationTestBase
    {
        [Test]
        public void CreateTlb()
        {
            const string projectBasePath = @"NET6DnaComServer\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET6DnaComServer.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.tlb", "NET6DnaComServer.tlb");
        }
    }
}
