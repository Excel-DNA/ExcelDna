using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SDKDnaComServerTests : IntegrationTestBase
    {
        [Test]
        public void CreateTlb()
        {
            const string projectBasePath = @"SDKDnaComServer\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKDnaComServer.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.tlb", "SDKDnaComServer.tlb");
        }
    }
}
