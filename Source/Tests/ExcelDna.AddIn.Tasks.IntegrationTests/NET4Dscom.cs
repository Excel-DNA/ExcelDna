using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class NET4Dscom : IntegrationTestBase
    {
        [Test]
        public void CreateTlb()
        {
            const string projectBasePath = @"NET4Dscom\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET4Dscom.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.tlb", "NET4Dscom.tlb");
        }
    }
}
