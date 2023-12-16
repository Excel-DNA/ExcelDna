using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class NET8MinimalTests : IntegrationTestBase
    {
        [Test]
        public void NET8MinimalTest()
        {
            const string projectBasePath = @"NET8Minimal\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET8Minimal.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertFileContains(Path.Combine(projectOutDir, "NET8Minimal-AddIn64.dna"), "RuntimeVersion=\"v8.0\"");
        }
    }
}
