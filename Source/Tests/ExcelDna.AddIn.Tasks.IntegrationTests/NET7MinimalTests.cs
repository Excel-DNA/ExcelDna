using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class NET7MinimalTests : IntegrationTestBase
    {
        [Test]
        public void NET7MinimalTest()
        {
            const string projectBasePath = @"NET7Minimal\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET7Minimal.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertFileContains(Path.Combine(projectOutDir, "NET7Minimal-AddIn64.dna"), "RuntimeVersion=\"v7.0\"");
        }
    }
}
