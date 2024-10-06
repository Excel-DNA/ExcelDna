using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class NET8RuntimeFrameworkVersionTests : IntegrationTestBase
    {
        [Test]
        public void RuntimeFrameworkVersion()
        {
            const string projectBasePath = @"NET8RuntimeFrameworkVersion\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET8RuntimeFrameworkVersion.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertFileContains(Path.Combine(projectOutDir, "NET8RuntimeFrameworkVersion-AddIn64.dna"), "RuntimeFrameworkVersion=\"8.0.8\"");
        }
    }
}
