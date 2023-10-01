using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class NET6RollForwardTests : IntegrationTestBase
    {
        [Test]
        public void RollForward()
        {
            const string projectBasePath = @"NET6RollForward\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET6RollForward.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertFileContains(Path.Combine(projectOutDir, "NET6RollForward-AddIn64.dna"), "RollForward=\"LatestMajor\"");
        }
    }
}
