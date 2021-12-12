using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SDKMinimalWithoutDnaTests : IntegrationTestBase
    {
        [Test]
        public void SDKMinimalWithoutDnaTest()
        {
            const string projectBasePath = @"SDKMinimalWithoutDna\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKMinimalWithoutDna.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dna", "SDKMinimalWithoutDna-AddIn.dna", "SDKMinimalWithoutDna-AddIn64.dna");
        }
    }
}
