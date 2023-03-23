using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SDKPublishDocTests : IntegrationTestBase
    {
        [Test]
        [Ignore("Requires ExcelDnaDoc package that creates .chm before addin publish.")]
        public void Packed()
        {
            const string projectBasePath = @"SDKPublishDoc\";
            const string projectOutDir = projectBasePath + @"bin\Release\";
            const string publishDir = projectOutDir + @"publish\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SDKPublishDoc.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(publishDir, "*.chm", "SDKPublishDoc-AddIn.chm");
        }
    }
}
