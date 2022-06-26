using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class NET6Sign : IntegrationTestBase
    {
        [Test]
        public void Sign()
        {
            const string projectBasePath = @"NET6Sign\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET6Sign.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));
        }
    }
}
