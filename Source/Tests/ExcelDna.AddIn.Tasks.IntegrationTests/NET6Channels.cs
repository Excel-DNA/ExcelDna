using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class NET6ChannelsTests : IntegrationTestBase
    {
        [Test]
        public void NET6ChannelsTest()
        {
            const string projectBasePath = @"NET6Channels\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET6Channels.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));
        }
    }
}
