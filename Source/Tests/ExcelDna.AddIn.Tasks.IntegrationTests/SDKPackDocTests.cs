using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SDKPackDocTests : IntegrationTestBase
    {
        [Test]
        public void Pack()
        {
            const string projectBasePath = @"SDKPackDoc\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            // The first build on a clean project doesn't create .chm at all, but the second run creates .chm and publishes it: 
            for (int i = 0; i < 2; ++i)
                MsBuild(projectBasePath + "SDKPackDoc.csproj /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));
        }

        [Test]
        public void CustomFileName()
        {
            const string projectBasePath = @"SDKPackDoc\";
            const string projectOutDir = projectBasePath + @"bin\ReleaseCustomFileName\";

            Clean(projectOutDir);

            // The first build on a clean project doesn't create .chm at all, but the second run creates .chm and publishes it: 
            for (int i = 0; i < 2; ++i)
                MsBuild(projectBasePath + "SDKPackDoc.csproj /p:ExcelAddInFileName=MyCustomFileName /t:Restore,Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\ReleaseCustomFileName\"));
        }
    }
}
