using NUnit.Framework;
using System.IO;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class NET4UnpackTests : IntegrationTestBase
    {
        [Test]
        public void Unpack()
        {
            const string projectBasePath = @"NET4Unpack\";
            const string projectOutDir = projectBasePath + @"bin\Release\";
            const string publishDir = projectOutDir + @"publish";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "NET4Unpack.csproj /p:ExcelDnaUnpack=true /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertFound(publishDir, "*.dll", new string[] { "ExcelDna.ManagedHost.dll", "ExcelDna.Integration.dll", "ExcelDna.Loader.dll" });
        }
    }
}
