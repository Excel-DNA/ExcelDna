using System.IO;
using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SharedOutDirProjectsTests : IntegrationTestBase
    {
        [Test]
        public void Multiple_AddIn_Projects_Built_To_The_Same_Directory_Should_Only_Clean_Themselves()
        {
            const string projectOneBasePath = @"SharedOutDirProjectOne\";
            const string projectTwoBasePath = @"SharedOutDirProjectTwo\";

            const string relativeProjectOutDir = @"..\out\SharedOutDir\";
            var projectOutDir = Path.Combine(projectOneBasePath, relativeProjectOutDir);

            MsBuild(projectOneBasePath + "SharedOutDirProjectOne.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", relativeProjectOutDir));
            MsBuild(projectTwoBasePath + "SharedOutDirProjectTwo.csproj /t:Build /p:Configuration=Release /p:ExcelDnaPublishPath=PublishTwo /v:m " + MsBuildParam("OutputPath", relativeProjectOutDir));
            MsBuild(projectTwoBasePath + "SharedOutDirProjectTwo.csproj /t:Clean /p:Configuration=Release /p:ExcelDnaPublishPath=PublishTwo /v:m " + MsBuildParam("OutputPath", relativeProjectOutDir));

            // The .DNA files, XLL + config files should remain for project one
            AssertFound(projectOutDir, "*.dna", "AddIn-One-x.dna", "AddIn-One-x64.dna");
            AssertFound(projectOutDir, "*.xll", "AddIn-One-x.xll", @"publish\AddIn-One-x-packed.xll", "AddIn-One-x64.xll", @"publish\AddIn-One-x64-packed.xll");
            AssertFound(projectOutDir, "*.xll.config", "AddIn-One-x64.xll.config", @"publish\AddIn-One-x64-packed.xll.config");

            AssertIdentical(projectOneBasePath + "AddIn-One-x64.dna", projectOutDir + "AddIn-One-x64.dna");
            AssertIdentical(projectOneBasePath + "AddIn-One-x64.config", projectOutDir + "AddIn-One-x64.xll.config");

            // Assert project two files have been removed
            AssertNotFound(projectOutDir + "AddIn-Two-x-packed.xll");
            AssertNotFound(projectOutDir + @"publish\AddIn-Two-x64-packed.xll.config");
            AssertNotFound(projectOutDir + @"publish\AddIn-Two-x64-packed.xll");
            AssertNotFound(projectOutDir + "AddIn-Two-x64.xll.config");
            AssertNotFound(projectOutDir + "AddIn-Two-x64.xll");
            AssertNotFound(projectOutDir + "AddIn-Two-x64.dna");
            AssertNotFound(projectOutDir + "AddIn-Two-x.xll");
            AssertNotFound(projectOutDir + "AddIn-Two-x.dna");
        }
    }
}

