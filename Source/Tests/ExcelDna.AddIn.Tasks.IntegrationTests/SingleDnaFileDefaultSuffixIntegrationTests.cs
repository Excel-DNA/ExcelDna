using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SingleDnaFileDefaultSuffixIntegrationTests : IntegrationTestBase
    {
        [Test]
        public void A_project_with_a_single_dna_file_using_default_suffix_copies_source_dna_to_32_and_64_bit_variants_and_copies_corresponding_xll_files()
        {
            const string projectBasePath = @"SingleDnaFileDefaultSuffix\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SingleDnaFileDefaultSuffix.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dna", "MyLibrary-AddIn.dna", "MyLibrary-AddIn64.dna");
            AssertOutput(projectOutDir, "*.xll", "MyLibrary-AddIn.xll", "MyLibrary-AddIn-packed.xll", "MyLibrary-AddIn64.xll", "MyLibrary-AddIn64-packed.xll");
            AssertOutput(projectOutDir, "*.xll.config", "MyLibrary-AddIn.xll.config", "MyLibrary-AddIn-packed.xll.config", "MyLibrary-AddIn64.xll.config", "MyLibrary-AddIn64-packed.xll.config");

            AssertIdentical(projectBasePath + "MyLibrary-AddIn.dna", projectOutDir + "MyLibrary-AddIn.dna");
            AssertIdentical(projectBasePath + "MyLibrary-AddIn.dna", projectOutDir + "MyLibrary-AddIn64.dna");

            Assert32BitXll(projectOutDir + "MyLibrary-AddIn.xll");
            Assert64BitXll(projectOutDir + "MyLibrary-AddIn64.xll");

            AssertIdentical(projectBasePath + "App.config", projectOutDir + "MyLibrary-AddIn.xll.config");
            AssertIdentical(projectBasePath + "App.config", projectOutDir + "MyLibrary-AddIn-packed.xll.config");

            AssertIdentical(projectBasePath + "App.config", projectOutDir + "MyLibrary-AddIn64.xll.config");
            AssertIdentical(projectBasePath + "App.config", projectOutDir + "MyLibrary-AddIn64-packed.xll.config");
        }

        [Test]
        public void A_project_with_a_single_dna_file_using_default_suffix_can_be_built_more_than_once_without_any_error()
        {
            const string projectBasePath = @"SingleDnaFileDefaultSuffix\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SingleDnaFileDefaultSuffix.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));
            MsBuild(projectBasePath + "SingleDnaFileDefaultSuffix.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));
        }

        [Test]
        public void A_project_with_a_single_dna_file_using_default_suffix_gets_its_output_clean_via_clean_build_target()
        {
            const string projectBasePath = @"SingleDnaFileDefaultSuffix\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SingleDnaFileDefaultSuffix.csproj /t:Build /p:RunExcelDnaPack=false /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dna", "MyLibrary-AddIn.dna", "MyLibrary-AddIn64.dna");
            AssertOutput(projectOutDir, "*.xll", "MyLibrary-AddIn.xll", "MyLibrary-AddIn64.xll");
            AssertOutput(projectOutDir, "*.xll.config", "MyLibrary-AddIn.xll.config", "MyLibrary-AddIn64.xll.config");

            MsBuild(projectBasePath + "SingleDnaFileDefaultSuffix.csproj /t:Clean /p:Configuration=Release /v:m");

            AssertOutput(projectOutDir, "*.dna", new string[0]);
            AssertOutput(projectOutDir, "*.xll", new string[0]);
            AssertOutput(projectOutDir, "*.xll.config", new string[0]);
        }
    }
}
