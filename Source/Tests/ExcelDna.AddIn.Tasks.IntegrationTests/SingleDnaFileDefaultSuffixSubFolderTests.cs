using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class SingleDnaFileDefaultSuffixSubFolderTests : IntegrationTestBase
    {
        [Test]
        public void A_project_with_a_single_dna_file_in_a_subfolder_using_default_suffix_copies_source_dna_to_32_and_64_bit_variants_and_copies_corresponding_xll_files_keeping_folder_structure()
        {
            const string projectBasePath = @"SingleDnaFileInSubFolderDefaultSuffix\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SingleDnaFileInSubFolderDefaultSuffix.csproj /t:Build /p:Configuration=Release /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir + @"MySubFolder\", "*.dna", "MyLibrary-AddIn.dna", "MyLibrary-AddIn64.dna");
            AssertOutput(projectOutDir + @"MySubFolder\", "*.xll", "MyLibrary-AddIn.xll", @"publish\MyLibrary-AddIn-packed.xll", "MyLibrary-AddIn64.xll", @"publish\MyLibrary-AddIn64-packed.xll");
            AssertOutput(projectOutDir + @"MySubFolder\", "*.xll.config", "MyLibrary-AddIn.xll.config", @"publish\MyLibrary-AddIn-packed.xll.config", "MyLibrary-AddIn64.xll.config", @"publish\MyLibrary-AddIn64-packed.xll.config");

            AssertIdentical(projectBasePath + @"MySubFolder\MyLibrary-AddIn.dna", projectOutDir + @"MySubFolder\MyLibrary-AddIn.dna");
            AssertIdentical(projectBasePath + @"MySubFolder\MyLibrary-AddIn.dna", projectOutDir + @"MySubFolder\MyLibrary-AddIn64.dna");

            Assert32BitXll(projectOutDir + @"MySubFolder\MyLibrary-AddIn.xll");
            Assert64BitXll(projectOutDir + @"MySubFolder\MyLibrary-AddIn64.xll");

            AssertIdentical(projectBasePath + "App.config", projectOutDir + @"MySubFolder\MyLibrary-AddIn.xll.config");
            AssertIdentical(projectBasePath + "App.config", projectOutDir + @"MySubFolder\publish\MyLibrary-AddIn-packed.xll.config");

            AssertIdentical(projectBasePath + "App.config", projectOutDir + @"MySubFolder\MyLibrary-AddIn64.xll.config");
            AssertIdentical(projectBasePath + "App.config", projectOutDir + @"MySubFolder\publish\MyLibrary-AddIn64-packed.xll.config");
        }
    }
}
