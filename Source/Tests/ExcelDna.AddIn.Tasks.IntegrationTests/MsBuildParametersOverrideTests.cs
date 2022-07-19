using NUnit.Framework;

namespace ExcelDna.AddIn.Tasks.IntegrationTests
{
    [TestFixture]
    public class MsBuildParametersOverrideTests : IntegrationTestBase
    {
        [Test]
        public void ExcelDnaBuild_can_be_disabled_by_setting_RunExcelDnaBuild_to_false()
        {
            const string projectBasePath = @"SingleDnaFileDefaultSuffix\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SingleDnaFileDefaultSuffix.csproj /t:Build /p:Configuration=Release /p:RunExcelDnaBuild=false /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dna", new string[0]);
            AssertOutput(projectOutDir, "*.xll", new string[0]);
            AssertOutput(projectOutDir, "*.xll.config", new string[0]);
        }

        [Test]
        public void ExcelDnaPack_can_be_disabled_by_setting_RunExcelDnaPack_to_false()
        {
            const string projectBasePath = @"SingleDnaFileDefaultSuffix\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SingleDnaFileDefaultSuffix.csproj /t:Build /p:Configuration=Release /p:RunExcelDnaPack=false /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dna", "MyLibrary-AddIn.dna", "MyLibrary-AddIn64.dna");
            AssertOutput(projectOutDir, "*.xll", "MyLibrary-AddIn.xll", "MyLibrary-AddIn64.xll");
            AssertOutput(projectOutDir, "*.xll.config", "MyLibrary-AddIn.xll.config", "MyLibrary-AddIn64.xll.config");
        }

        [Test]
        public void AddIn_32_bit_creation_can_be_disabled_by_setting_ExcelDnaCreate32BitAddIn_to_false()
        {
            const string projectBasePath = @"SingleDnaFileDefaultSuffix\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SingleDnaFileDefaultSuffix.csproj /t:Build /p:Configuration=Release /p:ExcelDnaCreate32BitAddIn=false /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dna", "MyLibrary-AddIn64.dna");
            AssertOutput(projectOutDir, "*.xll", "MyLibrary-AddIn64.xll", @"publish\MyLibrary-AddIn64-packed.xll");

            AssertIdentical(projectBasePath + "MyLibrary-AddIn.dna", projectOutDir + "MyLibrary-AddIn64.dna");

            Assert64BitXll(projectOutDir + "MyLibrary-AddIn64.xll");
        }

        [Test]
        public void AddIn_64_bit_creation_can_be_disabled_by_setting_ExcelDnaCreate64BitAddIn_to_false()
        {
            const string projectBasePath = @"SingleDnaFileDefaultSuffix\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SingleDnaFileDefaultSuffix.csproj /t:Build /p:Configuration=Release /p:ExcelDnaCreate64BitAddIn=false /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dna", "MyLibrary-AddIn.dna");
            AssertOutput(projectOutDir, "*.xll", "MyLibrary-AddIn.xll", @"publish\MyLibrary-AddIn-packed.xll");

            AssertIdentical(projectBasePath + "MyLibrary-AddIn.dna", projectOutDir + "MyLibrary-AddIn.dna");

            Assert32BitXll(projectOutDir + "MyLibrary-AddIn.xll");
        }

        [Test]
        public void No_addin_gets_created_if_both_ExcelDnaCreate32BitAddIn_andExcelDnaCreate64BitAddIn_are_false()
        {
            const string projectBasePath = @"SingleDnaFileDefaultSuffix\";
            const string projectOutDir = projectBasePath + @"bin\Release\";

            Clean(projectOutDir);

            MsBuild(projectBasePath + "SingleDnaFileDefaultSuffix.csproj /t:Build /p:Configuration=Release /p:ExcelDnaCreate32BitAddIn=false /p:ExcelDnaCreate64BitAddIn=false /v:m " + MsBuildParam("OutputPath", @"bin\Release\"));

            AssertOutput(projectOutDir, "*.dna", new string[0]);
            AssertOutput(projectOutDir, "*.xll", new string[0]);
            AssertOutput(projectOutDir, "*.xll.config", new string[0]);
        }
    }
}
