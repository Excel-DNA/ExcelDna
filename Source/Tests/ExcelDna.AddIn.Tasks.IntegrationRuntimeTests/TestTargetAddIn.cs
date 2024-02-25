using System.IO;
using ExcelDna.Testing;

namespace ExcelDna.AddIn.Tasks.IntegrationRuntimeTests
{
    internal class TestTargetAddIn
    {
        public static void BuildAndRegister(string projectName)
        {
            string testTargetPath = Path.GetFullPath(@"..\..\..\..\ExcelDna.AddIn.Tasks.IntegrationTests.TestTarget");
            string outputPath = @"bin\Release\";
            string projectDirPath = Path.Combine(testTargetPath, projectName);
            string xllPath = Path.Combine(projectDirPath, outputPath, $"{projectName}-AddIn64.xll");

            MsBuild.Build(Path.Combine(projectDirPath, $"{projectName}.csproj") + " /t:Restore,Build /p:Configuration=Release /v:m " + MsBuild.Param("OutputPath", outputPath));

            Util.Application.RegisterXLL(xllPath);
        }
    }
}
