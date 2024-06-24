using System.IO;
using ExcelDna.Testing;

namespace ExcelDna.AddIn.Tasks.IntegrationRuntimeTests
{
    internal class TestTargetAddIn
    {
        public static void Build(string projectName, string id, string? parameters = null)
        {
            try
            {
                Directory.Delete(Path.Combine(GetProjectDir(projectName), GetOutputPath(id)), true);
            }
            catch (DirectoryNotFoundException)
            {
            }

            MsBuild.Build(Path.Combine(GetProjectDir(projectName), $"{projectName}.csproj") + $" /t:Restore,Build /p:Configuration=Release {parameters} /v:m " + MsBuild.Param("OutputPath", GetOutputPath(id)));
        }

        public static bool Register(string projectName, string id, string xllName)
        {
            string xllPath = Path.Combine(GetProjectDir(projectName), GetOutputPath(id), xllName);

            return Util.Application.RegisterXLL(xllPath);
        }

        private static string GetProjectDir(string projectName)
        {
            string testTargetPath = Path.GetFullPath(@"..\..\..\..\ExcelDna.AddIn.Tasks.IntegrationTests.TestTarget");
            return Path.Combine(testTargetPath, projectName);
        }

        private static string GetOutputPath(string id)
        {
            return $@"bin\{id}\Release\";
        }
    }
}
