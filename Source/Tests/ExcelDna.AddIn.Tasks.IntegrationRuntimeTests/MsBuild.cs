using Microsoft.VisualStudio.Setup.Configuration;
using System.IO;
using System.Text;

namespace ExcelDna.AddIn.Tasks.IntegrationRuntimeTests
{
    internal class MsBuild
    {
        public static void Build(string commandLineArguments)
        {
            var allOutput = new StringBuilder();

            Action<string> writer = output =>
            {
                allOutput.AppendLine(output);
                Console.WriteLine(output);
            };

            var msBuildPath = GetMsBuildPath();

            var result = SilentProcessRunner.ExecuteCommand(msBuildPath, commandLineArguments, Environment.CurrentDirectory, writer, e => writer("ERROR: " + e));
            if (result != 0)
            {
                Assert.Fail($"MSBuild returned a non-zero exit code: {result}\n\r{allOutput}");
            }
        }

        public static string Param(string name, string value)
        {
            return string.Format("/p:{0}=\"{1}\\\"", name, value);
        }

        private static string GetMsBuildPath()
        {
            try
            {
                var vsConfiguration = new SetupConfiguration();
                var vsInstancesEnumerator = vsConfiguration.EnumAllInstances();
                int fetched;
                var vsInstances = new ISetupInstance[1];
                do
                {
                    vsInstancesEnumerator.Next(1, vsInstances, out fetched);
                    if (fetched > 0)
                    {
                        var vsInstance = vsInstances[0];
                        if (vsInstance.GetInstallationVersion().StartsWith("17.")) // Visual Studio 2022
                            return vsInstance.ResolvePath(@"Msbuild\Current\Bin\amd64\MSBuild.exe");
                    }
                }
                while (fetched > 0);
            }
            catch (Exception)
            {
            }

            string msBuildPath;

            var programFilesDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            foreach (var version in new[] { "14.0", "12.0" })
            {
                var buildDirectory = Path.Combine(programFilesDirectory, "MSBuild", version, "Bin");
                msBuildPath = Path.Combine(buildDirectory, "msbuild.exe");

                if (File.Exists(msBuildPath))
                {
                    return msBuildPath;
                }
            }

            var netFx = System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory();

            msBuildPath = Path.Combine(netFx, "msbuild.exe");
            if (!File.Exists(msBuildPath))
            {
                Assert.Fail("Could not find MSBuild at: " + msBuildPath);
            }

            return msBuildPath;
        }
    }
}
