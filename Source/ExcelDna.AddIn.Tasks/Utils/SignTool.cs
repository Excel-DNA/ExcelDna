using System;
using System.Diagnostics;
using ExcelDna.PackedResources.Logging;

namespace ExcelDna.AddIn.Tasks.Utils
{
    internal class SignTool
    {
        public static void Sign(string toolPath, string toolOptions, string target, IBuildLogger log)
        {
            _log = log;
            Process p = new Process
            {
                StartInfo =
                {
                    FileName = toolPath,
                    Arguments = $"sign {toolOptions} \"{target}\"",
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                }
            };

            p.OutputDataReceived += OnDataReceived;
            p.ErrorDataReceived += OnDataReceived;

            p.Start();
            p.BeginOutputReadLine();
            p.BeginErrorReadLine();

            p.WaitForExit();
            if (p.ExitCode != 0)
                throw new ApplicationException($"SignTool failed with exit code {p.ExitCode}.");
        }

        private static void OnDataReceived(object sender, DataReceivedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(e.Data))
                _log.Information(e.Data);
        }

        private static IBuildLogger _log;
    }
}
