using ExcelDna.AddIn.Tasks.Logging;
using System;
using System.Diagnostics;

namespace ExcelDna.AddIn.Tasks.Utils
{
    internal class ProcessRunner
    {
        public static void Run(string fileName, string arguments, string appName, IBuildLogger log)
        {
            _log = log;
            Process p = new Process
            {
                StartInfo =
                {
                    FileName = fileName,
                    Arguments = arguments,
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
                throw new ApplicationException($"{appName} failed with exit code {p.ExitCode}.");
        }

        private static void OnDataReceived(object sender, DataReceivedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(e.Data))
                _log.Information(e.Data);
        }

        private static IBuildLogger _log;
    }
}
