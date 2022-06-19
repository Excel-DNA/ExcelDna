using ExcelDna.AddIn.Tasks.Logging;
using System;
using System.Diagnostics;
using System.IO;

namespace ExcelDna.AddIn.Tasks.Utils
{
    internal class TlbExp
    {
        public static void Create(string toolPath, string outputFile, IBuildLogger log)
        {
            _log = log;
            Process p = new Process
            {
                StartInfo =
                {
                    FileName = toolPath,
                    Arguments = $"\"{outputFile}\" /out:\"{Path.ChangeExtension(outputFile,"tlb")}\"",
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
                throw new ApplicationException($"TlbExp failed with exit code {p.ExitCode}.");
        }

        private static void OnDataReceived(object sender, DataReceivedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(e.Data))
                _log.Information(e.Data);
        }

        private static IBuildLogger _log;
    }
}
