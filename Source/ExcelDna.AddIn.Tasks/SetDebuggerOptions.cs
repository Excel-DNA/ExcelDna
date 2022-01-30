using System;
using System.IO;
using System.Linq;
using ExcelDna.AddIn.Tasks.Logging;
using Microsoft.Build.Framework;
using ExcelDna.AddIn.Tasks.Utils;

namespace ExcelDna.AddIn.Tasks
{
    [RunInSTA]
    public class SetDebuggerOptions : ExcelDebugTask
    {
        private readonly Lazy<IExcelDnaProject> _projectLazy;
        private IExcelDnaProject _project;
        private readonly TimeSpan _defaultSynchronizationTimeout = TimeSpan.FromMinutes(1);

        public SetDebuggerOptions()
        {
            _log = new BuildLogger(this, "ExcelDnaSetDebuggerOptions");

            // We can only start logging when `Execute` runs... Until then, BuildEngine is `null`
            // Which is why we're deferring creating instances of ExcelDetector and ExcelDnaProject
            _excelDetectorLazy = new Lazy<IExcelDetector>(() => new ExcelDetector(_log));
            _projectLazy = new Lazy<IExcelDnaProject>(() => new ExcelDnaProject(_log, new DevToolsEnvironment(_log)));
        }

        internal SetDebuggerOptions(IBuildLogger log, IExcelDetector excelDetector, IExcelDnaProject project)
        {
            if (excelDetector == null) throw new ArgumentNullException(nameof(excelDetector));
            if (project == null) throw new ArgumentNullException(nameof(project));

            _log = log ?? throw new ArgumentNullException(nameof(log));
            _excelDetectorLazy = new Lazy<IExcelDetector>(() => excelDetector);
            _projectLazy = new Lazy<IExcelDnaProject>(() => project);
        }

        public override bool Execute()
        {
            try
            {
                _log.Debug("Running SetDebuggerOptions MSBuild Task");

                // Create instances of ExcelDetector and ExcelDnaProject
                _excelDetector = _excelDetectorLazy.Value;
                _project = _projectLazy.Value;

                FilesInProject = FilesInProject ?? new ITaskItem[0];
                _log.Debug("Number of files in project: " + FilesInProject.Length);

                var excelExePath = GetExcelPath();
                var addInForDebugging = GetAddInForDebugging(excelExePath);

                LogDiagnostics();

                if (!_project.TrySetDebuggerOptions(ProjectName, excelExePath, addInForDebugging))
                {
                    const string message = "Unable to set the debugger options within Visual Studio. " +
                                           "Please restart Visual Studio and try again.";
                    _log.Warning("DNA" + "PROJECT".GetHashCode(), message);
                }
            }
            catch (Exception ex)
            {
                _log.Warning(ex, ex.Message);
            }

            // Setting the debugger options is not essential to the build process, thus if anything
            // goes wrong, we'll report errors and warnings, but will not fail the build because of that
            return true;
        }

        private void LogDiagnostics()
        {
            _log.Debug("----Arguments----");
            _log.Debug("ProjectName: " + ProjectName);
            _log.Debug("ExcelExePath: " + ExcelExePath);
            _log.Debug("AddInForDebugging: " + AddInForDebugging);
            _log.Debug("FilesInProject: " + (FilesInProject ?? new ITaskItem[0]).Length);

            if (FilesInProject != null)
            {
                foreach (var f in FilesInProject)
                {
                    _log.Debug($"  {f.ItemSpec}");
                }
            }

            _log.Debug("OutDirectory: " + OutDirectory);
            _log.Debug("FileSuffix32Bit: " + FileSuffix32Bit);
            _log.Debug("FileSuffix64Bit: " + FileSuffix64Bit);
            _log.Debug("-----------------");
        }
    }
}
