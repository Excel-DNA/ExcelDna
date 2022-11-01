using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;
using ExcelDna.PackedResources.Logging;
using Process = System.Diagnostics.Process;

namespace ExcelDna.AddIn.Tasks.Utils
{
    internal class DevToolsEnvironment : IDevToolsEnvironment
    {
        private readonly IBuildLogger _log;

        private const string _visualStudioProcessName = "devenv";

        public DevToolsEnvironment(IBuildLogger log)
        {
            _log = log ?? throw new ArgumentNullException(nameof(log));
        }

        public EnvDTE.Project GetProjectByName(string projectName)
        {
            _log.Debug("Starting GetProjectByName");

            foreach (var dte in EnumerateDevToolsEnvironments())
            {
                if (!TryFindCurrentProject(dte, projectName, out var project))
                {
                    _log.Debug($"Project {projectName} was not inside instance of DTE {dte.Name}");
                    continue;
                }

                _log.Debug($"Found project {projectName} inside DTE {dte.Name}");
                return project;
            }

            return null;
        }

        private IEnumerable<Process> EnumerateVisualStudioProcesses()
        {
            var process = Process.GetCurrentProcess();

            if (process.ProcessName.ToLower().Contains(_visualStudioProcessName))
            {
                // We're being compiled directly by Visual Studio
                yield return process;
            }

            // We're being compiled by other tool (e.g. MSBuild) called from a Visual Studio instance
            // therefore, some Visual Studio instance is our parent process.

            // Because of nodeReuse, we can't guarantee that the parent process of our current process is the "right" Visual Studio
            // so we just have to go through them all, and try to find our project in one of the Visual Studio's that are open

            foreach (var visualStudioProcess in Process.GetProcessesByName(_visualStudioProcessName))
            {
                yield return visualStudioProcess;
            }
        }

        private IEnumerable<EnvDTE.DTE> EnumerateDevToolsEnvironments()
        {
            foreach (var visualStudioProcess in EnumerateVisualStudioProcesses())
            {
                _log.Debug($"Getting DTE for VS Process ID: {visualStudioProcess.Id}");

                var dte = GetDevToolsEnvironment(visualStudioProcess.Id);
                if (dte == null)
                {
                    _log.Debug($"Unable to get DTE for VS Process ID: {visualStudioProcess.Id}");
                    continue;
                }

                _log.Debug($"Got DTE instance {dte.Name} for VS Process ID: {visualStudioProcess.Id}");

                yield return dte;
            }
        }

        private EnvDTE.DTE GetDevToolsEnvironment(int processId)
        {
            object runningObject = null;

            IBindCtx bindCtx = null;
            IRunningObjectTable rot = null;
            IEnumMoniker enumMonikers = null;

            try
            {
                Marshal.ThrowExceptionForHR(CreateBindCtx(reserved: 0, ppbc: out bindCtx));
                bindCtx.GetRunningObjectTable(out rot);
                rot.EnumRunning(out enumMonikers);

                var moniker = new IMoniker[1];
                var numberFetched = IntPtr.Zero;

                while (enumMonikers.Next(1, moniker, numberFetched) == 0)
                {
                    var runningObjectMoniker = moniker[0];

                    string name = null;

                    try
                    {
                        runningObjectMoniker?.GetDisplayName(bindCtx, null, out name);
                    }
                    catch (UnauthorizedAccessException)
                    {
                        // Do nothing, there is something in the ROT that we do not have access to
                    }

                    var monikerRegex = new Regex(@"!VisualStudio.DTE\.\d+\.\d+\:" + processId, RegexOptions.IgnoreCase);
                    if (!string.IsNullOrEmpty(name) && monikerRegex.IsMatch(name))
                    {
                        Marshal.ThrowExceptionForHR(rot.GetObject(runningObjectMoniker, out runningObject));
                        break;
                    }
                }
            }
            finally
            {
                if (enumMonikers != null)
                {
                    Marshal.ReleaseComObject(enumMonikers);
                }

                if (rot != null)
                {
                    Marshal.ReleaseComObject(rot);
                }

                if (bindCtx != null)
                {
                    Marshal.ReleaseComObject(bindCtx);
                }
            }

            return runningObject as EnvDTE.DTE;
        }

        private static bool TryFindCurrentProject(EnvDTE.DTE dte, string projectName, out EnvDTE.Project project)
        {
            project = dte
                .Solution
                .Projects
                .OfType<EnvDTE.Project>()
                .SelectMany(GetProjectAndSubProjects)
                .SingleOrDefault(p =>
                    string.Compare(p.Name, projectName, StringComparison.OrdinalIgnoreCase) == 0);

            return project != null;
        }

        private static IEnumerable<EnvDTE.Project> GetProjectAndSubProjects(EnvDTE.Project project)
        {
            if (project.Kind == _vsProjectKindSolutionFolder)
            {
                return project.ProjectItems
                    .OfType<EnvDTE.ProjectItem>()
                    .Select(p => p.SubProject)
                    .Where(p => p != null)
                    .SelectMany(GetProjectAndSubProjects);
            }

            return new[] { project };
        }

		// Copied from EnvDTE80, instead of referencing it just because of this one string
		private const string _vsProjectKindSolutionFolder = "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}";
		
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);
    }
}
