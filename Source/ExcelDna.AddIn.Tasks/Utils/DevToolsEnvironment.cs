using System;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;

namespace ExcelDna.AddIn.Tasks.Utils
{
    internal class DevToolsEnvironment : IDisposable
    {
        private bool _isMessageFilterRegistered;

        public EnvDTE.Project GetProjectByName(string projectName)
        {
            var vsProcessId = GetVisualStudioProcessId();

            var dte = GetDevToolsEnvironment(vsProcessId);
            if (dte == null) return null;

            if (!_isMessageFilterRegistered)
            {
                MessageFilter.Register();
                _isMessageFilterRegistered = true;
            }

            var project = dte
                .Solution
                .Projects
                .OfType<EnvDTE.Project>()
                .SingleOrDefault(p =>
                    string.Compare(p.Name, projectName, StringComparison.OrdinalIgnoreCase) == 0);

            return project;
        }

        public void Dispose()
        {
            if (_isMessageFilterRegistered)
            {
                MessageFilter.Revoke();
            }
        }

        private int GetVisualStudioProcessId()
        {
            try
            {
                var process = Process.GetCurrentProcess();

                if (process.ProcessName.ToLower().Contains("devenv"))
                {
                    // We're being compiled directly by Visual Studio
                    return process.Id;
                }

                // We're being compiled by other tool (e.g. MSBuild) called from Visual Studio
                // therefore, Visual Studio is our parent process

                using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Process WHERE ProcessId = " + process.Id))
                {
                    foreach (var obj in searcher.Get())
                    {
                        var parentId = Convert.ToInt32((uint)obj["ParentProcessId"]);
                        var parentProcessName = Process.GetProcessById(parentId).ProcessName;

                        if (parentProcessName.ToLower().Contains("devenv"))
                        {
                            return parentId;
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Do nothing
            }

            return -1;
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
                        if (runningObjectMoniker != null)
                        {
                            runningObjectMoniker.GetDisplayName(bindCtx, null, out name);
                        }
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

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        // Implement the IOleMessageFilter interface
        [DllImport("Ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

        /// <summary>
        /// Contains the IOleMessageFilter thread error-handling functions
        /// See https://msdn.microsoft.com/en-us/library/ms228772
        /// </summary>
        private class MessageFilter : IOleMessageFilter
        {
            public static void Register()
            {
                // Register the IOleMessageFilter to handle any threading errors
                // See https://msdn.microsoft.com/en-us/library/ms228772

                IOleMessageFilter newFilter = new MessageFilter();

                IOleMessageFilter _;
                CoRegisterMessageFilter(newFilter, out _);
            }

            public static void Revoke()
            {
                IOleMessageFilter _;

                // Turn off the IOleMessageFilter
                CoRegisterMessageFilter(null, out _);
            }

            // IOleMessageFilter functions
            // Handle incoming thread requests
            int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
            {
                // Return the flag SERVERCALL_ISHANDLED
                return 0;
            }

            // Thread call was rejected, so try again
            int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
            {
                if (dwRejectType == 2) // SERVERCALL_RETRYLATER
                {
                    // Retry the thread call immediately if return >= 0 & < 100.
                    return 99;
                }

                // Too busy; cancel call
                return -1;
            }

            int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
            {
                // Return the flag PENDINGMSG_WAITDEFPROCESS
                return 2;
            }
        }

        [ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleMessageFilter
        {
            [PreserveSig]
            int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

            [PreserveSig]
            int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

            [PreserveSig]
            int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
        }
    }
}