//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ExcelDna.Integration.Rtd;
using ExcelDna.Logging;

namespace ExcelDna.Integration
{
    // TODO: Not sure how this should be used yet.... it was a good anchor for developing the rest.
    // CONSIDER: Do we want these to run 'as macro', or just run on the main thread?
    public class ExcelSynchronizationContext : SynchronizationContext
    {
        public override void Post(SendOrPostCallback d, object state)
        {
            SynchronizationManager.RunMacroSynchronization.RunAsMacroAsync(d, state);
        }

        public override void Send(SendOrPostCallback d, object state)
        {
            throw new NotImplementedException("ExcelSynchronizationContext does not currently allow synchronous calls.");
        }
    }

    // Just a this wrapper for the SynchronizationWindow - manages install / uninstall of the single instance, 
    // and access to the single instance.
    internal static class SynchronizationManager
    {
        static SynchronizationWindow _syncWindow;

        // Called from the Initialize (loading COM /RTD server) and/or from AutoOpen
        internal static void Install()
        {
            if (!ExcelDnaUtil.IsMainThread)
            {
                Logger.Initialization.Error("SynchronizationManager must be Installed from the main Excel thread.");
                return;
            }
            if (_syncWindow == null)
            {
                Logger.Initialization.Info("SynchronizationManager - Install");
                _syncWindow = new SynchronizationWindow();
            }
        }

        internal static void Uninstall()
        {
            if (_syncWindow != null)
            {
                Logger.Initialization.Info("SynchronizationManager - Uninstall");
                Debug.Assert(ExcelDnaUtil.IsMainThread, "SynchronizationManager must be Uninstalled from the main Excel thread.");
                _syncWindow.Dispose();
                _syncWindow = null;
            }
        }

        internal static bool IsInstalled
        {
            get { return (_syncWindow != null && _syncWindow.RunMacroSynchronization.IsRegistered); }
        }

        internal static RtdUpdateSynchronization RtdUpdateSynchronization
        {
            get
            {
                if (_syncWindow != null)
                    return _syncWindow.RtdUpdateSynchronization;

                return null;
            }
        }

        internal static RunMacroSynchronization RunMacroSynchronization
        {
            get
            {
                if (_syncWindow != null && _syncWindow.RunMacroSynchronization.IsRegistered) 
                    return _syncWindow.RunMacroSynchronization;

                return null;
            }
        }
    }

    internal class RtdUpdateSynchronization
    {
        SynchronizationWindow _syncWindow;

        // We'll use the Key collection of a Dictionary as a HashSet (which is not available on .NET 2.0).
        readonly Dictionary<IRTDUpdateEvent, object> _registeredRtdUpdates = new Dictionary<IRTDUpdateEvent, object>();
        Dictionary<IRTDUpdateEvent, object> _pendingRtdUpdates = new Dictionary<IRTDUpdateEvent, object>();
        readonly object _lockObject = new object();

        public RtdUpdateSynchronization(SynchronizationWindow syncWindow)
        {
            _syncWindow = syncWindow;
        }

        // Support for pushing UpdateNotify onto the main thread.
        // RTD server may be alive or now - here we don't worry.
        // We assume UpdateNotify is not called too often here (need RTD server to be careful).
        public void UpdateNotify(IRTDUpdateEvent updateEvent)
        {
            lock (_lockObject)
            {
                _pendingRtdUpdates[updateEvent] = null;
                _syncWindow.PostUpdateNotify();
            }
        }

        // Should only be called from the RTD server ServerStart
        // Keeps track of the 'alive' RTD servers.
        public void RegisterUpdateNotify(IRTDUpdateEvent updateEvent)
        {
            _registeredRtdUpdates[updateEvent] = null;
        }

        // Should be called from the RTD server ServerTerminate
        // This doesn't really solve the problem of another thread calling UpdateNotify after ServerTerminate....?
        public void DeregisterUpdateNotify(IRTDUpdateEvent updateEvent)
        {
            _registeredRtdUpdates.Remove(updateEvent);
        }

        // Runs on the main thread.
        // TODO: Do we need to clear these...?
        // TODO: Check again whether the UpdateNotify might fail. See: https://social.msdn.microsoft.com/Forums/office/en-US/0e3dfd86-62e1-4557-80a3-2b864c31cc52/excel-rtd-updatenotify-throws-exception-the-application-is-busy?forum=exceldev
        //       However, the problem reported there was surely when called from another thread. Does suspension of the object model affect us here?
        public void ProcessUpdateNotifications()
        {
            // CONSIDER: Do temp swap trick to reduce locking?
            lock (_lockObject)
            {
                // Only update servers that are still registered.
                foreach (IRTDUpdateEvent pendingRtdUpdate in _pendingRtdUpdates.Keys)
                {
                    if (_registeredRtdUpdates.ContainsKey(pendingRtdUpdate))
                    {
                        pendingRtdUpdate.UpdateNotify();
                    }
                }
            }
        }
    }

    internal class RunMacroSynchronization : IDisposable
    {
        #region Timer related stuff to create Application.Run retry timer
        // Not clear whether timer should be here on in the SynchronizationWindow.
        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        static extern IntPtr SetTimer(HandleRef hWnd, int nIDEvent, int uElapse, IntPtr lpTimerFunc);
        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        static extern bool KillTimer(HandleRef hwnd, int idEvent);
        int _timerId;   // Only gets values 0 or 1. Used both as the uIDEvent parameter for the Windows Timer, and as a flag to indicate whether there is an active timer.
        const int RETRY_INTERVAL_MS = 250;
        #endregion

        SynchronizationWindow _syncWindow;
        readonly object _sendOrPostLock = new object();

        bool _isRunningMacros = false;
        bool _syncPosted = false;
        string _syncMacroName;
        object _syncMacroRegistrationId = null;

        readonly Queue<SendOrPostCallback> _postCallbacks = new Queue<SendOrPostCallback>();
        readonly Queue<Object> _postStates = new Queue<Object>();

        public RunMacroSynchronization(SynchronizationWindow syncWindow)
        {
            _syncWindow = syncWindow;
            Register();
        }

        // Called from outside on any thread to enqueue work.
        public void RunAsMacroAsync(SendOrPostCallback d, object state)
        {
            if (!SynchronizationManager.IsInstalled)
            {
                throw new InvalidOperationException("SynchronizationManager is not registered. This is an unexpected error.");
            }

            lock (_sendOrPostLock)
            {
                _postCallbacks.Enqueue(d);
                _postStates.Enqueue(state);

                // CAREFUL: This check needs to be in the same lock used in SyncMacro when running
                if (!_isRunningMacros && !_syncPosted)
                {
                    _syncWindow.PostRunSyncMacro();
                    _syncPosted = true;
                }
            }
        }

        // Called from the SyncWindow on the main thread, when the posted message is received.
        // Tries to get into a macro context by calling Application.Run(_syncWindow) and if that fails, 
        // set a timer to retry again soon and back off for now. 
        // The timer is hooked on to the _syncWindow, and its WM_TIMER message received in the _syncWindow calls this function again.
        // TODO: Consider this discussion (maybe allow user to register callback to check other conditions before running): 
        //       https://exceldna.codeplex.com/discussions/565495
        public void ProcessRunSyncMacroMessage()
        {
            try
            {
                bool runOK = COMRunMacro(_syncMacroName);
                if (!runOK && _timerId == 0)
                {
                    // Timer is not yet set - so set a new timer
                    // The Timer will be stopped when SyncMacro actually runs.
                    _timerId = 1; // We're always on the main thread, so no race condition between checking and setting the id/flag
                    IntPtr result = SetTimer(new HandleRef(_syncWindow, _syncWindow.Handle), _timerId, RETRY_INTERVAL_MS, IntPtr.Zero);
                    if (result == IntPtr.Zero)
                    {
                        // TODO: Handle unexpected error in setting timer
                        Debug.Print("SynchronizationWindow timer could not be set.");
                    }
                }
            }
            catch (InvalidOperationException)
            {
                // Expected when Excel is shutting down - abandon
                Debug.Print("Error trying to run SyncMacro - Excel is shutting down. Async macro queue abandoned.");
            }
            catch (Exception ex)
            {
                // TODO: Handle unexpected error
                Debug.Print("Unexpected error trying to run SyncMacro: " + ex);
            }
        }

        // This is the helper macro that runs (on the main thread)
        void SyncMacro(double _unused_)
        {
            // Check for timer and disable
            if (_timerId != 0)
            {
                KillTimer(new HandleRef(_syncWindow, _syncWindow.Handle), _timerId);
                _timerId = 0; // No race condition between the check and the reset here, since we're always on the main thread.
            }
            // Run everything currently in the queue
            _isRunningMacros = true;
            while (_isRunningMacros)
            {
                SendOrPostCallback work = null;
                object state = null;
                lock (_sendOrPostLock)
                {
                    _isRunningMacros = _postCallbacks.Count > 0;
                    if (_isRunningMacros)
                    {
                        work = _postCallbacks.Dequeue();
                        state = _postStates.Dequeue();
                    }
                    else
                    {
                        // set flag that we need to post again for work
                        // and exit
                        _syncPosted = false;
                        return;
                    }
                }
                try
                {
                    work(state);
                }
                catch (Exception ex)
                {
                    // CONSIDER: Integrate with Logging here...
                    Debug.Print("Async delegate exception: " + ex);
                }
            }
        }

        // Register the helper macro with Excel, so that Application.Run can call it.
        void Register()
        {
            // CONSIDER: Can this be cleaned up by calling ExcelDna.Loader?
            // We must not be in a function when this is run, nor in an RTD method call.
            _syncMacroName = "SyncMacro_" + ExcelDnaUtil.XllGuid.ToString("N");
            ExcelIntegration.SetSyncMacro(SyncMacro);

            object[] registerParameters = new object[6];
            registerParameters[0] = DnaLibrary.XllPath;
            registerParameters[1] = "SyncMacro";
            registerParameters[2] = ">B"; // Takes double, abuse in-place flag '>' to return void
            registerParameters[3] = _syncMacroName;
            registerParameters[4] = "value";
            registerParameters[5] = 2; // macro

            object xlCallResult;
            XlCall.TryExcel(XlCall.xlfRegister, out xlCallResult, registerParameters);
            Logger.Registration.Verbose("Register SyncMacro - XllPath={0}, ProcName={1}, FunctionType={2}, MethodName={3} - Result={4}", 
                registerParameters[0], registerParameters[1], registerParameters[2], registerParameters[3], xlCallResult);
            if (xlCallResult is double)
            {
                _syncMacroRegistrationId = (double)xlCallResult;
            }
            else
            {
                throw new InvalidOperationException("Synchronization macro registration failed.");
            }
        }

        public bool IsRegistered
        {
            get { return _syncMacroRegistrationId != null; }
        }

        void Unregister()
        {
            // Clear the name and unregister
            XlCall.Excel(XlCall.xlfSetName, _syncMacroName);
            XlCall.Excel(XlCall.xlfUnregister, _syncMacroRegistrationId);
            _syncMacroRegistrationId = null;
        }

        public void Dispose()
        {
            Unregister();
        }

        static readonly CultureInfo _enUsCulture = new CultureInfo(1033); // Don't know if this is useful...

        // Invoke Application.Run to run a macro
        static bool COMRunMacro(string macroName)
        {
            try
            {
                // If busy editing, don't even try to call Application.Run.
                if (IsInFormulaEditMode()) return false;

                object xlApp = ExcelDnaUtil.Application;
                Type appType = xlApp.GetType();

                // Now try Application.Run(macroName) if we are still alive.
                appType.InvokeMember("Run", BindingFlags.InvokeMethod, null, xlApp, new object[] { macroName, 0.0 }, _enUsCulture);
                return true;
            }
            catch (TargetInvocationException tie)
            {
                // Deal with the COM exception that we get if the Application object does not allow us to call 'Run' right now.
                COMException cex = tie.InnerException as COMException;
                if (cex != null && IsRetry(cex))
                    return false;

                // Unexpected error
                throw;
            }
        }

        #region Checks for known COM errors
        const uint RPC_E_SERVERCALL_RETRYLATER = 0x8001010A;
        const uint RPC_E_CALL_REJECTED = 0x80010001; // Not sure when we get this one?
                                                     // Maybe when trying to get the Application object from another thread, 
                                                     // triggered by a ribbon handler, while Excel is editing a cell.
        const uint VBA_E_IGNORE = 0x800AC472;        // Excel has suspended the object browser
        const uint NAME_NOT_FOUND = 0x800A03EC;      // When called from the main thread, but Excel is busy.

        static bool IsRetry(COMException e)
        {
            uint errorCode = (uint)e.ErrorCode;
            switch (errorCode)
            {
                case RPC_E_SERVERCALL_RETRYLATER:
                case VBA_E_IGNORE:
                case NAME_NOT_FOUND:
                case RPC_E_CALL_REJECTED:
                    return true;
                default:
                    return false;
            }
        }
        #endregion

        #region IsInFormulaEditMode helpers
        // It's hard to know if Excel is in Edit mode without causing side effects.
        // Using a COM call to check (Application.Interaction = ... / Application.ReferenceStyle = ...)
        //   gives the right answer, but the resulting exception still causes the editing 
        // abberations we see when calling Application.Run.
        // Checking the focused window class doesn't always work (focus might be on another application)
        // One proper solution might be a message hook to detect focus changes, but I'd not want to do that from 
        //   managed code in every add-in, to hook every message on the main app.
        // The UI Automation stuff would work, but might be hard to do from .NET 2.0.
        // So as a first I attempt I try to check which window has the keyboard focus, 
        //   and if it is not an Excel window (GetFocus returns 0), 
        //   I fall back to the standard menu enabled check
        //   (which still works under Excel 2007+, presumably for backward compatibility)

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr GetFocus();
        
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern IntPtr GetProcAddress([In] IntPtr hModule, [In, MarshalAs(UnmanagedType.LPStr)] string lpProcName);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)] static extern bool IsWindow(IntPtr hWnd);

        static string GetWindowClassName(IntPtr hWnd)
        {
            StringBuilder buffer = new StringBuilder(256);
            GetClassName(hWnd, buffer, buffer.Capacity);
            return buffer.ToString();
        }

        static bool IsFileOpenMenuEnabled()
        {
            CustomUI.CommandBars commandBars = CustomUI.ExcelCommandBarUtil.GetCommandBars();
            CustomUI.CommandBar worksheetMenu = commandBars[1]; // Worksheet Menu Bar
            CustomUI.CommandBarControl openMenuButton = worksheetMenu.FindControl(Missing.Value, /* ID:= */ 23, Missing.Value, Missing.Value, /* Recursive:= */ true);
            return openMenuButton.Enabled;
        }

        // The call to LPenHelper will cause an AccessViolation after Excel starts shutting down.
        // .NET40: If this library is recompiled to target .NET 4+, we need to add an attribute to indicate that this exception 
        // (which might indicate corrupted state) should be handled in our code.
        // For now, we target .NET 2.0, and even when running under .NET 4.0 we'll see the exception and be able to handle is.
        // See: http://msdn.microsoft.com/en-us/magazine/dd419661.aspx
        // (Also for CheckExcelApiAvailable())

        // [HandleProcessCorruptedStateExceptions]
        static int CallPenHelper(int wCode, ref XlCall.FmlaInfo fmlaInfo)
        {
            try
            {
                // (If Excel is shutting down, we see an Access Violation here, reading at 0x00000018.)
                return XlCall.LPenHelper(XlCall.xlGetFmlaInfo, ref fmlaInfo);
            }
            catch (AccessViolationException ave)
            {
                throw new InvalidOperationException("LPenHelper call failed. Excel is shutting down.", ave);
            }
        }

        static bool IsInFormulaEditMode()
        {
            // I assume LPenHelper is available under Excel 2007+
            if (ExcelDnaUtil.ExcelVersion >= 12.0)
            {
                // check edit state directly
                var fmlaInfo = new XlCall.FmlaInfo();

                // If Excel is shutting down, CallPenHelper will throw an InvalidOperationException.
                var result = CallPenHelper(XlCall.xlGetFmlaInfo, ref fmlaInfo);
                if (result == 0)
                {
                    // Succeeded
                    return fmlaInfo.wPointMode != XlCall.xlModeReady;
                }
            }

            // Otherwise try Focus windows check, else menu check.
            IntPtr focusedWindow = GetFocus();
            if (focusedWindow == IntPtr.Zero)
            {
                // Excel (this thread) does not have the keyboard focus. Use the Menu check instead.
                bool menuEnabled = IsFileOpenMenuEnabled();
                // Debug.Print("Menus Enabled: " + menuEnabled);
                return !menuEnabled;
            }

            string className = GetWindowClassName(focusedWindow);
            // Debug.Print("Focused window class: " + className);

            return className == "EXCEL<" || className == "EXCEL6";
        }
        #endregion
    }

    // SynchronizationWindow installs a window on the main Excel message loop, 
    // to allow us to jump onto the main thread for calling RTD update notifications and running macros.
    sealed class SynchronizationWindow : NativeWindow, IDisposable
    {
        [DllImport("user32.dll")]
        static extern bool PostMessage(HandleRef hwnd, int msg, IntPtr wparam, IntPtr lparam);

        // Helpers for the two sync aspects
        public RtdUpdateSynchronization RtdUpdateSynchronization;
        public RunMacroSynchronization RunMacroSynchronization;

        readonly IntPtr HWND_MESSAGE = (IntPtr)(-3);
        const int WM_TIMER = 0x0113;
        const int WM_USER = 0x400;
        const int WM_UPDATENOTIFY = WM_USER + 1;
        const int WM_SYNCMACRO = WM_USER + 2;
        //const int WM_SYNCMACRO_DIRECT = WM_USER + 3;

        public SynchronizationWindow()
        {
            CreateParams cp = new CreateParams();
            if (Environment.OSVersion.Version.Major >= 5)
                cp.Parent = HWND_MESSAGE;

            CreateHandle(cp);
            RtdUpdateSynchronization = new RtdUpdateSynchronization(this);
            RunMacroSynchronization = new RunMacroSynchronization(this);
        }

        internal void PostUpdateNotify()
        {
 	         PostMessage(new HandleRef(this, Handle), WM_UPDATENOTIFY, IntPtr.Zero, IntPtr.Zero);
        }

        internal void PostRunSyncMacro()
        {
            PostMessage(new HandleRef(this, Handle), WM_SYNCMACRO, IntPtr.Zero, IntPtr.Zero);
        }

        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case WM_UPDATENOTIFY:
                    RtdUpdateSynchronization.ProcessUpdateNotifications();
                    break;
                case WM_SYNCMACRO:
                case WM_TIMER:
                    RunMacroSynchronization.ProcessRunSyncMacroMessage();
                    break;
                default:
                    base.WndProc(ref m);
                    break;
            }
        }

        public void Dispose()
        {
            RunMacroSynchronization.Dispose();
            DestroyHandle();
        }
    }
}
