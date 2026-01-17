//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using ExcelDna.Integration.Rtd;
using ExcelDna.Logging;

#if USE_WINDOWS_FORMS
using System.Windows.Forms;
#else
using ExcelDna.Integration.Win32;
#endif

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
        internal static void Install(bool isAutoOpen)
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
            if (isAutoOpen)
            {
                // Safe to call more than once
                _syncWindow.Register();
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
            Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.f}] RtdUpdateSynchronization.UpdateNotify");
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
            Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.f}] ProcessUpdateNotifications");
            // CONSIDER: Do temp swap trick to reduce locking?
            lock (_lockObject)
            {
                try
                {
                    // Only update servers that are still registered.
                    foreach (IRTDUpdateEvent pendingRtdUpdate in _pendingRtdUpdates.Keys)
                    {
                        if (_registeredRtdUpdates.ContainsKey(pendingRtdUpdate))
                        {
                            Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.f}] Calling UpdateNotify - Thread: {System.Threading.Thread.CurrentThread.ManagedThreadId}");
                            pendingRtdUpdate.UpdateNotify();
                        }
                    }
                    // NOTE: Added the Clear() call on 2017/09/07 (after v0.34)
                    _pendingRtdUpdates.Clear();
                }
                catch (InvalidCastException ice)
                {
                    // UpdateNotify has been reported to fail after the Thomson Reuters Client is installed, 
                    // with an InvalidCastException:
                    //     at System.StubHelpers.StubHelpers.GetCOMIPFromRCW(System.Object, IntPtr, IntPtr ByRef, Boolean ByRef) 
                    //     at ExcelDna.Integration.Rtd.IRTDUpdateEvent.UpdateNotify()
                    //     ....
                    Logger.RtdServer.Error(ice, "There was an error when calling the Excel RTD UpdateNotify callback.\r\nThis has been reported after installing a conflicting real-time add-in, and might require a repair or re-install of Excel.\r\nReal-time updates and/or async functions could not be processed.");
                }
                catch (Exception ex)
                {
                    // An unhandled exception here will terminate the process.
                    Logger.RtdServer.Error(ex, "There was an error when calling the Excel RTD UpdateNotify callback.\r\nReal-time updates and/or async functions could not be processed.");
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

        // Here we catch an InvalidOperationException that we throw from CallPenHelper below
        // This seems to need the special attribute too !?
        // TODO: NET6+: See notes at CallPenHelper
#if NETFRAMEWORK
        [HandleProcessCorruptedStateExceptions]
#endif
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
            catch (InvalidOperationException /*ioe*/)
            {
                // Expected when Excel is shutting down - abandon
                Logger.Runtime.Warn("Error (InvalidOperationException) trying to run SyncMacro - Excel is shutting down. Queued macro execution abandoned.");
            }
            catch (Exception ex)
            {
                // Includes unexpected TargetInvocationExceptions which are not known COMExceptions
                Logger.Runtime.Error(ex, "Unexpected error trying to run SyncMacro for queued macro execution.");
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
                    UnhandledExceptionHandler handler = ExcelIntegration.GetRegisterUnhandledExceptionHandler();
                    if (handler != null)
                    {
                        try
                        {
                            handler(ex);
                        }
                        catch (Exception uehex)
                        {
                            Logger.Runtime.Error(ex, "Unhandled exception in async delegate call.");
                            Logger.Runtime.Error(uehex, "Unhandled exception in UnhandledExceptionHandler after async delegate call.");
                        }
                    }
                    else
                    {
                        Logger.Runtime.Error(ex, "Unhandled exception in async delegate call.");
                    }
                }
            }
        }

        // Register the helper macro with Excel, so that Application.Run can call it.
        public void Register()
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
        // Returns true if it ran OK, false if not and we should retry
        static bool COMRunMacro(string macroName)
        {
            try
            {
                // If busy editing, don't even try to call Application.Run.
                if (IsInFormulaEditMode())
                    return false;

                object xlApp = ExcelDnaUtil.GetApplicationNotProtectedNoThrow();
                if (xlApp == null)
                {
                    // Some possibilities that get us here:
                    // * Can't get Application object at all - first time we're trying is here, no workbook open and hence C API is needed but not available
                    // * Excel is shutting down (would have abandoned in the past, now we keep re-trying)

                    return false;
                }

                // Now try Application.Run(macroName) if we are still alive.
                object result = ComInterop.Util.TypeAdapter.Invoke("Run", new object[] { macroName, 0.0 }, xlApp);
                // Sometimes (e.g. when the paste live preview feature is active) Application.Run returns the integer value for E_FAIL, 
                // and not a COM Error that is converted to an exception.
                if (!result.Equals(0.0))    // We expect our "void" macro to just return 0.0.
                {
#if DEBUG
                    // Some extra checks to see if we can understand the return values better
                    if (!(result is int))
                    {
                        Logger.Registration.Error("Unexpected return type from Application.Run(\"SyncMacro_...\") - " + result);
                    }
                    else
                    {
                        if ((int)result != E_FAIL
                            && (int)result != E_NA)
                        {
                            Logger.Registration.Error("Unexpected return value from Application.Run(\"SyncMacro_...\") - " + result);
                        }
                    }
#endif
                    return false;
                }
                return true;
            }
            catch (TargetInvocationException tie)
            {
                // Deal with the COM exception that we get if the Application object does not allow us to call 'Run' right now.
                COMException cex = tie.InnerException as COMException;
                if (cex != null && IsRetry(cex))
                    return false;

                // Unexpected error - very bad - we abandon the whole QueueAsMacro plan forever - the exception is handled higher up and logged
                throw;
            }
        }

        const int E_FAIL = unchecked((int)0x80004005);
        const int E_NA = unchecked((int)0x800A07FA);   // Not sure why we get this back from Application.Run("SyncMacro...")

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

        public static bool IsInFormulaEditMode()
        {
            // check edit state directly
            var fmlaInfo = new XlCall.FmlaInfo();

            // If Excel is shutting down, PenHelper will throw an InvalidOperationException.
            var result = XlCall.PenHelper(XlCall.xlGetFmlaInfo, ref fmlaInfo);
            if (result == 0)
            {
                // Succeeded
                return fmlaInfo.wPointMode != XlCall.xlModeReady;
            }
            else
            {
                // Log and return true (the safer option) ???
                // Decreased to Warn to prevent LogDisplay pop-up loop
                Logger.Registration.Warn("IsInFormulaEditMode - PenHelper failed, result " + result);
                return true;
            }
        }
        #endregion
    }

    // SynchronizationWindow installs a window on the main Excel message loop, 
    // to allow us to jump onto the main thread for calling RTD update notifications and running macros.
    sealed class SynchronizationWindow : NativeWindow, IDisposable
    {
        [DllImport("user32.dll", SetLastError = true)]
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

        // We keep track of messages that are pending (after failed PostMessage calls)
        Queue<int> _pendingPostMessages = new Queue<int>();
        readonly object _pendingPostMessageLock = new object();
        bool _pendingPostMessageThreadRunning = false;
        const int RetryPostMessageDelayMs = 250;

        public SynchronizationWindow()
        {
            CreateParams cp = new CreateParams();
            if (Environment.OSVersion.Version.Major >= 5)
                cp.Parent = HWND_MESSAGE;

            CreateHandle(cp);
            RtdUpdateSynchronization = new RtdUpdateSynchronization(this);
            RunMacroSynchronization = new RunMacroSynchronization(this);
        }

        // Complete the initialization when we are called in AutoOpen, and can try to install the macro
        public void Register()
        {
            // This might fail under some expected circumstances - e.g. when called from Cluster host of RegSvr32 - exception thrown in thoses cases
            if (!RunMacroSynchronization.IsRegistered)
            {
                RunMacroSynchronization.Register();
            }
        }

        internal void PostUpdateNotify()
        {
            ReliablePostMessage(WM_UPDATENOTIFY);
        }

        internal void PostRunSyncMacro()
        {
            ReliablePostMessage(WM_SYNCMACRO);
        }

        // ReliablePostMessage ensures that there will be a successful PostMessage call with this msg
        internal void ReliablePostMessage(int msg)
        {
            if (!_pendingPostMessageThreadRunning)
            {
                // Normal case - try to post directly
                var postOK = PostMessage(new HandleRef(this, Handle), msg, IntPtr.Zero, IntPtr.Zero);
                if (postOK)
                    return;

                // There was an error - log it and continue to retry behaviour
                // Instead of dealing with the error code, we could do:
                //     throw new Win32Exception() 
                // which calls GetLastWin32Error and creates a formatted message string
                // See: https://blogs.msdn.microsoft.com/adam_nathan/2003/04/25/getlasterror-and-managed-code/
                int err = Marshal.GetLastWin32Error();
                Logger.Runtime.Warn("SynchronizationWindow - PostMessage Error {0}", err);
            }

            // Rarely we might have a PostMessage call that fails
            // (when Excel is very busy, the main thread is not processing messages and the queue gets too big?)
            // Then we switch to a reliable retry mode

            // Either PostMessage failed or we didn't even try, since we are (or were) busy running queued messages
            lock (_pendingPostMessageLock)
            {
                Logger.Runtime.Verbose("SynchronizationWindow - Enqueueing message {0}", msg);
                _pendingPostMessages.Enqueue(msg);
            }
            // Retry or ensure retry thread is running
            AttemptPostMessages();
        }

        // Can be called from any thread to try to process the messages in _pendingPostMessages
        internal void AttemptPostMessages()
        {
            lock (_pendingPostMessageLock)
            {
                Logger.Runtime.Verbose("SynchronizationWindow - AttemptPostMessages for {0} message(s)", _pendingPostMessages.Count);
                bool postOK = true;
                while (_pendingPostMessages.Count > 0 && postOK)
                {
                    int msg = _pendingPostMessages.Peek();
                    postOK = PostMessage(new HandleRef(this, Handle), msg, IntPtr.Zero, IntPtr.Zero);
                    if (postOK)
                    {
                        _pendingPostMessages.Dequeue();
                    }
                    else
                    {
                        // Error - we'll stop the loop since postOK == false
                        int err = Marshal.GetLastWin32Error();
                        Logger.Runtime.Warn("SynchronizationWindow - PostMessage Error {0}", err);
                    }
                }
                // Were we finished?
                if (_pendingPostMessages.Count == 0)
                {
                    // No more pending messages - retry thread can exit if it is running
                    _pendingPostMessageThreadRunning = false;
                    Logger.Runtime.Verbose("SynchronizationWindow - AttemptPostMessages complete - all messages posted");
                }
                else if (!_pendingPostMessageThreadRunning)
                {
                    // We still have pending messages and no thread was running
                    // so we start a new retry thread
                    Logger.Runtime.Verbose("SynchronizationWindow - AttemptPostMessages complete - starting retry thread");
                    _pendingPostMessageThreadRunning = true;
                    new Thread(RetryPostMessages).Start();
                }
                else
                {
                    Logger.Runtime.Verbose("SynchronizationWindow - AttemptPostMessages ending - {0} message(s) remain", _pendingPostMessages.Count);
                }
            }
        }

        // This is the retry thread routine
        void RetryPostMessages(object _unused_)
        {
            Logger.Runtime.Verbose("SynchronizationWindow - RetryPostMessages starting");
            while (true)
            {
                Thread.Sleep(RetryPostMessageDelayMs);
                lock (_pendingPostMessageLock)
                {
                    AttemptPostMessages();
                    if (!_pendingPostMessageThreadRunning)
                    {
                        // Stop running
                        Logger.Runtime.Verbose("SynchronizationWindow - RetryPostMessages stopping");
                        return;
                    }
                }
            }
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
