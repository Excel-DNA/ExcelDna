/*
  Copyright (C) 2005-2012 Govert van Drimmelen

  This software is provided 'as-is', without any express or implied
  warranty.  In no event will the authors be held liable for any damages
  arising from the use of this software.

  Permission is granted to anyone to use this software for any purpose,
  including commercial applications, and to alter it and redistribute it
  freely, subject to the following restrictions:

  1. The origin of this software must not be misrepresented; you must not
     claim that you wrote the original software. If you use this software
     in a product, an acknowledgment in the product documentation would be
     appreciated but is not required.
  2. Altered source versions must be plainly marked as such, and must not be
     misrepresented as being the original software.
  3. This notice may not be removed or altered from any source distribution.


  Govert van Drimmelen
  govert@icon.co.za
*/

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using ExcelDna.Integration.Rtd;

namespace ExcelDna.Integration
{
    public static class SynchronizationManager
    {
        static SynchronizationWindow _syncWindow;
        static int _registerCount = 0;

        // Don't want to call it 'Install' since that is used for installing the SynchronizationContext as the 'Current' SyncContext in the thread.
        // TODO: Check that this happens on the main Excel thread, and not in a 'function' context.
        // TODO: Reference count for Register, matched by unregister in ExcelRtdServer...?
        public static void Register()
        {
            _registerCount++;
            if (_syncWindow == null)
            {
                _syncWindow = new SynchronizationWindow();
                _syncWindow.Register();
            }
        }

        // TODO: Check that this happens on the main Excel thread.
        public static void Unregister()
        {
            _registerCount--;
            if (_registerCount == 0 && _syncWindow != null)
            {
                _syncWindow.Unregister();
                _syncWindow = null;
            }
        }

        public static bool IsRegistered
        {
            get { return (_syncWindow != null); }
        }

        internal static SynchronizationWindow SynchronizationWindow
        {
            get { return _syncWindow; }
        }
    }

    // TODO: Not sure how this should be used yet....
    // TODO: How to deal with 'installing' etc.
    // CONSIDER: Do we want these to run 'as macro', or just run on the main thread?
    public class ExcelSynchronizationContext : SynchronizationContext
    {
        public override void Post(SendOrPostCallback d, object state)
        {
            if (!SynchronizationManager.IsRegistered)
                throw new InvalidOperationException("SynchronizationManager is not registered.");

            SynchronizationManager.SynchronizationWindow.RunAsMacroAsync(d, state);
        }
    }

    // SynchronizationWindow supports running code on the main Excel thread.
    // TODO: Make this into a static class?
    internal sealed class SynchronizationWindow : NativeWindow, IDisposable
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        static extern IntPtr SetTimer(HandleRef hWnd, int nIDEvent, int uElapse, IntPtr lpTimerFunc);
        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        static extern bool KillTimer(HandleRef hwnd, int idEvent);
        [DllImport("user32.dll")]
        static extern bool PostMessage(HandleRef hwnd, int msg, IntPtr wparam, IntPtr lparam);

        readonly object _lockObject = new object();
        int _timerId;
        // We'll use the Key collection of a Dictionary as a HashSet (which is not available on .NET 2.0).
        readonly Dictionary<IRTDUpdateEvent, object> _pendingRtdUpdates = new Dictionary<IRTDUpdateEvent, object>();

        readonly IntPtr HWND_MESSAGE = (IntPtr)(-3);
        const int WM_TIMER = 0x0113;
        const int WM_USER = 0x400;
        const int WM_UPDATENOTIFY = WM_USER + 1;
        const int WM_SYNCMACRO = WM_USER + 2;
        //const int WM_SYNCMACRO_DIRECT = WM_USER + 3;
        const int RETRY_INTERVAL_MS = 250;

        public SynchronizationWindow()
        {
            CreateParams cp = new CreateParams();
            if (Environment.OSVersion.Version.Major >= 5)
                cp.Parent = HWND_MESSAGE;

            CreateHandle(cp);
        }

        #region RTD UpdateNotify support
        // Support for pushing UpdateNotify onto the main thread.
        // TODO: How do we know that the RTD server is still alive when it runs?
        public void UpdateNotify(IRTDUpdateEvent updateEvent)
        {
            lock (_lockObject)
            {
                _pendingRtdUpdates[updateEvent] = null;
                PostMessage(new HandleRef(this, Handle), WM_UPDATENOTIFY, IntPtr.Zero, IntPtr.Zero);
            }
        }

        // Should be called from the RTD server ServerTerminate
        // CONSIDER: Make an updater object that can be disposed automatically?
        // This doesn't really solve the problem of another thread calling UpdateNotify after ServerTerminate....?
        public void CancelUpdateNotify(IRTDUpdateEvent updateEvent)
        {
            lock (_lockObject)
            {
                _pendingRtdUpdates.Remove(updateEvent);
            }
        }
        #endregion

        public void RunAsMacroAsync(SendOrPostCallback d, object state)
        {
            lock (_sendOrPostLock)
            {
                _postCallbacks.Enqueue(d);
                _postStates.Enqueue(state);

                // CAREFUL: This check needs to be in the same lock used in SyncMacro when running
                if (!_isRunningMacros)
                {
#if DEBUG
                    Debug.Print("About to enqueue a macro: " + state);
//                    d(state);
#endif
                    PostMessage(new HandleRef(this, Handle), WM_SYNCMACRO, IntPtr.Zero, IntPtr.Zero);
                }
            }
        }

        void ProcessRunSyncMacro()
        {
            try
            {
                bool runOK = COMRunMacro();
                if (!runOK && _timerId == 0)
                {
                    // Timer not yet set - start it
                    // Timer will be stopped when SyncMacro actually runs.
                    IntPtr result = SetTimer(new HandleRef(this, Handle), _timerId++, RETRY_INTERVAL_MS, IntPtr.Zero);
                    if (result == IntPtr.Zero)
                    {
                        // TODO: Handle unexpected error in setting timer
                        Debug.Print("SynchronizationWindow timer could not be set.");
                    }
                }
            }
            catch (Exception ex)
            {
                // TODO: Handle unexpected error
                Debug.Print("Unexpected error trying to run SyncMacro: " + ex);
            }
        }

        bool COMRunMacro()
        {
            try
            {
                object xlApp = ExcelDnaUtil.Application;
                xlApp.GetType().InvokeMember("Run", BindingFlags.InvokeMethod, null, xlApp, new object[] { _syncMacroName, 0.0 });
                return true;
            }
            catch (TargetInvocationException tie)
            {
                COMException cex = tie.InnerException as COMException;
                if (cex != null && IsRetry(cex))
                    return false;

                // Unexpected error
                throw;
            }
            // Not releasing the Application object here, since we are on the main thread, and might get the ribbon-cached Application.
        }

        const uint RPC_E_SERVERCALL_RETRYLATER = 0x8001010A;
        const uint RPC_E_CALL_REJECTED = 0x80010001; // Not sure when we get this one?
                                                     // Maybe when trying to get the Application object from another thread, 
                                                     // triggered by a ribbon handler, while Excel is editing a cell.
        const uint VBA_E_IGNORE = 0x800AC472;        // Excel has suspended the object browser
        const uint UNKNOWN_E_UNKNOWN = 0x800A03EC;   // When called from the main thread, but Excel is busy.

        static bool IsRetry(COMException e)
        {
            uint errorCode = (uint)e.ErrorCode;
            switch (errorCode)
            {
                case RPC_E_SERVERCALL_RETRYLATER:
                case VBA_E_IGNORE:
                case UNKNOWN_E_UNKNOWN:
                case RPC_E_CALL_REJECTED:
                    return true;
                default:
                    return false;
            }
        }

        void ProcessUpdateNotifications()
        {
            lock (_lockObject)
            {
                Debug.Print("Calling UpdateNotify");
                foreach (IRTDUpdateEvent pendingRtdUpdate in _pendingRtdUpdates.Keys)
                {
                    pendingRtdUpdate.UpdateNotify();
                }
                _pendingRtdUpdates.Clear();
            }
        }

        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case WM_UPDATENOTIFY:
                    ProcessUpdateNotifications();
                    break;
                case WM_SYNCMACRO:
                case WM_TIMER:
                    ProcessRunSyncMacro();
                    break;
                // TODO: case WM_CLOSE / WM_DESTROY: ????
                default:
                    base.WndProc(ref m);
                    break;
            }
        }

        bool _isRunningMacros = false;
        string _syncMacroName;
        object _syncMacroRegistrationId = null;

        readonly Queue<SendOrPostCallback> _postCallbacks = new Queue<SendOrPostCallback>();
        readonly Queue<Object> _postStates = new Queue<Object>();
        readonly object _sendOrPostLock = new object();

        void SyncMacro(double _unused_)
        {
            // Check for timer and disable
            if (_timerId != 0)
            {
                KillTimer(new HandleRef(this, Handle), _timerId);
                _timerId = 0;
            }
            // Run everything currently in the queue
            _isRunningMacros = true;
            while (_isRunningMacros)
            {
                SendOrPostCallback work = null;
                object state = null;
                // PostMessage
                lock (_sendOrPostLock)
                {
                    _isRunningMacros = _postCallbacks.Count > 0;
                    if (_isRunningMacros)
                    {
                        work = _postCallbacks.Dequeue();
                        state = _postStates.Dequeue();
                    }
                }
                if (_isRunningMacros)
                {
                    try
                    {
                        work(state);
                    }
                    catch (Exception ex)
                    {
                        Debug.Print("Async delegate exception: " + ex);
                    }
                }
            }
        }

        public void Register()
        {
            // CONSIDER: Can this be cleaned up by calling ExcelDna.Loader?
            // We must not be in a function when this is run, nor in an RTD method call.
            _syncMacroName = "SyncMacro_" + Guid.NewGuid().ToString("N");
            Integration.SetSyncMacro(SyncMacro);

            object[] registerParameters = new object[6];
            registerParameters[0] = DnaLibrary.XllPath;
            registerParameters[1] = "SyncMacro";
            registerParameters[2] = ">B"; // Takes double, abuse in-place flag '>' to return void
            registerParameters[3] = _syncMacroName;
            registerParameters[4] = "value";
            registerParameters[5] = 2; // macro

            object xlCallResult;
            XlCall.TryExcel(XlCall.xlfRegister, out xlCallResult, registerParameters);
            Debug.Print("Register SyncMacro - XllPath={0}, ProcName={1}, FunctionType={2}, MethodName={3} - Result={4}", 
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

        public void Unregister()
        {
            XlCall.Excel(XlCall.xlfUnregister, _syncMacroRegistrationId);
            _syncMacroRegistrationId = null;
        }

        public void Dispose()
        {
            // CONSIDER: Must this also Unregister?
            DestroyHandle();
        }
    }
}
