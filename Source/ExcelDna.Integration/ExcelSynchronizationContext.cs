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
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using ExcelDna.Integration.Rtd;

namespace ExcelDna.Integration
{

    // TODO: Not sure how this should be used yet.... it was a good anchor for developing the rest.
    // CONSIDER: Do we want these to run 'as macro', or just run on the main thread?
    public class ExcelSynchronizationContext : SynchronizationContext
    {
        public override void Post(SendOrPostCallback d, object state)
        {
            if (!SynchronizationManager.IsInstalled)
            {
                throw new InvalidOperationException("SynchronizationManager is not registered. Call ExcelAsyncUtil.Initialize() before use.");
            }

            SynchronizationManager.RunMacroSynchronization.RunAsMacroAsync(d, state);
        }
    }

    // Just a this wrapper for the SynchronizationWindow - manages install / uninstall of the single instance, 
    // and access to the single instance.
    internal static class SynchronizationManager
    {
        static SynchronizationWindow _syncWindow;
        static int _installCount = 0;

        // TODO: Check that this does not happen in a 'function' context.
        internal static void Install()
        {
            if (!ExcelDnaUtil.IsMainThread())
            {
                throw new InvalidOperationException("SynchronizationManager must be installed from the main Excel thread. Ensure that ExcelAsyncUtil.Initialize() is called from AutoOpen() or a macro on the main Excel thread.");
            }
            if (_syncWindow == null)
            {
                _syncWindow = new SynchronizationWindow();
            }
            _installCount++;
        }

        internal static void Uninstall()
        {
            if (!ExcelDnaUtil.IsMainThread())
            {
                throw new InvalidOperationException("SynchronizationManager must be uninstalled from the main Excel thread. Ensure that ExcelAsyncUtil.Uninitialize() is called from AutoOpen() or a macro on the main Excel thread.");
            }
            _installCount--;
            if (_installCount == 0 && _syncWindow != null)
            {
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
        int _timerId;
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
        public void ProcessRunSyncMacroMessage()
        {
            try
            {
                bool runOK = COMRunMacro(_syncMacroName);
                if (!runOK && _timerId == 0)
                {
                    // Timer not yet set - start it
                    // Timer will be stopped when SyncMacro actually runs.
                    IntPtr result = SetTimer(new HandleRef(_syncWindow, _syncWindow.Handle), _timerId++, RETRY_INTERVAL_MS, IntPtr.Zero);
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

        // This is the helper macro that runs (on the main thread)
        void SyncMacro(double _unused_)
        {
            // Check for timer and disable
            if (_timerId != 0)
            {
                KillTimer(new HandleRef(_syncWindow, _syncWindow.Handle), _timerId);
                _timerId = 0;
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

        // Regsiter the helper macro with Excel, so that Application.Run can call it.
        void Register()
        {
            // CONSIDER: Can this be cleaned up by calling ExcelDna.Loader?
            // We must not be in a function when this is run, nor in an RTD method call.
            _syncMacroName = "SyncMacro_" + Guid.NewGuid().ToString("N");
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

        void Unregister()
        {
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
                object xlApp = ExcelDnaUtil.Application;
                xlApp.GetType().InvokeMember("Run", BindingFlags.InvokeMethod, null, xlApp, new object[] { macroName, 0.0 }, _enUsCulture);
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
            // Not releasing the Application object here, since we are on the main thread.
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
                // TODO: case WM_CLOSE / WM_DESTROY: ????
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
