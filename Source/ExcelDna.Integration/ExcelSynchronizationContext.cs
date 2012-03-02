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
    public class ExcelSynchronizationContext : SynchronizationContext
    {
        static SynchronizationWindow _syncWindow;

        // TODO: Do we need reference counts to manage the Register / Unregister?
        public void Register()
        {
            // TODO: Check whether we are running on the main Excel thread.
            if (_syncWindow == null)
            {
                _syncWindow = new SynchronizationWindow();
                _syncWindow.Register();
            }
        }

        public void Unregister()
        {
            _syncWindow.Unregister();
            _syncWindow = null;
        }

        public override void Send(SendOrPostCallback d, object state)
        {
            // We can get SyncMacro to run.
            // SynchronizationWindow.
        }

        public override void Post(SendOrPostCallback d, object state)
        {
            
        }
    }

    // SynchronizationWindow supports running code on the main Excel thread.
    // TODO: Make this into a static class?
    internal sealed class SynchronizationWindow : NativeWindow, IDisposable
    {
        object _lockObject = new object();
        // We'll use the Key collection of a Dictionary as a HashSet (which is not available on .NET 2.0).
        Dictionary<IRTDUpdateEvent, object> _pendingRtdUpdates = new Dictionary<IRTDUpdateEvent,object>();

        [DllImport("user32.dll")]
        static extern bool PostMessage(HandleRef hwnd, int msg, IntPtr wparam, IntPtr lparam);
        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(HandleRef hWnd, int msg, IntPtr wParam, IntPtr lParam);

        readonly IntPtr HWND_MESSAGE = (IntPtr)(-3);
        const int WM_USER = 0x400;
        const int WM_UPDATENOTIFY   = WM_USER + 1;
        const int WM_SYNCMACRO_0 = WM_USER + 2;
        const int WM_SYNCMACRO_1 = WM_USER + 3;

        public SynchronizationWindow()
        {
            CreateParams cp = new CreateParams();
            if (Environment.OSVersion.Version.Major >= 5)
                cp.Parent = HWND_MESSAGE;

            CreateHandle(cp);
        }

        // Support for pushing UpdateNotify onto the main thread.
        // TODO: How do we know that the RTD server is still alive when 
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

        public void RunSyncMacro(SendOrPostCallback d, object state)
        {
            lock (_sendOrPostLock)
            {
                Debug.Assert(_sendCallback == null);
                _sendCallback = d;
                _sendState = state;
                SendMessage(new HandleRef(this, Handle), WM_SYNCMACRO_0, IntPtr.Zero, IntPtr.Zero);
            }
        }

        public void RunSyncMacroAsync(SendOrPostCallback d, object state)
        {
            lock (_sendOrPostLock)
            {
                _postCallbacks.Enqueue(d);
                _postStates.Enqueue(d);
                PostMessage(new HandleRef(this, Handle), WM_SYNCMACRO_1, IntPtr.Zero, IntPtr.Zero);
            }
        }

        void ProcessRunSyncMacro(double dValue)
        {
            // To safely transition to Excel's macro-running context. we 
            object app = ExcelDnaUtil.Application;
            app.GetType().InvokeMember("Run", BindingFlags.InvokeMethod, null, app, new object[] {_syncMacroName, dValue}, new System.Globalization.CultureInfo(1033));
        }

        void ProcessUpdateNotifications()
        {
            lock (_lockObject)
            {
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
                case WM_SYNCMACRO_0:
                    ProcessRunSyncMacro(0);
                    break;
                case WM_SYNCMACRO_1:
                    ProcessRunSyncMacro(1);
                    break;
                default:
                    base.WndProc(ref m);
                    break;
            }
        }

        string _syncMacroName;
        object _syncMacroRegistrationId = null;

        SendOrPostCallback _sendCallback;
        object _sendState;

        Queue<SendOrPostCallback> _postCallbacks = new Queue<SendOrPostCallback>();
        Queue<Object> _postStates = new Queue<Object>();

        object _sendOrPostLock = new object();

        void SyncMacro(double dValue)
        {
            Debug.Print("SyncMacro called on thread: " + Thread.CurrentThread.ManagedThreadId);
            SendOrPostCallback work;
            object state;
            if (dValue == 0.0)
            {
                // SendMessage
                lock (_sendOrPostLock)
                {
                    work = _sendCallback;
                    state = _sendState;
                    _sendCallback = null;
                    _sendState = null;
                }
            }
            else if (dValue == 1.0)
            {
                // PostMessage
                lock (_sendOrPostLock)
                {
                    work = _postCallbacks.Dequeue();
                    state = _postStates.Dequeue();
                }
            }
            else
            {
                Debug.Fail("Unexpected SyncMacro argument: " + dValue);
                return;
            }
            work(state);
        }

        // CONSIDER: Move this to Integration class?
        void RegisterSyncMacro(string name, SyncMacroDelegate syncMacro)
        {
            Integration.SetSyncMacro(syncMacro);

            object[] registerParameters = new object[6];
            registerParameters[0] = DnaLibrary.XllPath;
            registerParameters[1] = "SyncMacro";
            registerParameters[2] = ">B"; // Takes double, abuse in-place flag '>'
            registerParameters[3] = name;
            registerParameters[4] = "value";
            registerParameters[5] = 2; // macro
            _syncMacroRegistrationId = XlCall.Excel(XlCall.xlfRegister, registerParameters);
        }

        // Don't want to call it 'Install' since that is used for installing the SynchronizationContext as the 'Current' SyncContext in the thread.
        public void Register()
        {
            _syncMacroName = "SyncMacro_" + Guid.NewGuid().ToString("N");
            RegisterSyncMacro(_syncMacroName, SyncMacro);
        }

        public void Unregister()
        {
            if (_syncMacroRegistrationId != null)
            {
                XlCall.Excel(XlCall.xlfUnregister, _syncMacroRegistrationId);
                _syncMacroRegistrationId = null;
            }
        }

        ~SynchronizationWindow()
        {
            Dispose(false);
        }

        void Dispose(bool disposing)
        {
            if (disposing)
                DestroyHandle();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}
