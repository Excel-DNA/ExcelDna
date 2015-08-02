//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace ExcelDna.Integration.Rtd
{
    internal static class AsyncObservableImpl
    {
        static readonly Dictionary<AsyncCallInfo, Guid> _asyncCallIds = new Dictionary<AsyncCallInfo, Guid>();
        static readonly Dictionary<Guid, AsyncObservableState> _observableStates = new Dictionary<Guid, AsyncObservableState>();

        // This is the most general RTD registration
        // This should not be called from a ThreadSafe function. Checked in the callers.
        public static object ProcessObservable(string functionName, object parameters, ExcelObservableSource getObservable)
        {
            if (!SynchronizationManager.IsInstalled)
            {
                throw new InvalidOperationException("ExcelAsyncUtil has not been initialized. This is an unexpected error.");
            }
            if (!ExcelDnaUtil.IsMainThread)
            {
                throw new InvalidOperationException("ExcelAsyncUtil.Run / ExcelAsyncUtil.Observe may not be called from a ThreadSafe function.");
            }

            // CONSIDER: Why not same problems with all RTD servers?
            AsyncCallInfo callInfo = new AsyncCallInfo(functionName, parameters);

            // Shortcut if already registered
            Guid id;
            if (_asyncCallIds.TryGetValue(callInfo, out id))
            {
                // Already registered.
                Debug.Print("AsyncObservableImpl GetValueIfRegistered - Found Id: {0}", id);
                AsyncObservableState state = _observableStates[id];
                object value;
                // The TryGetValue call here is a big deal - it eventually calls Excel's RTD function 
                // (or not, it the observable is complete).
                // The return value of TryGetValue indicates the special array-call where RTD fails, which we ignore here.
                bool unused = state.TryGetValue(out value);
                return value;
            }

            // Not registered before - actually register as a new Observable
            IExcelObservable observable = getObservable();
            return RegisterObservable(callInfo, observable);
        }

        // Make a one-shot 'Observable' from the func
        public static object ProcessFunc(string functionName, object parameters, ExcelFunc func)
        {
            return ProcessObservable(functionName, parameters,
                delegate { return new ThreadPoolDelegateObservable(func); });
        }

        public static object ProcessFuncAsyncHandle(string functionName, object parameters, ExcelFuncAsyncHandle func)
        {
            return ProcessObservable(functionName, parameters,
                delegate
                {
                    ExcelAsyncHandleObservable asyncHandleObservable = new ExcelAsyncHandleObservable();
                    func(asyncHandleObservable);
                    return asyncHandleObservable;
                });
        }

        // Register a new observable
        // Returns null if it failed (due to RTD array-caller first call)
        static object RegisterObservable(AsyncCallInfo callInfo, IExcelObservable observable)
        {
            // Check it's not registered already
            Debug.Assert(!_asyncCallIds.ContainsKey(callInfo));

            // Set up ObservableState and keep track of things
            // Caller might be null if not from worksheet
            ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            Guid id = Guid.NewGuid();
            Debug.Print("AsyncObservableImpl.RegisterObservable - Id: {0}", id);
            _asyncCallIds[callInfo] = id;
            AsyncObservableState state = new AsyncObservableState(id, callInfo, caller, observable);
            _observableStates[id] = state;

            // Will spin up RTD server and topic if required, causing us to be called again...
            object value;
            if (!state.TryGetValue(out value))
            {
                Debug.Print("AsyncObservableImpl.RegisterObservable (GetValue Error) - Remove Id: {0}", id);
                // Problem case - array-caller with RTD call that failed.
                // Clean up state and return null - we'll be called again later and everything will be better.
                _observableStates.Remove(id);
                _asyncCallIds.Remove(callInfo);
                return null;
            }
            return value;
        }

        // Safe to call with an invalid Id, but that's not expected.
        internal static void ConnectObserver(Guid id, ExcelRtdObserver rtdObserver)
        {
            Debug.Print("AsyncObservableImpl.ConnectObserver - Id: {0}", id);

            // It's an error if the id is not on the list - the ExcelObserverRtdServer should protect is from the onw known case 
            // - when RTD is called from sheet open
            AsyncObservableState state;
            if (_observableStates.TryGetValue(id, out state))
            {
                // Start the work for this AsyncCallInfo, and subscribe the topic to the result
                state.Subscribe(rtdObserver);
            }
            else
            {
                // Not expected - the ExcelObserverRtdServer should have protected us, 
                // since the only invalid id call would be for sheet open with direct RTD refresh.
                Debug.Fail("AsyncObservableImpl.ConnectObserver - Invalid Id: " + id);
            }
        }

        // Safe to call even if the id is not valid
        internal static void DisconnectObserver(Guid id)
        {
            Debug.Print("AsyncObservableImpl.DisconnectObserver - Id: {0}", id);

            AsyncObservableState state;
            if (_observableStates.TryGetValue(id, out state))
            {
                state.Unsubscribe();
                _observableStates.Remove(id);   // Remove is safe even if key is not found.
                _asyncCallIds.Remove(state.GetCallInfo());
            }
        }
    }

    // Pushes data to Excel via the RTD topic.
    // Observes a single Observable. 
    // Needs to coordinate the IsCompleted status with all the other Observers for the Caller, to ensure coordinated 'completion'.

    internal class ExcelRtdObserver : IExcelObserver
    {
        readonly ExcelRtdServer.Topic _topic;

        // Indicates whether the RTD Topic should be shut down
        // Set to true if the work is completed or if an error is signalled.
        public bool IsCompleted { get; private set; }
        // Keeping our own value, since the RTD Topic.Value is insufficient (e.g. string length limitation)
        // This may be an issue if we want to support startup OldValues (currently we don't)
        public object Value { get; private set; }

        internal ExcelRtdObserver(ExcelRtdServer.Topic topic)
        {
            _topic = topic;
            Value = ExcelError.ExcelErrorNA;
        }

        public void OnCompleted()
        {
            IsCompleted = true;
            // Force another update to ensure DisconnectData is called.
            // CONSIDER: Do we need to UpdateNotify here?
            //           Not necessarily. Next recalc will call the function, not call RTD and that will trigger DisconnectData.
            //           However, this ensures a more deterministic call to DisconnectData
            _topic.UpdateNotify();
        }

        public void OnError(Exception exception)
        {
            Value = ExcelIntegration.HandleUnhandledException(exception);
            // Set the topic value to #VALUE (not used!?) - converted to COM code in Topic
            _topic.UpdateValue(ExcelError.ExcelErrorValue);
            OnCompleted();
        }

        public void OnNext(object value)
        {
            Value = value;
            // Not actually setting the topic value, just poking it
            // TODO: Using the 'fake' RTD value should be optional - to give us a way to deal with 'newValues' one day.
            // BUGBUG: The ToOADate truncates, things might happy in the same millisecond etc.
            //         See https://exceldna.codeplex.com/workitem/9472
            //_topic.UpdateValue(DateTime.UtcNow.ToOADate());
            _topic.UpdateNotify();
        }
    }

    [ComVisible(true)]
    internal class ExcelObserverRtdServer : ExcelRtdServer
    {
        class ObserverRtdTopic : Topic
        {
            public readonly Guid Id;
            public ObserverRtdTopic(ExcelObserverRtdServer server, int topicId, Guid id)
                : base(server, topicId)
            {
                Id = id;
            }
        }

        protected override Topic CreateTopic(int topicId, IList<string> topicInfo)
        {
            Guid id = new Guid(topicInfo[0]);
            return new ObserverRtdTopic(this, topicId, id);
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            Debug.Print("ExcelObserverRtdServer.ConnectData: ProgId: {0}, TopicId: {1}, TopicInfo: {2}, NewValues: {3}", RegisteredProgId, topic.TopicId, topicInfo[0], newValues);

            if (newValues == false)
            {
                // Excel has a cached value, and we are being called from the file open refresh.

                // Calling UpdateNotify here seems to work (it causes the wrapper function to recalc, 
                //    which Disconnects the bad topic, and allows a fresh one to be created)
                // Not needed if we return a new 'fake' value, which is safe since it is consistent with normal updates.
                // Result should be a Disconnect followed by a proper Connect via the wrapper.
                // topic.UpdateNotify();   

                newValues = true;
                return DateTime.UtcNow.ToOADate();
            }
            // Retrieve and store the GUID from the topic's first info string - used to hook up to the Async state
            Guid id = ((ObserverRtdTopic)topic).Id;

            // Create a new ExcelRtdObserver, for the Topic, which will listen to the Observable
            // (Internally this will also set the initial value to #N/A)
            ExcelRtdObserver rtdObserver = new ExcelRtdObserver(topic);
            // ... and subscribe it
            AsyncObservableImpl.ConnectObserver(id, rtdObserver);

            // Now ConnectData needs to return some value, which will only be used by Excel internally (and saved in the book's RTD topic value).
            // Our wrapper function (ExcelAsyncUtil.Run or ExcelAsyncUtil.Observe) will return #N/A no matter what we return here.
            // However, it seems that Excel handles the special 'busy' error #N/A here (return ExcelErrorUtil.ToComError(ExcelError.ExcelErrorNA))
            // in a special way (<tp t="e"><v>#N/A</v> in volatileDependencies.xml) - while other values seem to trigger a recalculate on file open, 
            // when Excel attempts to restart the RTD server and fails (due to transient ProgId).
            // So we already return the same kind of value we'd return for updates, putting Excel into the 'value has been updated' state
            // even if the sheet is saved. That will trigger a proper formula recalcs on file open.
            return DateTime.UtcNow.ToOADate();
        }

        protected override void DisconnectData(Topic topic)
        {
            Debug.Print("ExcelObserverRtdServer.DisconnectData: ProgId: {0}, TopicId: {1}", RegisteredProgId, topic.TopicId);

            // Retrieve the GUID from the topic's first info string - used to hook up to the Async state
            Guid id = ((ObserverRtdTopic)topic).Id;
            // ... and unsubscribe it
            AsyncObservableImpl.DisconnectObserver(id);
        }

#if DEBUG
        protected override bool ServerStart()
        {
            Debug.Print("ExcelObserverRtdServer.ServerStart");
            return true;
        }

        protected override void ServerTerminate()
        {
            Debug.Print("ExcelObserverRtdServer.ServerTerminate");
        }
#endif

        // This makes sure the hook up with the registration-free RTD loading is in place.
        // For a user RTD server the add-in loading would ensure this, but not for this class since it is declared inside Excel-DNA.
        static bool _isRegistered = false;
        internal static void EnsureRtdServerRegistered()
        {
            if (!_isRegistered)
            {
                RtdRegistration.RegisterRtdServerTypes(new Type[] { typeof(ExcelObserverRtdServer) });
            }
            _isRegistered = true;
        }
    }

    // Encapsulates the information that defines and async call or observable hook-up.
    // Checked for equality and stored in a Dictionary, so we have to be careful
    // to define value equality and a consistent HashCode.

    // Used as Keys in a Dictionary - should be immutable. 
    // We allow parameters to be null or primitives or ExcelReference objects, 
    // or a 1D array or 2D array with primitives or arrays.
    internal struct AsyncCallInfo : IEquatable<AsyncCallInfo>
    {
        readonly string _functionName;
        readonly object _parameters;
        readonly int _hashCode;

        public AsyncCallInfo(string functionName, object parameters)
        {
            _functionName = functionName;
            _parameters = parameters;
            _hashCode = 0; // Need to set to some value before we call a method.
            _hashCode = ComputeHashCode();
        }

        // Jon Skeet: http://stackoverflow.com/questions/263400/what-is-the-best-algorithm-for-an-overridden-system-object-gethashcode
        int ComputeHashCode()
        {
            unchecked
            {
                int hash = 17;
                hash = hash * 23 + (_functionName == null ? 0 : _functionName.GetHashCode());
                hash = hash * 23 + ComputeHashCode(_parameters);
                return hash;
            }
        }

        // Computes a hash code for the parameters, consistent with the value equality that we define.
        // Also checks that the data types passed for parameters are among those we handle properly for value equality.
        // For now no string[]. But invalid types passed in will causes exception immediately.
        static int ComputeHashCode(object obj)
        {
            if (obj == null) return 0;

            // CONSIDER: All of this could be replaced by a check for (obj is ValueType || obj is ExcelReference)
            //           which would allow a more flexible set of parameters, at the risk of comparisons going wrong.
            //           We can reconsider if this arises, or when we implement async automatically or custom marshaling 
            //           to other data types. For now this allows everything that can be passed as parameters from Excel-DNA.

            // We also support using an opaque byte[] hash as the parameters 'key'.
            // In cases with huge amounts of active topics, especially using string parameters, this can improve the memory usage significantly.

            if (obj is double ||
                obj is float ||
                obj is string ||
                obj is bool ||
                obj is DateTime ||
                obj is ExcelReference ||
                obj is ExcelError ||
                obj is ExcelEmpty ||
                obj is ExcelMissing ||
                obj is int ||
                obj is uint ||
                obj is long ||
                obj is ulong ||
                obj is short ||
                obj is ushort ||
                obj is byte ||
                obj is sbyte ||
                obj is decimal ||
                obj.GetType().IsEnum)
            {
                return obj.GetHashCode();
            }

            unchecked
            {
                int hash = 17;

                double[] doubles = obj as double[];
                if (doubles != null)
                {
                    foreach (double item in doubles)
                    {
                        hash = hash * 23 + item.GetHashCode();
                    }
                    return hash;
                }

                double[,] doubles2 = obj as double[,];
                if (doubles2 != null)
                {
                    foreach (double item in doubles2)
                    {
                        hash = hash * 23 + item.GetHashCode();
                    }
                    return hash;
                }

                object[] objects = obj as object[];
                if (objects != null)
                {
                    foreach (object item in objects)
                    {
                        hash = hash * 23 + ((item == null) ? 0 : ComputeHashCode(item));
                    }
                    return hash;
                }

                object[,] objects2 = obj as object[,];
                if (objects2 != null)
                {
                    foreach (object item in objects2)
                    {
                        hash = hash * 23 + ((item == null) ? 0 : ComputeHashCode(item));
                    }
                    return hash;
                }

                byte[] bytes = obj as byte[];
                if (bytes != null)
                {
                    foreach (byte b in bytes)
                    {
                        hash = hash * 23 + b;
                    }
                    return hash;
                }
            }
            throw new ArgumentException("Invalid type used for async parameter(s)", "parameters");
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (obj.GetType() != typeof(AsyncCallInfo)) return false;
            return Equals((AsyncCallInfo)obj);
        }

        public bool Equals(AsyncCallInfo other)
        {
            if (_hashCode != other._hashCode) return false;
            return Equals(other._functionName, _functionName)
                   && ValueEquals(_parameters, other._parameters);
        }

        #region Helpers to implement value equality
        // The value equality we check here is for the types we allow in CheckParameterTypes above.
        static bool ValueEquals(object a, object b)
        {
            if (Equals(a, b)) return true; // Includes check for both null
            if (a is double[] && b is double[]) return ArrayEquals((double[])a, (double[])b);
            if (a is double[,] && b is double[,]) return ArrayEquals((double[,])a, (double[,])b);
            if (a is object[] && b is object[]) return ArrayEquals((object[])a, (object[])b);
            if (a is object[,] && b is object[,]) return ArrayEquals((object[,])a, (object[,])b);
            if (a is byte[] && b is byte[]) return ArrayEquals((byte[])a, (byte[])b);
            return false;
        }

        static bool ArrayEquals(double[] a, double[] b)
        {
            if (a.Length != b.Length)
                return false;
            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] != b[i]) return false;
            }
            return true;
        }

        static bool ArrayEquals(double[,] a, double[,] b)
        {
            int rows = a.GetLength(0);
            int cols = a.GetLength(1);
            if (rows != b.GetLength(0) ||
                cols != b.GetLength(1))
            {
                return false;
            }
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (a[i, j] != b[i, j]) return false;
                }
            }
            return true;
        }

        static bool ArrayEquals(object[] a, object[] b)
        {
            if (a.Length != b.Length)
                return false;
            for (int i = 0; i < a.Length; i++)
            {
                if (!ValueEquals(a[i], b[i]))
                    return false;
            }
            return true;
        }

        static bool ArrayEquals(object[,] a, object[,] b)
        {
            int rows = a.GetLength(0);
            int cols = a.GetLength(1);
            if (rows != b.GetLength(0) ||
                cols != b.GetLength(1))
            {
                return false;
            }
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (!ValueEquals(a[i, j], b[i, j]))
                        return false;
                }
            }
            return true;
        }

        static bool ArrayEquals(byte[] a, byte[] b)
        {
            if (a.Length != b.Length)
                return false;
            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] != b[i]) return false;
            }
            return true;
        }

        #endregion

        public override int GetHashCode()
        {
            return _hashCode;
        }

        public static bool operator ==(AsyncCallInfo asyncCallInfo1, AsyncCallInfo asyncCallInfo2)
        {
            return asyncCallInfo1.Equals(asyncCallInfo2);
        }

        public static bool operator !=(AsyncCallInfo asyncCallInfo1, AsyncCallInfo asyncCallInfo2)
        {
            return !(asyncCallInfo1.Equals(asyncCallInfo2));
        }
    }

    // This manages the information for a single Caller (maybe multiple UDF+callinfos)
    // Added to allow IsComplete synchronization per Caller
    // We expect all calls into the class to be on the main thread
    internal class AsyncCallerState
    {
        static readonly Dictionary<ExcelReference, AsyncCallerState> _callerStates = new Dictionary<ExcelReference, AsyncCallerState>();
        // caller might be null
        public static AsyncCallerState GetCallerState(ExcelReference caller)
        {
            if (caller == null) return new AsyncCallerState(null);

            AsyncCallerState callerState;
            if (!_callerStates.TryGetValue(caller, out callerState))
            {
                callerState = new AsyncCallerState(caller);
                _callerStates[caller] = callerState;
            }
            return callerState;
        }

        readonly ExcelReference _caller; // Might be null
        readonly List<ExcelRtdObserver> _observers = new List<ExcelRtdObserver>();
        AsyncCallerState(ExcelReference caller)
        {
            _caller = caller;
        }

        public void AddObserver(ExcelRtdObserver observer)
        {
            _observers.Add(observer);
        }

        public void RemoveObserver(ExcelRtdObserver observer)
        {
            _observers.Remove(observer);
            if (_observers.Count == 0 && _caller != null)
            {
                _callerStates.Remove(_caller);
            }
        }

        // Called on the main thread
        public bool AreObserversCompleted()
        {
            foreach (ExcelRtdObserver observer in _observers)
            {
                if (!observer.IsCompleted) return false;
            }
            return true;
        }
    }

    // This manages the information for a single Observable (one UDF+callinfo).
    internal class AsyncObservableState
    {
        const string _observerRtdServerProgId = "ExcelDna.Integration.Rtd.ExcelObserverRtdServer";
        readonly string _id;
        readonly AsyncCallerState _callerState;
        readonly AsyncCallInfo _callInfo; // Bit ugly having this here - need a bi-directional dictionary...
        readonly IExcelObservable _observable;
        ExcelRtdObserver _currentObserver;
        IDisposable _currentSubscription;

        // caller may be null when not called as a worksheet function.
        public AsyncObservableState(Guid id, AsyncCallInfo callInfo, ExcelReference caller, IExcelObservable observable)
        {
            _id = id.ToString();
            _callInfo = callInfo;
            _observable = observable;
            _callerState = AsyncCallerState.GetCallerState(caller); // caller might be null, _callerState should not be
        }

        public bool TryGetValue(out object value)
        {
            // We need to be careful when this is called from an array formula.
            // In the 'completed' case we actually still have to call xlfRtd, then only skip if for the next (single-cell calller) call.
            // That gives us a proper Disconnect...
            ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            bool isCallerArray = caller != null &&
                                 (caller.RowFirst != caller.RowLast ||
                                  caller.ColumnFirst != caller.ColumnLast);
            if (_currentObserver == null || isCallerArray || !IsCompleted())
            {
                // NOTE: At this point the SynchronizationManager must be registered!
                if (!SynchronizationManager.IsInstalled)
                {
                    Debug.Print("SynchronizationManager not registered!");
                    throw new InvalidOperationException("SynchronizationManager must be registered for async and observable support. Call ExcelAsyncUtil.Initialize() in an IExcelAddIn.AutoOpen() handler.");
                }

                // Ensure that Excel-DNA knows about the RTD Server, since it would not have been registered when loading
                ExcelObserverRtdServer.EnsureRtdServerRegistered();

                // Refresh RTD call
                // NOTE: First time this will result in a call to ConnectData, which will call Subscribe and set the _currentObserver
                //       For the first array-group call, this returns null (due to xlUncalced error), 
                //       but Excel will call us again... (I hope).
                if (!RtdRegistration.TryRTD(out value, _observerRtdServerProgId, null, _id))
                {
                    // This is the special case...
                    // We return false - to the state creation function that indicates the state should not be saved.
                    value = ExcelError.ExcelErrorNA;
                    return false;
                }
            }
            else if (_currentObserver != null && IsCompleted())
            {
                // Special call for the Excel 2010 bug helper to indicate we are not refreshing (due to completion)
                if (ExcelRtd2010BugHelper.ExcelVersionHasRtdBug)
                {
                    ExcelRtd2010BugHelper.RecordRtdComplete(_observerRtdServerProgId, _id);
                }
            }

            // No assumptions about previous state here - could have re-entered this class

            // We use #N/A as the 'busy' indicator, as RTD does normally.
            // Add-in creator can remap the 'busy' result in the UDF or another wrapper.
            if (_currentObserver == null)
            {
                value = ExcelError.ExcelErrorNA;
                return true;
            }

            // Subsequent calls get value from Observer
            value = _currentObserver.Value;
            return true;
        }

        public void Subscribe(ExcelRtdObserver rtdObserver)
        {
            _currentObserver = rtdObserver;
            _callerState.AddObserver(_currentObserver);
            _currentSubscription = _observable.Subscribe(rtdObserver);
        }

        // Under unpatched Excel 2010, we rely on the ExcelRtd2010BugHelper to ensure we get a good Unsubscribe...
        public void Unsubscribe()
        {
            Debug.Assert(_currentSubscription != null);
            _currentSubscription.Dispose();
            _currentSubscription = null;
            _callerState.RemoveObserver(_currentObserver);
            _currentObserver = null;
        }

        public AsyncCallInfo GetCallInfo()
        {
            return _callInfo;
        }

        bool IsCompleted()
        {
            if (!_currentObserver.IsCompleted) return false;
            return _callerState.AreObserversCompleted();
        }
    }

    // This is not a very elegant IObservable implementation - should not be public.
    // It basically represents a Subject 
    internal class ThreadPoolDelegateObservable : IExcelObservable
    {
        readonly ExcelFunc _func;
        bool _subscribed;

        public ThreadPoolDelegateObservable(ExcelFunc func)
        {
            _func = func;
        }

        public IDisposable Subscribe(IExcelObserver observer)
        {
            if (_subscribed) throw new InvalidOperationException("Only single Subscription allowed.");
            _subscribed = true;

            ThreadPool.QueueUserWorkItem(delegate
            {
                try
                {
                    object result = _func();
                    observer.OnNext(result);
                    observer.OnCompleted();
                }
                catch (Exception ex)
                {
                    // TODO: Log somehow?
                    observer.OnError(ex);
                }
            });

            return DummyDisposable.Instance;
        }

        class DummyDisposable : IDisposable
        {
            public static readonly DummyDisposable Instance = new DummyDisposable();

            private DummyDisposable()
            {
            }

            public void Dispose()
            {
            }
        }
    }
}