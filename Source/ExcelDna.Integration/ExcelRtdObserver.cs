//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelDna.Integration.Rtd
{
    // ThreadSafe
    internal static class AsyncObservableImpl
    {
        static readonly ConcurrentDictionary<AsyncCallInfo, Guid> _asyncCallIds = new ConcurrentDictionary<AsyncCallInfo, Guid>();
        static readonly ConcurrentDictionary<Guid, AsyncObservableState> _observableStates = new ConcurrentDictionary<Guid, AsyncObservableState>();
        static readonly object _registrationLock = new object();

        // This is the most general RTD registration
        public static object ProcessObservable(string functionName, object parameters, ExcelObservableOptions options, ExcelObservableSource getObservable)
        {
            if (!SynchronizationManager.IsInstalled)
            {
                throw new InvalidOperationException("ExcelAsyncUtil has not been initialized. This is an unexpected error.");
            }

            // CONSIDER: Why not same problems with all RTD servers?
            AsyncCallInfo callInfo = new AsyncCallInfo(functionName, parameters);
            AsyncObservableState state = null;
            lock (_registrationLock)
            {
                Guid id;
                if (_asyncCallIds.TryGetValue(callInfo, out id))
                {
                    // Already registered.
                    Debug.Print("AsyncObservableImpl GetValueIfRegistered - Found Id: {0}", id);
                    state = _observableStates[id];
                }
            }

            // Shortcut if already registered
            if (state != null)
            {
                object value;
                // The TryGetValue call here is a big deal - it eventually calls Excel's RTD function 
                // (or not, it the observable is complete).
                // The return value of TryGetValue indicates the special array-call where RTD fails, which we ignore here.
                bool unused = state.TryGetValue(out value);
                return value;
            }

            // Not registered before - actually register as a new Observable
            IExcelObservable observable = getObservable();
            return RegisterObservable(callInfo, options, observable);
        }

        // Make a one-shot 'Observable' from the func
        public static object ProcessFunc(string functionName, object parameters, ExcelFunc func)
        {
            return ProcessObservable(functionName, parameters, ExcelObservableOptions.None,
                delegate { return new ThreadPoolDelegateObservable(func); });
        }

        public static object ProcessFuncAsyncHandle(string functionName, object parameters, ExcelFuncAsyncHandle func)
        {
            return ProcessObservable(functionName, parameters, ExcelObservableOptions.None,
                delegate
                {
                    ExcelAsyncHandleObservable asyncHandleObservable = new ExcelAsyncHandleObservable();
                    func(asyncHandleObservable);
                    return asyncHandleObservable;
                });
        }

        // Register a new observable
        // Returns null if it failed (due to RTD array-caller first call)
        static object RegisterObservable(AsyncCallInfo callInfo, ExcelObservableOptions options, IExcelObservable observable)
        {
            AsyncObservableState state;
            Guid id;
            lock (_registrationLock)
            {
                // Check it's not registered already
                if (_asyncCallIds.TryGetValue(callInfo, out id))
                {
                    // Already registered.
                    Debug.Print("AsyncObservableImpl GetValueIfRegistered - Found Id: {0}", id);
                    state = _observableStates[id];
                }
                else
                {
                    // Set up ObservableState and keep track of things
                    // Caller might be null if not from worksheet
                    ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
                    id = Guid.NewGuid();
                    Debug.Print("AsyncObservableImpl.RegisterObservable - Id: {0}", id);
                    _asyncCallIds[callInfo] = id;
                    state = new AsyncObservableState(id, callInfo, options, caller, observable);
                    _observableStates[id] = state;
                }
            }

            // Will spin up RTD server and topic if required, causing us to be called again...
            object value;
            if (!state.TryGetValue(out value))
            {
                Debug.Print("AsyncObservableImpl.RegisterObservable (GetValue Error) - Remove Id: {0}", id);
                // Problem case - array-caller with RTD call that failed.
                // Clean up state and return null - we'll be called again later and everything will be better.
                lock (_registrationLock)
                {
                    _observableStates.TryRemove(id, out _);
                    _asyncCallIds.TryRemove(callInfo, out _);
                }
                return null;
            }
            return value;
        }

        // Safe to call with an invalid Id, but that's not expected.
        internal static void ConnectObserver(Guid id, IExcelRtdObserver rtdObserver)
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

            lock (_registrationLock)
            {
                AsyncObservableState state;
                if (_observableStates.TryGetValue(id, out state))
                {
                    state.Unsubscribe();
                    _observableStates.TryRemove(id, out _);   // Remove is safe even if key is not found.
                    _asyncCallIds.TryRemove(state.GetCallInfo(), out _);
                }
            }
        }

        internal static ExcelObservableOptions GetObservableOptions(Guid id)
        {
            if (_observableStates.TryGetValue(id, out AsyncObservableState state))
            {
                return state.GetOptions();
            }

            return ExcelObservableOptions.None;
        }
    }

    internal interface IExcelRtdObserver : IExcelObserver
    {
        object GetValue();
        bool IsCompleted();
    }

    internal class TopicUpdater
    {
        readonly ExcelRtdServer.Topic _topic;

        // 2^52 -1 - we'll roll over the TopicValue here, so that the double representation of the incrementing integer is always unique
        const long MaxTopicValue = 4503599627370495;

        public TopicUpdater(ExcelRtdServer.Topic topic)
        {
            _topic = topic;
        }

        public void Update()
        {
            _topic.UpdateValue((((double)_topic.Value) + 1) % MaxTopicValue);
        }

        public void Complete()
        {
            _topic.UpdateValue(ExcelObserverRtdServer.TopicValueCompleted);
        }
    }

    // Pushes data to Excel via the RTD topic.
    // Observes a single Observable. 
    // Needs to coordinate the IsCompleted status with all the other Observers for the Caller, to ensure coordinated 'completion'.
    internal class ExcelRtdObserver : IExcelRtdObserver
    {
        private TopicUpdater _topicUpdater;

        // Indicates whether the RTD Topic should be shut down
        // Set to true if the work is completed or if an error is signalled.
        private bool _isCompleted;

        // Keeping our own value, since the RTD Topic.Value is insufficient (e.g. string length limitation)
        // This may be an issue if we want to support startup OldValues (currently we don't)
        private object _value;

        internal ExcelRtdObserver(ExcelRtdServer.Topic topic)
        {
            _topicUpdater = new TopicUpdater(topic);

            // Set our wrapper Value, but not the internal Topic value 
            // (which must never be #N/A if we want re-open restart).
            _value = ExcelError.ExcelErrorNA;
        }

        public void OnCompleted()
        {
            _isCompleted = true;

            // 2016-11-04: We have to ensure that the UpdateNotify call is still called if 
            //             OnCompleted happens during the ConnectData
            //             To ensure this we set the internal topic value

            _topicUpdater.Complete();
        }

        public void OnError(Exception exception)
        {
            _value = ExcelIntegration.HandleUnhandledException(exception);
            OnCompleted();
        }

        public void OnNext(object value)
        {
            _value = value;

            // This code has had  alot of churn
            // Old versions set the Topic Value, then many versions did not
            // From v1.1 (2020-03-02) we set it again to a dummy value

            // TODO: Using the 'fake' RTD value should be optional - to give us a way to deal with 'newValues' one day.
            //       But then we'd need to be really careful with error values (which prevent restart).
            // BUGBUG: The ToOADate truncates, things might happy in the same millisecond etc.
            //         See https://exceldna.codeplex.com/workitem/9472
            //_topic.UpdateValue(DateTime.UtcNow.ToOADate());

            // 2016-03-25: Further bug here - We can't leave the Topic value as #N/A (which is what happens if we don't update it)
            //             since that prevents restart when a book is re-opened. (See details in ExcelObserverRtdServer.ConnectData)
            //             So we now initialize the Topic with a new value.

            // 2016-11-04: ExcelRtdServer.ConnectData was changed so that this call, 
            //             if it happens during the ConnectData, will have no effect (saving an extra function call).

            // 2020-03-02: Now Excel's behaviour seems to have changed, and without an updated value the topic does not seem to update even if present in RefreshData
            //             Previously we just added the topic to RefreshData:
            //                  _topic.UpdateNotify();
            //             Now we really update again, but with something more careful than before to not have the millisecond problem
            //             We're not worried about a race condition here - multiple thread might update at once, but we just need one to succeed in setting a new value

            _topicUpdater.Update();
        }

        public object GetValue()
        {
            return _value;
        }

        public bool IsCompleted()
        {
            return _isCompleted;
        }
    }

    internal class ExcelRtdLosslessObserver : IExcelRtdObserver
    {
        private TopicUpdater _topicUpdater;
        private bool _isCompleted;
        private Queue<object> _values = new Queue<object>();
        private object _lastValue;

        internal ExcelRtdLosslessObserver(ExcelRtdServer.Topic topic)
        {
            _topicUpdater = new TopicUpdater(topic);
            _lastValue = ExcelError.ExcelErrorNA;
        }

        public void OnCompleted()
        {
            if (_isCompleted)
                return;

            _isCompleted = true;

            if (IsCompleted())
                _topicUpdater.Complete();
        }

        public void OnError(Exception exception)
        {
            _lastValue = ExcelIntegration.HandleUnhandledException(exception);
            _values.Clear();
            _isCompleted = true;

            _topicUpdater.Complete();
        }

        public void OnNext(object value)
        {
            _values.Enqueue(value);
            _topicUpdater.Update();
        }

        public object GetValue()
        {
            if (_values.Count > 0)
            {
                _lastValue = _values.Dequeue();
                _topicUpdater.Update();

                if (IsCompleted())
                    _topicUpdater.Complete();
            }

            return _lastValue;
        }

        public bool IsCompleted()
        {
            return _isCompleted && (_values.Count == 0);
        }
    }

    [ComVisible(true)]
    internal class ExcelObserverRtdServer : ExcelRtdServer
    {
        // The topic has an internal value that is returned to Excel in the RTD UpdateValues,
        // but never returned by us to the UDF wrapper.
        // This internal value is stored by Excel in the volatileDependencies.xml part.
        // If it is the default #N/A error value, then the topic is not restarted when re-opening the sheet.
        // So we never want to be in the #N/A state.
        // CONSIDER: What do we do use for the alternative values - something more obviously invalid?

        // 2020-03-02: It seems Excel no longer updates if the RefreshData returns the same value.
        //             So for regular updates we will set the internal topic value to a new value on every update

        internal static readonly object TopicValueInitial = 1.0;    // We'll increment from here with every new value
        internal static readonly object TopicValueCompleted = -2.0; // We park here after OnCompleted

        class ObserverRtdTopic : Topic
        {
            public readonly Guid Id;

            // NOTE: That the topic is initialized with the value TopicValueActive is 
            //       important to the ConnectData implementation below
            //       - there is some interaction between topic values and the return value from ConnectData.
            public ObserverRtdTopic(ExcelObserverRtdServer server, int topicId, Guid id, object valueActive)
                : base(server, topicId, valueActive)
            {
                Id = id;
            }
        }

        protected override Topic CreateTopic(int topicId, IList<string> topicInfo)
        {
            Guid id = new Guid(topicInfo[0]);
            return new ObserverRtdTopic(this, topicId, id, TopicValueInitial);
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            Debug.Print("ExcelObserverRtdServer.ConnectData: ProgId: {0}, TopicId: {1}, TopicInfo: {2}, NewValues: {3}", RegisteredProgId, topic.TopicId, topicInfo[0], newValues);

            // The topic might be "completed" or incremented on a separate thread, so we need to return the initial, intended active value
            object initialValue = TopicValueInitial;

            if (newValues == false)
            {
                // Excel has a cached value, and we are being called from the file open refresh.

                // Indicating "newValues", should be safe since it is consistent with normal updates.
                // Result should be a Disconnect followed by a proper Connect via the wrapper.

                newValues = true;
                return initialValue;
            }
            // Retrieve and store the GUID from the topic's first info string - used to hook up to the Async state
            Guid id = ((ObserverRtdTopic)topic).Id;

            ExcelObservableOptions observableOptions = AsyncObservableImpl.GetObservableOptions(id);

            // Create a new ExcelRtdObserver, for the Topic, which will listen to the Observable
            // (Internally this will also set the initial value of the Observer wrapper to #N/A)
            IExcelRtdObserver rtdObserver = observableOptions.HasFlag(ExcelObservableOptions.Lossless) ? (IExcelRtdObserver)new ExcelRtdLosslessObserver(topic) : (IExcelRtdObserver)new ExcelRtdObserver(topic);
            // ... and subscribe it
            AsyncObservableImpl.ConnectObserver(id, rtdObserver);

            // Now ConnectData needs to return some value, which will only be used by Excel internally (and saved in the book's RTD topic value).
            // Our wrapper function (ExcelAsyncUtil.Run or ExcelAsyncUtil.Observe) will return #N/A no matter what we return here.
            // However, it seems that Excel handles the special 'busy' error #N/A here (return ExcelErrorUtil.ToComError(ExcelError.ExcelErrorNA))
            // in a special way (<tp t="e"><v>#N/A</v> in volatileDependencies.xml) - while other values seem to trigger a recalculate on file open, 
            // when Excel attempts to restart the RTD server and fails (due to transient ProgId).
            // So for the ObserverRtdTopic we ensure the internal value is not an error,
            // (it is initialized to TopicValueActive)
            // which we return from here.

            // 2016-11-04: We are no longer returning the current value of topic.Value here.
            //             Since calls to UpdateValue inside the ConnectData no longer raise an
            //             UpdateNotify automatically, we need to ensure a different value
            //             is returned for a completed topic (so that ConnectData.returned != topic.Value)
            //             to raise an extra UpdateNotify, for the Disconnect of the already completed topic
            //             (I.e. if the completion happened during the ConnectData call).

            return initialValue;
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
        static object _isRegisteredLock = new object();
        internal static void EnsureRtdServerRegistered()
        {
            lock (_isRegisteredLock)
            {
                if (!_isRegistered)
                {
                    RtdRegistration.RegisterRtdServerTypes(new Type[] { typeof(ExcelObserverRtdServer) });
                }
                _isRegistered = true;
            }
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
    // ThreadSafe
    internal class AsyncCallerState
    {
        static readonly Dictionary<ExcelReference, AsyncCallerState> _callerStates = new Dictionary<ExcelReference, AsyncCallerState>();
        static object _callerStatesLock = new object();
        // caller might be null
        public static AsyncCallerState GetCallerState(ExcelReference caller)
        {
            if (caller == null) return new AsyncCallerState(null);

            AsyncCallerState callerState;
            lock (_callerStatesLock)
            {
                if (!_callerStates.TryGetValue(caller, out callerState))
                {
                    callerState = new AsyncCallerState(caller);
                    _callerStates[caller] = callerState;
                }
            }
            return callerState;
        }

        readonly ExcelReference _caller; // Might be null
        readonly List<IExcelRtdObserver> _observers = new List<IExcelRtdObserver>();
        AsyncCallerState(ExcelReference caller)
        {
            _caller = caller;
        }

        public void AddObserver(IExcelRtdObserver observer)
        {
            _observers.Add(observer);
        }

        public void RemoveObserver(IExcelRtdObserver observer)
        {
            _observers.Remove(observer);
            if (_observers.Count == 0 && _caller != null)
            {
                lock (_callerStatesLock)
                {
                    _callerStates.Remove(_caller);
                }
            }
        }

        // Called on the main thread
        public bool AreObserversCompleted()
        {
            foreach (IExcelRtdObserver observer in _observers)
            {
                if (!observer.IsCompleted()) return false;
            }
            return true;
        }
    }

    // This manages the information for a single Observable (one UDF+callinfo).
    // ThreadSafe
    internal class AsyncObservableState
    {
        const string _observerRtdServerProgId = "ExcelDna.Integration.Rtd.ExcelObserverRtdServer";
        readonly AsyncCallerState _callerState;
        readonly AsyncCallInfo _callInfo; // Bit ugly having this here - need a bi-directional dictionary...
        readonly string[] _topics; // Contains id (Guid.ToString()) and possibly Integer retpresentation of the ExcelObservableOptions enum value
        readonly IExcelObservable _observable;
        IExcelRtdObserver _currentObserver;
        IDisposable _currentSubscription;
        readonly object _lock = new object();
        ExcelObservableOptions _options;

        // caller may be null when not called as a worksheet function.
        public AsyncObservableState(Guid id, AsyncCallInfo callInfo, ExcelObservableOptions options, ExcelReference caller, IExcelObservable observable)
        {
            _callInfo = callInfo;
            _options = options;
            _topics = new string[] { id.ToString() };
            _observable = observable;
            _callerState = AsyncCallerState.GetCallerState(caller); // caller might be null, _callerState should not be
        }

        public bool TryGetValue(out object value)
        {
            // We need to be careful when this is called from an array formula.
            // In the 'completed' case we actually still have to call xlfRtd, then only skip if for the next (single-cell caller) call.
            // That gives us a proper Disconnect...
            ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            bool isCallerArray = caller != null &&
                                 (caller.RowFirst != caller.RowLast ||
                                  caller.ColumnFirst != caller.ColumnLast);

            bool refreshRTDCall;
            lock (_lock)
            {
                refreshRTDCall = (_currentObserver == null || isCallerArray || !IsCompleted());
            }

            if (refreshRTDCall)
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
                if (!RtdRegistration.TryRTD(out value, _observerRtdServerProgId, null, _topics))
                {
                    // This is the special case...
                    // We return false - to the state creation function that indicates the state should not be saved.
                    value = ExcelError.ExcelErrorNA;
                    return false;
                }
            }

            // No assumptions about previous state here - could have re-entered this class

            lock (_lock)
            {
                // We use #N/A as the 'busy' indicator, as RTD does normally.
                // Add-in creator can remap the 'busy' result in the UDF or another wrapper.
                if (_currentObserver == null)
                {
                    value = ExcelError.ExcelErrorNA;
                    return true;
                }

                // Subsequent calls get value from Observer
                value = _currentObserver.GetValue();
                return true;
            }
        }

        public void Subscribe(IExcelRtdObserver rtdObserver)
        {
            lock (_lock)
            {
                _currentObserver = rtdObserver;
                _callerState.AddObserver(_currentObserver);
                _currentSubscription = _observable.Subscribe(rtdObserver);
            }
        }

        public void Unsubscribe()
        {
            Debug.Assert(_currentSubscription != null);
            lock (_lock)
            {
                _currentSubscription.Dispose();
                _currentSubscription = null;
                _callerState.RemoveObserver(_currentObserver);
                _currentObserver = null;
            }
        }

        public AsyncCallInfo GetCallInfo()
        {
            return _callInfo;
        }

        public ExcelObservableOptions GetOptions()
        {
            return _options;
        }

        bool IsCompleted()
        {
            lock (_lock)
            {
                if (!_currentObserver.IsCompleted()) return false;
                return _callerState.AreObserversCompleted();
            }
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

    // Helper class to wrap a Task in an Observable - allowing one Subscriber.
    internal class ExcelTaskObservable<TResult> : IExcelObservable
    {
        readonly Task<TResult> _task;
        readonly CancellationTokenSource _cts;

        public ExcelTaskObservable(Task<TResult> task)
        {
            _task = task;
        }

        public ExcelTaskObservable(Task<TResult> task, CancellationTokenSource cts)
            : this(task)
        {
            _cts = cts;
        }

        public IDisposable Subscribe(IExcelObserver observer)
        {
            // Start with a disposable that does nothing
            // Possibly set to a CancellationDisposable later
            IDisposable disp = DefaultDisposable.Instance;

            switch (_task.Status)
            {
                case TaskStatus.RanToCompletion:
                    observer.OnNext(_task.Result);
                    observer.OnCompleted();
                    break;
                case TaskStatus.Faulted:
                    observer.OnError(_task.Exception.InnerException);
                    break;
                case TaskStatus.Canceled:
                    observer.OnError(new TaskCanceledException(_task));
                    break;
                default:
                    var task = _task;
                    // OK - the Task has not completed synchronously
                    // First set up a continuation that will suppress Cancel after the Task completes
                    if (_cts != null)
                    {
                        var cancelDisp = new CancellationDisposable(_cts);
                        task = _task.ContinueWith(t =>
                        {
                            cancelDisp.SuppressCancel();
                            return t;
                        }).Unwrap();

                        // Then this will be the IDisposable we return from Subscribe
                        disp = cancelDisp;
                    }
                    // And handle the Task completion
                    task.ContinueWith(t =>
                    {
                        switch (t.Status)
                        {
                            case TaskStatus.RanToCompletion:
                                observer.OnNext(t.Result);
                                observer.OnCompleted();
                                break;
                            case TaskStatus.Faulted:
                                observer.OnError(t.Exception.InnerException);
                                break;
                            case TaskStatus.Canceled:
                                observer.OnError(new TaskCanceledException(t));
                                break;
                        }
                    });
                    break;
            }

            return disp;
        }

        sealed class DefaultDisposable : IDisposable
        {
            public static readonly DefaultDisposable Instance = new DefaultDisposable();

            // Prevent external instantiation
            DefaultDisposable()
            {
            }

            public void Dispose()
            {
                // no op
            }
        }

        sealed class CancellationDisposable : IDisposable
        {
            bool _suppress;
            readonly CancellationTokenSource _cts;
            public CancellationDisposable(CancellationTokenSource cts)
            {
                if (cts == null)
                {
                    throw new ArgumentNullException("cts");
                }

                _cts = cts;
            }

            public CancellationDisposable()
                : this(new CancellationTokenSource())
            {
            }

            public void SuppressCancel()
            {
                _suppress = true;
            }

            public CancellationToken Token
            {
                get { return _cts.Token; }
            }

            public void Dispose()
            {
                if (!_suppress) _cts.Cancel();
                _cts.Dispose();  // Not really needed...
            }
        }
    }

    internal class ExcelTaskObjectObservable<TResult> : IExcelObservable
    {
        readonly Task<TResult> _task;
        readonly CancellationTokenSource _cts;

        public ExcelTaskObjectObservable(Task<TResult> task)
        {
            _task = task;
        }

        public ExcelTaskObjectObservable(Task<TResult> task, CancellationTokenSource cts)
            : this(task)
        {
            _cts = cts;
        }

        public IDisposable Subscribe(IExcelObserver observer)
        {
            // Start with a disposable that does nothing
            // Possibly set to a CancellationDisposable later
            IDisposable disp = DefaultDisposable.Instance;

            switch (_task.Status)
            {
                case TaskStatus.RanToCompletion:
                    observer.OnNext(_task.Result);
                    //observer.OnCompleted();
                    break;
                case TaskStatus.Faulted:
                    observer.OnError(_task.Exception.InnerException);
                    break;
                case TaskStatus.Canceled:
                    observer.OnError(new TaskCanceledException(_task));
                    break;
                default:
                    var task = _task;
                    // OK - the Task has not completed synchronously
                    // First set up a continuation that will suppress Cancel after the Task completes
                    if (_cts != null)
                    {
                        var cancelDisp = new CancellationDisposable(_cts);
                        task = _task.ContinueWith(t =>
                        {
                            cancelDisp.SuppressCancel();
                            return t;
                        }).Unwrap();

                        // Then this will be the IDisposable we return from Subscribe
                        disp = cancelDisp;
                    }
                    // And handle the Task completion
                    task.ContinueWith(t =>
                    {
                        switch (t.Status)
                        {
                            case TaskStatus.RanToCompletion:
                                observer.OnNext(t.Result);
                                //observer.OnCompleted();
                                break;
                            case TaskStatus.Faulted:
                                observer.OnError(t.Exception.InnerException);
                                break;
                            case TaskStatus.Canceled:
                                observer.OnError(new TaskCanceledException(t));
                                break;
                        }
                    });
                    break;
            }

            return disp;
        }

        sealed class DefaultDisposable : IDisposable
        {
            public static readonly DefaultDisposable Instance = new DefaultDisposable();

            // Prevent external instantiation
            DefaultDisposable()
            {
            }

            public void Dispose()
            {
                // no op
            }
        }

        sealed class CancellationDisposable : IDisposable
        {
            bool _suppress;
            readonly CancellationTokenSource _cts;
            public CancellationDisposable(CancellationTokenSource cts)
            {
                if (cts == null)
                {
                    throw new ArgumentNullException("cts");
                }

                _cts = cts;
            }

            public CancellationDisposable()
                : this(new CancellationTokenSource())
            {
            }

            public void SuppressCancel()
            {
                _suppress = true;
            }

            public CancellationToken Token
            {
                get { return _cts.Token; }
            }

            public void Dispose()
            {
                if (!_suppress) _cts.Cancel();
                _cts.Dispose();  // Not really needed...
            }
        }
    }

    // An IExcelObservable that wraps an IObservable
    internal class ExcelObservable<T> : IExcelObservable
    {
        readonly IObservable<T> _observable;

        public ExcelObservable(IObservable<T> observable)
        {
            _observable = observable;
        }

        public IDisposable Subscribe(IExcelObserver excelObserver)
        {
            var observer = new AnonymousObserver<T>(value => excelObserver.OnNext(value), excelObserver.OnError, excelObserver.OnCompleted);
            return _observable.Subscribe(observer);
        }

        // An IObserver that forwards the inputs to given methods.
        class AnonymousObserver<OT> : IObserver<OT>
        {
            readonly Action<OT> _onNext;
            readonly Action<Exception> _onError;
            readonly Action _onCompleted;

            public AnonymousObserver(Action<OT> onNext, Action<Exception> onError, Action onCompleted)
            {
                if (onNext == null)
                {
                    throw new ArgumentNullException("onNext");
                }
                if (onError == null)
                {
                    throw new ArgumentNullException("onError");
                }
                if (onCompleted == null)
                {
                    throw new ArgumentNullException("onCompleted");
                }
                _onNext = onNext;
                _onError = onError;
                _onCompleted = onCompleted;
            }

            public void OnNext(OT value)
            {
                _onNext(value);
            }

            public void OnError(Exception error)
            {
                _onError(error);
            }

            public void OnCompleted()
            {
                _onCompleted();
            }
        }
    }
}
