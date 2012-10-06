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
using System.Runtime.InteropServices;
using System.Threading;

namespace ExcelDna.Integration.Rtd
{
    internal static class AsyncObservableImpl
    {
        static readonly Dictionary<AsyncCallInfo, Guid> _asyncCallIds = new Dictionary<AsyncCallInfo, Guid>();
        static readonly Dictionary<Guid, AsyncObservableState> _observableStates = new Dictionary<Guid, AsyncObservableState>();

        // This is the most general RTD registration
        // TODO: This should not be called from a ThreadSafe function. Check...?
        public static object ProcessObservable(string functionName, object parameters, ExcelObservableSource getObservable)
        {
            // TODO: Check here that registration has happened.
            // CONSIDER: Why not same problems with all RTD servers?

            AsyncCallInfo callInfo = new AsyncCallInfo(functionName, parameters);

            // Shortcut if already registered
            object value;
            if (GetValueIfRegistered(callInfo, out value))
            {
                return value;
            }

            // Actually register as a new Observable
            IExcelObservable observable = getObservable();
            return RegisterObservable(callInfo, observable);
        }

        // Make a one-shot 'Observable' from the func
        public static object ProcessFunc(string functionName, object parameters, ExcelFunc func)
        {
            return ProcessObservable(functionName, parameters,
                delegate { return new ThreadPoolDelegateObservable(func); });
        }

        static bool GetValueIfRegistered(AsyncCallInfo callInfo, out object value)
        {
            Guid id;
            if (_asyncCallIds.TryGetValue(callInfo, out id))
            {
                // Already registered.
                AsyncObservableState state = _observableStates[id];
                value = state.GetValue();
                return true;
            }
            value = null;
            return false;
        }

        // Register a new observable
        static object RegisterObservable(AsyncCallInfo callInfo, IExcelObservable observable)
        {
            // Check it's not registered already
            Debug.Assert(!_asyncCallIds.ContainsKey(callInfo));

            // Set up a new Id and ObservableState and keep track of things
            Guid id = Guid.NewGuid();
            _asyncCallIds[callInfo] = id;
            AsyncObservableState state = new AsyncObservableState(id, callInfo, observable);
            _observableStates[id] = state;

            // Will spin up RTD server and topic if required, causing us to be called again...
            return state.GetValue();
        }

        internal static void ConnectObserver(Guid id, ExcelRtdObserver rtdObserver)
        {
            // TODO: Checking...(huh?)
            AsyncObservableState state = _observableStates[id];
            // Start the work for this AsyncCallInfo, and subscribe the topic to the result
            state.Subscribe(rtdObserver);
        }

        internal static void DisconnectObserver(Guid id)
        {
            AsyncObservableState state = _observableStates[id];
            state.Unsubscribe();
            _observableStates.Remove(id);
            _asyncCallIds.Remove(state.GetCallInfo());
        }
    }

    // Pushes data to Excel via the RTD topic.
    internal class ExcelRtdObserver : IExcelObserver
    {
        readonly ExcelRtdServer.Topic _topic;

        // Indicates whether the RTD Topic should be shut down
        // Set to true if the work is completed or if an error is signalled.
        public bool IsCompleted { get; private set; }
        public object Value { get; private set; } // Keeping our own value, since the RTD Topic.Value is insufficient (e.g. string length limitation)

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
            _topic.UpdateNotify();
        }

        public void OnError(Exception exception)
        {
            // TODO: Is the sequence here important?
            Value = ExcelIntegration.HandleUnhandledException(exception);
            // Set the topic value to #VALUE (not used!?) - converted to COM code in Topic
            _topic.UpdateValue(ExcelError.ExcelErrorValue);
            IsCompleted = true;
        }

        public void OnNext(object value)
        {
            Value = value;
            // Not actually setting the topic value, just poking it
            // TODO: Using the 'fake' RTD value should be optional - to give us a way to deal with 'newValues' one day.
            _topic.UpdateValue(DateTime.UtcNow.ToOADate());
        }
    }

    [ComVisible(true)]
    internal class ExcelObserverRtdServer : ExcelRtdServer
    {
        Dictionary<Topic, Guid> _topicGuids = new Dictionary<Topic, Guid>();

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            // Retrieve and store the GUID from the topic's first info string - used to hook up to the Async state
            Guid id = new Guid(topicInfo[0]);
            _topicGuids[topic] = id;

            // Create a new ExcelRtdObserver, for the Topic, which will listen to the Observable
            // (Internally also set initial value - #N/A for now)
            ExcelRtdObserver rtdObserver = new ExcelRtdObserver(topic);
            // ... and subscribe it
            AsyncObservableImpl.ConnectObserver(id, rtdObserver);

            // Return something: #N/A for now. Not currently used.
            // TODO: Allow customize?
            return ExcelErrorUtil.ToComError(ExcelError.ExcelErrorNA);
        }

        protected override void DisconnectData(Topic topic)
        {
            // Retrieve the GUID from the topic's first info string - used to hook up to the Async state
            Guid id = _topicGuids[topic];

            // ... and unsubscribe it
            AsyncObservableImpl.DisconnectObserver(id);
            _topicGuids.Remove(topic);
        }

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
    internal struct AsyncCallInfo
    {
        readonly string _functionName;
        readonly object _parameters;
        readonly int    _hashCode;

        public AsyncCallInfo(string functionName, object parameters)
        {
            _functionName = functionName;
            _parameters = parameters;
            _hashCode = 0; // Need to set to smoe value before we call a method.
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
            //           to other data types. For now this allow everything that can be passed as parameters from Excel-DNA.
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
                obj is decimal)
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
                        hash = hash * 23 +  item.GetHashCode(); 
                    }
                    return hash;
                }

                double[,] doubles2 = obj as double[,];
                if (doubles2 != null) 
                {
                    foreach (double item in doubles2)
                    {
                        hash = hash * 23 +  item.GetHashCode(); 
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
            }
            throw new ArgumentException("Invalid type used for async parameter(s)", "parameters");
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (obj.GetType() != typeof(AsyncCallInfo)) return false;
            return Equals((AsyncCallInfo)obj);
        }

        bool Equals(AsyncCallInfo other)
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
        #endregion

        public override int GetHashCode()
        {
            return _hashCode;
        }
    }

    internal class AsyncObservableState
    {
        readonly Guid _id;
        readonly AsyncCallInfo _callInfo; // Bit ugly having this here - need a bi-directional dictionary...
        readonly IExcelObservable _observable;
        ExcelRtdObserver _currentObserver;
        IDisposable _currentSubscription;

        public AsyncObservableState(Guid id, AsyncCallInfo callInfo, IExcelObservable observable)
        {
            _id = id;
            _callInfo = callInfo;
            _observable = observable;
        }

        public object GetValue()
        {
            if (_currentObserver == null || !_currentObserver.IsCompleted)
            {
                // NOTE: At this post the SynchronizationManager must be registered!
                if (!SynchronizationManager.IsInstalled)
                {
                    Debug.Print("SynchronizationManager not registered!");
                    throw new InvalidOperationException("SynchronizationManager must be registered for async and observable support");
                }
                
                // Ensure that Excel-DNA knows about the RTD Server, since it would not have been registered when loading
                ExcelObserverRtdServer.EnsureRtdServerRegistered();

                // Refresh RTD call
                // NOTE: First time this will result in a call to ConnectData, which will call Subscribe and set the _currentObserver
                object unused = XlCall.RTD("ExcelDna.Integration.Rtd.ExcelObserverRtdServer", null, _id.ToString());
            }

            // No assumptions about previous state here - could have re-entered this class.

            // TODO: Allow customize this value?
            //       Not too serious since the user can remap in the UDF.
            if (_currentObserver == null) return ExcelError.ExcelErrorNA;

            // Subsequent calls get value from Observer
            return _currentObserver.Value;
        }

        public void Subscribe(ExcelRtdObserver rtdObserver)
        {
            _currentObserver = rtdObserver;
            _currentSubscription = _observable.Subscribe(rtdObserver);
        }

        public void Unsubscribe()
        {
            Debug.Assert(_currentSubscription != null);
            _currentSubscription.Dispose();
            _currentSubscription = null;
            _currentObserver = null;
        }

        public AsyncCallInfo GetCallInfo()
        {
            return _callInfo;
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
                    // TODO: Log somehow
                    observer.OnError(ex);
                }
            });

            return new DummyDisposable();
        }

        class DummyDisposable : IDisposable
        {
            public void Dispose() { }
        }
    }
}