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
        // TODO: Different name RunObservable?
        public static object ProcessObservable(string functionName, string parametersToken, bool useCaller, ExcelObservableFunc observableFunc)
        {
            AsyncCallInfo callInfo = new AsyncCallInfo(functionName, useCaller ? XlCall.Excel(XlCall.xlfCaller) : null, parametersToken);

            // Shortcut if already registered
            object value;
            if (GetValueIfRegistered(callInfo, out value))
            {
                return value;
            }

            // Actually register as a new Observable
            return RegisterObservable(callInfo, observableFunc());
        }

        public static object ProcessFunc(string functionName, string parametersToken, bool useCaller, ExcelFunc func)
        {
            AsyncCallInfo callInfo = new AsyncCallInfo(functionName, useCaller ? XlCall.Excel(XlCall.xlfCaller) : null, parametersToken);
            object value;
            if (GetValueIfRegistered(callInfo, out value))
            {
                return value;
            }

            // Not registered - ???
            ThreadPoolDelegateObservable obs = new ThreadPoolDelegateObservable(func);
            return RegisterObservable(callInfo, obs);
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
            // TODO: Checking
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
            Value = Integration.HandleUnhandledException(exception);
            // Set the topic value to #VALUE (not used!?) - converted to COM code in Topic
            _topic.UpdateValue(ExcelError.ExcelErrorValue);
            IsCompleted = true;
        }

        public void OnNext(object value)
        {
            Value = value;
            // Not actually setting the topic value, just poking it
            // TODO: This must be an option - to give us a way to deal with 'newValue' one day.
            _topic.UpdateValue(DateTime.Now.ToOADate());
        }
    }


    // TODO: How to register!?


    // TODO: Whether the caller is part of the info used for checking should be configurable
    // TODO: Need a helper to check for equality even with array parameters

    [ComVisible(true)]
    internal class ExcelObserverRtdServer : ExcelRtdServer
    {
        protected override object ConnectData(Topic topic, ref bool newValues)
        {
            // Retrieve the GUID from the topic's first info string - used to hook up to the Async state
            Guid id = new Guid(topic.TopicInfo[0]);

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
            Guid id = new Guid(topic.TopicInfo[0]);

            // ... and unsubscribe it
            AsyncObservableImpl.DisconnectObserver(id);
        }

        // This makes sure the hook up with the registration-free RTD loading is in place.
        // For a user RTD server the add-in loading would ensure this, but not for this class since it is declared inside Excel-DNA.
        // CONSDIDER: How do we deal with saved RTD values to make a per-addin unique persistent Progid?
        static bool _isRegistered = false;
        internal static void EnsureRtdServerRegistered()
        {
            if (!_isRegistered)
            {
                RtdRegistration.RegisterRtdServerTypes(new Type[] { typeof(ExcelDna.Integration.Rtd.ExcelObserverRtdServer) });
            }
            _isRegistered = true;
        }

    }

    // Used as Keys in a Dictionary - should be immutable.
    internal struct AsyncCallInfo
    {
        readonly string FunctionName;
        readonly object Caller;
        readonly string ParametersToken;

        public AsyncCallInfo(string functionName, object caller, string parametersToken)
        {
            FunctionName = functionName;
            Caller = caller;
            ParametersToken = parametersToken;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (obj.GetType() != typeof(AsyncCallInfo)) return false;
            return Equals((AsyncCallInfo)obj);
        }

        bool Equals(AsyncCallInfo other)
        {
            bool nameEqual = Equals(other.FunctionName, FunctionName);
            bool callerEqual = Equals(other.Caller, Caller);
            bool paramEqual = Equals(other.ParametersToken, ParametersToken);
            return nameEqual && callerEqual && paramEqual;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int result = (FunctionName != null ? FunctionName.GetHashCode() : 0);
                result = (result * 397) ^ (Caller != null ? Caller.GetHashCode() : 0);
                result = (result * 397) ^ (ParametersToken != null ? ParametersToken.GetHashCode() : 0);
                return result;
            }
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
                // Ensure that Excel-DNA knows about the RTD Server, since it would not have been registered when loading
                ExcelObserverRtdServer.EnsureRtdServerRegistered();

                // NOTE: At this post the SynchronizationManager must be registered!
                if (!SynchronizationManager.IsRegistered)
                {
                    Debug.Print("SynchronizationManager not registered!");
                    throw new InvalidOperationException("SynchronizationManager must be registered for async and observable support");
                }
                
                // Refresh RTD call
                // NOTE: First time this will result in a call to ConnectData, which will call Subscribe and set the _currentObserver
                object unused = XlCall.RTD("ExcelDna.Integration.Rtd.ExcelObserverRtdServer", null, _id.ToString());
            }

            // No assumptions about previous state here - could have re-entered this class.

            // TODO: Allow customize this value?
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
        ExcelFunc _func;
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

