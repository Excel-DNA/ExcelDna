//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using ExcelDna.Integration.Rtd;

namespace ExcelDna.Integration
{
    // We try to work around the Excel 2010 RTM bug whereby DisconnectData is not called on RTD topics 
    // when the calling formula has inputs depending on another cell.
    // E.g. A1: 5, B1: =MyRtdFunc(A1) where MyRtdFunc calls some RTD server with the input as a parameter.
    //      When A1 is changed, MyRtfFunc is called, the new RTD topic created, but DisconnectData never called.
    // The bug is discussed in this therad: http://social.msdn.microsoft.com/Forums/en-US/exceldev/thread/ba06ac78-7b64-449b-bce4-9a03ac91f0eb/

    // TODO: If a formula is moved from one cell to another, or some rows / cols are deleted, the UDFs will not be called automatically again.
    //       Consider whether this might be a problem for us here.
    class ExcelRtd2010BugHelper : IRtdServer
    {
        public static readonly bool ExcelVersionHasRtdBug;
        static ExcelRtd2010BugHelper()
        {
            // RTM was:                                     14.0.4760.1000
            // the Excel version with the hotfix* is:       14.0.5128.5000
            // SP1 where the bug is also fixed has version: 14.0.6129.5000
            // * according to http://support.microsoft.com/kb/2405840
            FileVersionInfo ver = ExcelDnaUtil.ExcelExecutableInfo;
            if (ver.FileMajorPart == 14 && ver.FileMinorPart == 0 && ver.FileBuildPart < 5128)
            {
                ExcelVersionHasRtdBug = true;
                return;
            }
//#if DEBUG
//            ExcelVersionHasRtdBug = true;   // To test a bit...
//#else
            ExcelVersionHasRtdBug = false;
//#endif
        }

        static Dictionary<string, ExcelRtd2010BugHelper> _wrappers;
        readonly string _wrappedServerRegisteredProgId;
        readonly Type _wrappedServerType;
        IRtdServer _wrappedServer;
        IRTDUpdateEvent _callbackObject;
        public ExcelRtd2010BugHelper(string wrappedServerRegisteredProgId, Type wrappedServerType)
        {
            // Since we don't provide any wrapping, we should only accept ExcelRtdServer subclasses.
            Debug.Assert(wrappedServerType.IsSubclassOf(typeof(ExcelRtdServer)));

            _wrappedServerRegisteredProgId = wrappedServerRegisteredProgId;
            _wrappedServerType = wrappedServerType;
        }

        #region ExcelRtdServer overrides
        public int ServerStart(IRTDUpdateEvent CallbackObject)
        {
            Debug.Print("### ServerStart " + _wrappedServerRegisteredProgId);

            if (_wrappers == null)
            {
                _wrappers = new Dictionary<string, ExcelRtd2010BugHelper>();
            }
            Debug.Assert(!_wrappers.ContainsKey(_wrappedServerRegisteredProgId));
            _wrappers[_wrappedServerRegisteredProgId] = this;

            // ServerStart should never be called when ServerTerminate has not been called.
            // So we should never have a wrapped server at this point.
            Debug.Assert(_wrappedServer == null);
            InstallAfterCalculateHandler();
            CreateWrappedServer();
            _callbackObject = CallbackObject;
            return _wrappedServer.ServerStart(_callbackObject);
        }

        public void ServerTerminate()
        {
            Debug.Print("### ServerTerminate " + _wrappedServerRegisteredProgId);

            // We might already have terminated the wrapped server ....
            if (_wrappedServer != null)
            {
                // TODO: This assertion is not quite right
                //       If Excel is shutting down, the RTD server will receive a ServerTerminate, 
                //       even though there weren't DisconnectData calls for each Topic.
                //       (Is there a way to know that we're shutting down, so that we can 
                //        make the check only when required?)
                Debug.Assert(_activeTopics.Count == 0);
                _wrappedServer.ServerTerminate();
                _wrappedServer = null;
            }
            // Now remove this wrapper from the list
            _wrappers.Remove(_wrappedServerRegisteredProgId);
            UninstallAfterCalculateHandler();
        }

        public object ConnectData(int topicId, ref Array strings, ref bool newValues)
        {
            Debug.Print("ConnectData " + _wrappedServerRegisteredProgId + ":" + topicId);
            if (_wrappedServer == null)
            {
                // We have to create and start a new server instance
                CreateWrappedServer();
                int startResult = _wrappedServer.ServerStart(_callbackObject);
                if (startResult == 0)
                {
                    // TODO: Deal with error here...
                    _wrappedServer.ServerTerminate();
                    _wrappedServer = null;
                    return ExcelErrorUtil.ToComError(ExcelError.ExcelErrorValue);
                }
            }

            string topicKey = GetTopicKey(strings);
            _activeTopics[topicKey] = topicId;
            _activeTopicCallers[topicId] = new ExcelReferenceSet();
            // Everything is fine
            return _wrappedServer.ConnectData(topicId, ref strings, ref newValues);
        }

        public void DisconnectData(int topicId)
        {
            // CONSIDER: But what if we previously ServerTerminated, and only now Excel is figuring out it must DisconnectData?
            Debug.Assert(_wrappedServer != null);

            // This is called both by the DisconnectOrphans handler and Excel regularly
            foreach (KeyValuePair<string, int> activeTopic in _activeTopics)
            {
                if (activeTopic.Value != topicId) continue;

                // Only call the wrapped Disconnect if the topic is still active
                _wrappedServer.DisconnectData(topicId);
                Debug.Print("DisconnectData " + _wrappedServerRegisteredProgId + ":" + topicId);
                _activeTopics.Remove(activeTopic.Key);
                foreach (ExcelReference caller in _activeTopicCallers[topicId])
                {
                    TopicIdList callerTopics;
                    if (_activeCallerTopics.TryGetValue(caller, out callerTopics))
                    {
                        callerTopics.Remove(topicId);
                        if (callerTopics.Count == 0)
                        {
                            _activeCallerTopics.Remove(caller);
                        }
                    }
                }
                _activeTopicCallers.Remove(topicId);
                return;
            }

            // We might have to let the server instance go.
            // (Even though Excel thinks there are some active topics, there aren't really.)
            if (_activeTopics.Count == 0)
            {
                _wrappedServer.ServerTerminate();
                // CONSIDER: Is this safe ...? What if Excel tried to Connect again later...?
                _wrappedServer = null;
            }
        }

        public Array RefreshData(ref int topicCount)
        {
            return _wrappedServer.RefreshData(ref topicCount);
        }

        public int Heartbeat()
        {
            return _wrappedServer.Heartbeat();
        }

        #endregion

        void CreateWrappedServer()
        {
            _wrappedServer = (IRtdServer)Activator.CreateInstance(_wrappedServerType);
            // Would be nice to combine the initialization here and in ExcelRtd.Rtd ...
            ExcelRtdServer rtdServer = _wrappedServer as ExcelRtdServer;
            if (rtdServer != null)
            {
                rtdServer.RegisteredProgId = _wrappedServerRegisteredProgId;
            }
        }

        // Per-server state and processing of RTD calls

        readonly List<RtdCall> _rtdCalls = new List<RtdCall>(); // Stores every good call to xlfRtd
        readonly List<RtdCall> _rtdCompletes = new List<RtdCall>(); // Stores the special completion non-calls
        // All the topics that we believe are active for this server (map from topic key to topic id)
        readonly Dictionary<string, int> _activeTopics = new Dictionary<string, int>();
        // All the RTD calls (relevant to us) since the last AfterCalculate event
        // All the Callers that we've ever seen for each topic.
        readonly Dictionary<int, ExcelReferenceSet> _activeTopicCallers = new Dictionary<int, ExcelReferenceSet>();
        readonly Dictionary<ExcelReference, TopicIdList> _activeCallerTopics = new Dictionary<ExcelReference, TopicIdList>();

        // The progId we get here is the one we're internally registered with (NOT the RtdSrv.XXX)
        // CONSIDER: How big an issue is duplicate calls?
        //           (Since we process and clear after every calculation, 
        //            and we don't expect many duplicates inside a calculation...?)
        static public void RecordRtdCall(string progId, string[] topicInfo)
        {
            ExcelRtd2010BugHelper wrapper;
            if (_wrappers.TryGetValue(progId, out wrapper))
            {
                ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
                if (caller == null) return; // Can't do much in this case
                wrapper._rtdCalls.Add(new RtdCall(caller, topicInfo));
            }
        }

        // New part added for Excel 2010 not disconnecting after completion
        // (when xlfRtd is _not_ called from the cell, and normally Excel would take that as the queue to disconnect)
        static public void RecordRtdComplete(string progId, params string[] topicInfo)
        {
            ExcelRtd2010BugHelper wrapper;
            if (_wrappers.TryGetValue(progId, out wrapper))
            {
                ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
                if (caller == null) return; // Can't do much in this case
                wrapper._rtdCompletes.Add(new RtdCall(caller, topicInfo));
            }
        }

        void ProcessRtdCalls()
        {
            if (_rtdCalls.Count == 0 && _rtdCompletes.Count == 0) return;
            // First, build a dictionary of what we saw,
            // and add any new active callers.
            Dictionary<ExcelReference, TopicIdList> callerTopicMap = new Dictionary<ExcelReference, TopicIdList>();
            foreach (RtdCall call in _rtdCalls)
            {
                // find the corresponding RtdTopicInfo
                int topic;
                if (!_activeTopics.TryGetValue(call.TopicKey, out topic))
                {
                    Debug.Print("!!! Unknown Rtd Call: " + call.TopicKey);
                    continue;
                }

                // This is a call to a topic we know
                // Check if we already have an entry for this caller in the callerTopicMap....
                TopicIdList callerTopics;
                if (!callerTopicMap.TryGetValue(call.Caller, out callerTopics))
                {
                    // ... no - it's a new entry.
                    // Add the caller, and the topic map
                    callerTopics = new TopicIdList();
                    callerTopicMap[call.Caller] = callerTopics;
                }

                // Get the known callers
                ExcelReferenceSet callers = _activeTopicCallers[topic];
                if (!callers.Contains(call.Caller))
                {
                    // Previously unknown caller for this topic - add to _activeTopicCallers
                    callers.Add(call.Caller);

                    // Add the Topic to the list of topic to watch for this caller
                    TopicIdList topics;
                    if (!_activeCallerTopics.TryGetValue(call.Caller, out topics))
                    {
                        // Not seen this caller before for this topic - record for future use
                        topics = new TopicIdList();
                        _activeCallerTopics[call.Caller] = topics;
                    }
                    // This is a caller we've dealt with before
                    // This topic should not be in the list
                    // TODO: What if it is called twice from a single formula...?
                    Debug.Assert(!topics.Contains(topic));
                    topics.Add(topic);

                    // NOTE: topics might include the orphans!
                }
                // One of the known callers
                // Anyway - record that we saw it in this calc
                callerTopics.Add(topic);
            }

            // Process calls that are 'complete' - we can record that we didn't see (i.e. call xlfRtd for) the topic in this call
            foreach (var call in _rtdCompletes)
            {
                // These callers were called, but not with the topics we expected (since there was no real RTD call)
                // find the corresponding RtdTopicInfo
                int topic;
                if (!_activeTopics.TryGetValue(call.TopicKey, out topic))
                {
                    Debug.Fail("!!! Unknown Rtd Call: " + call.TopicKey);
                    continue;
                }

                // This is a call to a topic we know
                // Check if we already have an entry for this caller in the callerTopicMap....
                TopicIdList callerTopics;
                if (!callerTopicMap.TryGetValue(call.Caller, out callerTopics))
                {
                    // This caller has no topics (in this calculation)
                    // Note that it was called (but we'll add no topics...)
                    // Otherwise it's fine - we've listed this as a caller to examine, but we won't put this topic in
                    callerTopics = new TopicIdList();
                    callerTopicMap[call.Caller] = callerTopics;
                }

                if (callerTopics.Contains(topic))
                {
                    Debug.Fail("!!! Inconsistent Rtd Call (RtdCalls contains the RtdComplete call): " + call.TopicKey);
                }
            }


            // Now figure out what to clean up

            // For each caller and its topics that we saw in this calc ...
            TopicIdList orphans = new TopicIdList();
            foreach (KeyValuePair<ExcelReference, TopicIdList> callerTopics in callerTopicMap)
            {
                ExcelReference thisCalcCaller = callerTopics.Key;
                TopicIdList thisCalcTopics = callerTopics.Value;

                // ... Check the topics in the _activeCallerTopics list for this caller.
                TopicIdList activeTopics = _activeCallerTopics[thisCalcCaller];
                TopicIdList activeTopicsToRemove = null; // Lazy initialize
                foreach (int activeTopic in activeTopics)
                {
                    // If we've seen the topic in this calc, all is fine.
                    if (thisCalcTopics.Contains(activeTopic)) continue;

                    // ... Any topic not seen in this calc might be an orphan (so check if it has other callers).
                    // ... ensure that the active topic also does not have the caller in its activeCallers list any more.
                    ExcelReferenceSet activeCallers = _activeTopicCallers[activeTopic];
                    if (activeCallers.Remove(thisCalcCaller))
                    {
                        // - now check if this topic is an orphan
                        if (activeCallers.Count == 0)
                        {
                            orphans.Add(activeTopic);
                        }
                    }

                    // The activeTopic was one of the topics for thisCalcCaller, but is no longer.
                    // Should now be removed from the list of topics for this caller.
                    if (activeTopicsToRemove == null)
                        activeTopicsToRemove = new TopicIdList();
                    activeTopicsToRemove.Add(activeTopic);
                }

                if (activeTopicsToRemove != null)
                {
                    foreach (int topicToRemove in activeTopicsToRemove)
                    {
                        activeTopics.Remove(topicToRemove);
                        if (activeTopics.Count == 0)
                        {
                            // This happens if we have completed the topic, and Excel might not disconnect
                            // but the topic is no longer active.
                            // I think Excel will try to Disconnect later, e.g. if we delete the formula or something.
                            _activeCallerTopics.Remove(thisCalcCaller);
                        }
                    }
                }
            }

            // Clear our recording and disconnect the orphans
            _rtdCalls.Clear();
            DisconnectOrphanedTopics(orphans);
        }

        void DisconnectOrphanedTopics(IEnumerable<int> topicIds)
        {
            foreach (int topicId in topicIds)
            {
                DisconnectData(topicId);
            }
        }

        static string GetTopicKey(string[] topicInfo)
        {
            return string.Join("\t", topicInfo);
        }

        static string GetTopicKey(Array topicInfo)
        {
            Debug.Assert(topicInfo.Length > 0);
            StringBuilder key = new StringBuilder((string)topicInfo.GetValue(0));
            for (int i = 1; i < topicInfo.Length; i++)
            {
                key.Append("\t");
                key.Append((string)topicInfo.GetValue(i));
            }
            return key.ToString();
        }

        #region AfterCalculate Handler

        // We could try the Excel 2010 CalculationEnded event, but there is some Access Violation
        // even though the registration seems OK.
        // Anyway, mixing the C API calls and the COM stuff seems a bit risky, and we want to be able to make COM calls to the RTD server.
        // Implementing in terms of Application.AfterCalculate is not pretty but seems to work OK.

        delegate void AfterCalculateEventHandler();
        static AppEventSink _appEventSink;
        static IConnectionPoint _connectionPoint;
        static int _adviseCookie;
        static int _activeServerCount;

        static void InstallAfterCalculateHandler()
        {
            if (_appEventSink == null)
            {
                object app = ExcelDnaUtil.Application;
                _appEventSink = new AppEventSink();
                IConnectionPointContainer connectionPointContainer = (IConnectionPointContainer)app;
                Guid appEventsInterfaceId = new Guid("00024413-0000-0000-c000-000000000046");
                connectionPointContainer.FindConnectionPoint(ref appEventsInterfaceId, out _connectionPoint);
                _connectionPoint.Advise(_appEventSink, out _adviseCookie);

                _appEventSink.AfterCalculate = AfterCalculate;
            }
            _activeServerCount++;
        }

        static void UninstallAfterCalculateHandler()
        {
            _activeServerCount--;
            if (_activeServerCount <= 0 && _appEventSink != null)
            {
                _connectionPoint.Unadvise(_adviseCookie);
                Marshal.ReleaseComObject(_connectionPoint);
                _connectionPoint = null;
                _appEventSink = null;
            }
        }

        static void AfterCalculate()
        {
            foreach (ExcelRtd2010BugHelper helper in _wrappers.Values)
                helper.ProcessRtdCalls();
        }
        #endregion

        // CONSIDER: Might add equality and GetHashCode, so we can keep only unique calls.
        struct RtdCall
        {
            public ExcelReference Caller;
            public string[] TopicInfo;

            public RtdCall(ExcelReference caller, string[] topicInfo)
            {
                Caller = caller;
                TopicInfo = topicInfo;
            }

            public string TopicKey
            {
                get
                {
                    return GetTopicKey(TopicInfo);
                }
            }
        }

        // We need something like HashSet with fast contains lookup
        // For .NET 2.0, Dictionary is a good replacement.
        class SimpleSet<T> : IEnumerable<T>
        {
            readonly Dictionary<T, object> _impl = new Dictionary<T, object>();

            public bool Add(T element)
            {
                if (_impl.ContainsKey(element))
                    return false;

                _impl.Add(element, null);
                return true;
            }

            public bool Remove(T element)
            {
                return _impl.Remove(element);
            }

            public bool Contains(T element)
            {
                return _impl.ContainsKey(element);
            }

            public int Count { get { return _impl.Count; } }

            public IEnumerator<T> GetEnumerator()
            {
                return _impl.Keys.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }

        class ExcelReferenceSet : SimpleSet<ExcelReference> { }
        class TopicIdList : List<int> { }

        // To provide a private implementation of the AppEventSink, we implement it in terms of IReflect.
        // The .NET handling for IDispatch will call into IReflect (even for private types).
        // (This might have been nice as an F# object expression.)
        class AppEventSink : IReflect
        {
            public AfterCalculateEventHandler AfterCalculate;

            object IReflect.InvokeMember(string name, BindingFlags invokeAttr, Binder binder, object target, object[] args, ParameterModifier[] modifiers, CultureInfo culture, string[] namedParameters)
            {
                // We're not called with the name, since it was not provided via a MethodInfo.
                // So we get the DISPID identifier (the magic number comes from the Excel type library).
                if (name == "[DISPID=2612]") AfterCalculate();
                return null;
            }

            // All the rest is unimportant
            FieldInfo IReflect.GetField(string name, BindingFlags bindingAttr)
            {
                throw new NotImplementedException();
            }

            FieldInfo[] IReflect.GetFields(BindingFlags bindingAttr)
            {
                return null;
            }

            MemberInfo[] IReflect.GetMember(string name, BindingFlags bindingAttr)
            {
                throw new NotImplementedException();
            }

            MemberInfo[] IReflect.GetMembers(BindingFlags bindingAttr)
            {
                throw new NotImplementedException();
            }

            MethodInfo IReflect.GetMethod(string name, BindingFlags bindingAttr)
            {
                throw new NotImplementedException();
            }

            MethodInfo IReflect.GetMethod(string name, BindingFlags bindingAttr, Binder binder, Type[] types, ParameterModifier[] modifiers)
            {
                throw new NotImplementedException();
            }

            MethodInfo[] IReflect.GetMethods(BindingFlags bindingAttr)
            {
                // Not so useful here... we are not public so the MethodInfos returned here will not be used,
                // and so we will be called via InvokeMember anyway.
                return null;
            }

            PropertyInfo[] IReflect.GetProperties(BindingFlags bindingAttr)
            {
                return null;
            }

            PropertyInfo IReflect.GetProperty(string name, BindingFlags bindingAttr, Binder binder, Type returnType, Type[] types, ParameterModifier[] modifiers)
            {
                throw new NotImplementedException();
            }

            PropertyInfo IReflect.GetProperty(string name, BindingFlags bindingAttr)
            {
                throw new NotImplementedException();
            }

            Type IReflect.UnderlyingSystemType
            {
                get { throw new NotImplementedException(); }
            }
        }
    }
}
