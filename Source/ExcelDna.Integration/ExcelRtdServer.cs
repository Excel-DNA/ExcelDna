/*
  Copyright (C) 2005-2013 Govert van Drimmelen

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
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace ExcelDna.Integration.Rtd
{
    // ExcelRtdServer provides a thread-safe, simplified implementation of an RTD server.
    // Uses the SynchronizationWindow, so the AutoOpen checks for derived classes and if any are found it calls SynchronizationManager.Register().

    // Derive from this class for an easy RTD Server
    // CONSIDER: How to support COM Server registration and 'newvalues=false'
    public abstract class ExcelRtdServer : IRtdServer
    {
        public class Topic
        {
            internal readonly int TopicId;
            readonly ExcelRtdServer _server;
            object _value;

            // Setting the Value must be thread-safe!
            // [Obsolete("Rather call ExcelRtdServer.Topic.UpdateValue(value) explicitly.")]
            public object Value
            {
                get { return _value; }
                // Removing set on Topic.Value, must use explicit Topic.UpdateValue(...) call.
                //set
                //{
                //    UpdateValue(value);
                //}
            }

            public ExcelRtdServer Server { get { return _server; } }

            /// <summary>
            /// Sets the topic value and calls UpdateNotify on the RTD Server to refresh.
            /// </summary>
            /// <param name="value"></param>
            public void UpdateValue(object value)
            {
                object fixedValue = FixValue(value);
                lock (_server._updateLock)
                {
                    if (!object.Equals(_value, fixedValue))
                    {
                        _value = fixedValue;
                        UpdateNotify();
                    }
                }
            }

            /// <summary>
            /// Calls UpdateNotify on the RTD server to refresh.
            /// </summary>
            public void UpdateNotify()
            {
                _server.SetDirty(this);
            }

            internal Topic(ExcelRtdServer server, int topicId)
            {
                _server = server;
                TopicId = topicId;
            }

            object FixValue(object value)
            {
                if (value is ExcelError)
                {
                    value = ExcelErrorUtil.ToComError((ExcelError)value);
                }

                // Long strings will cause the topic update to fail horribly.
                // See: http://social.msdn.microsoft.com/Forums/nl-BE/exceldev/thread/436f1aa4-c950-4486-ba58-22a6a12fbf19
                // We truncate long strings.
                string valueString = value as string;
                if (valueString != null && valueString.Length > 255)
                {
                    value = valueString.Substring(0, 255);
                }

                // CONSIDER: Check valid data types
                return value;
            }

            public event EventHandler Disconnected;
            internal void OnDisconnected()
            {
                EventHandler disconnected = Disconnected;
                if (disconnected != null)
                {
                    disconnected(this, EventArgs.Empty);
                }
            }
        }

        internal string RegisteredProgId;

        readonly Dictionary<int, Topic> _activeTopics = new Dictionary<int, Topic>();
        // Using a Dictionary for the dirty topics instead of a HashSet, since we are targeting .NET 2.0
        Dictionary<Topic, object> _dirtyTopics = new Dictionary<Topic, object>();
        bool _notified;
        RtdUpdateSynchronization _updateSync;
        IRTDUpdateEvent _callbackObject;
        readonly object _updateLock = new object();

        // The next few are the core RTD methods to be overridden by implementations
        protected virtual bool ServerStart()
        {
            return true;
        }

        protected virtual void ServerTerminate()
        {
        }

        protected virtual object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            return null;
        }


        protected virtual void DisconnectData(Topic topic)
        {
        }

        // TODO: Threading protection here...?
        protected int HeartbeatInterval
        {
            get { return _callbackObject.HeartbeatInterval; }
            set { _callbackObject.HeartbeatInterval = value; }
        }

        protected virtual int Heartbeat()
        {
            return 1;
        }

        // Add the topic to the dirty set and calls UpdateNotify()
        void SetDirty(Topic topic)
        {
            lock (_updateLock)
            {
                // Check that this topic is still active 
                // (we might be processing the update from another thread, after DisconnectData has been called)
                // and ensure the active topic really is this one.
                Topic activeTopic;
                if (!_activeTopics.TryGetValue(topic.TopicId, out activeTopic) || 
                    !ReferenceEquals(topic, activeTopic))
                {
                    return;
                }
                // Ensure that the topic is in the current dirty list, and call UpdateNotify if needed.
                _dirtyTopics[topic] = null;
                if (!_notified)
                {
                    _updateSync.UpdateNotify(_callbackObject);
                }
                _notified = true;
            }
        }

        // This is the private implementation of the IRtdServer interface
        int IRtdServer.ServerStart(IRTDUpdateEvent callbackObject)
        {
            try
            {
                _updateSync = SynchronizationManager.RtdUpdateSynchronization;
                if (_updateSync == null)
                {
                    // CONSIDER: A better message to alert user of problem here?
                    return 0;
                }

                if (Excel2010RtdBugHelper.ExcelVersionHasRtdBug)
                {
                    Excel2010RtdBugHelper.RecordRtdServerStart();
                }

                _callbackObject = callbackObject;
                _updateSync.RegisterUpdateNotify(_callbackObject);
                using (XlCall.Suspend())
                {
                    return ServerStart() ? 1 : 0;
                }
            }
            catch (Exception e)
            {
                Logging.LogDisplay.WriteLine("Error in RTD server {0} ServerStart: {1}", GetType().Name, e.ToString());
                return 0;
            }
        }

        object IRtdServer.ConnectData(int topicId, ref Array strings, ref bool newValues)
        {
            try
            {
                // Check for an active topic with the same topicId 
                // - this is unexpected, but is reported as a bug in an earlier Excel version.
                // (Thanks ngm)
                
                // (Does not address the Excel 2010 bug documented here:
                // http://social.msdn.microsoft.com/Forums/en-US/exceldev/thread/ba06ac78-7b64-449b-bce4-9a03ac91f0eb/
                // fixed by hotfix: http://support.microsoft.com/kb/2405840 
                // and SP1 )
                if (_activeTopics.ContainsKey(topicId))
                {
                    ((IRtdServer)this).DisconnectData(topicId);
                }

                List<string> topicInfo = new List<string>(strings.Length);
                for (int i = 0; i < strings.Length; i++) topicInfo.Add((string)strings.GetValue(i));

                Topic topic = new Topic(this, topicId);
                _activeTopics[topicId] = topic;

                if (Excel2010RtdBugHelper.ExcelVersionHasRtdBug)
                {
                    Excel2010RtdBugHelper.RecordRtdConnectData(topic, topicInfo);
                }

                using (XlCall.Suspend())
                {
                    return ConnectData(topic, topicInfo, ref newValues);
                }
            }
            catch (Exception e)
            {
                Logging.LogDisplay.WriteLine("Error in RTD server {0} ConnectData: {1}", GetType().Name, e.ToString());
                return null;
            }
        }

        Array IRtdServer.RefreshData(ref int topicCount)
        {
            // Get a copy of the dirty topics to work with, 
            // locking as briefly as possible (thanks Naju).
            Dictionary<Topic, object> temp;
            Dictionary<Topic, object> newDirtyTopics = new Dictionary<Topic, object>();
            lock (_updateLock)
            {
                temp = _dirtyTopics;
                _dirtyTopics = newDirtyTopics;
                _notified = false;
            }

            // The topics in _dirtyTopics may have been Disconnected already.
            // (With another thread updating the value and setting dirty afterwards)
            // We assume Excel doesn't mind being notified of Disconnected topics.
            Dictionary<Topic, object>.KeyCollection dirtyTopics = temp.Keys;
            topicCount = dirtyTopics.Count;
            object[,] result = new object[2, topicCount];
            int i = 0;
            foreach (Topic topic in dirtyTopics)
            {
                result[0, i] = topic.TopicId;
                result[1, i] = topic.Value;
                i++;
            }
            return result;
        }

        void IRtdServer.DisconnectData(int topicId)
        {
            try
            {
                Topic topic;
                if (!_activeTopics.TryGetValue(topicId, out topic))
                {
                    return;
                }

                if (Excel2010RtdBugHelper.ExcelVersionHasRtdBug)
                {
                    Excel2010RtdBugHelper.RecordRtdDisconnectData(topic);
                }

                _activeTopics.Remove(topicId);
                using (XlCall.Suspend())
                {
                    DisconnectData(topic);
                    topic.OnDisconnected();
                }
            }
            catch (Exception e)
            {
                Logging.LogDisplay.WriteLine("Error in RTD server {0} DisconnectData: {1}", GetType().Name, e.ToString());
            }
        }

        int IRtdServer.Heartbeat()
        {
            try
            {
                using (XlCall.Suspend())
                {
                    return Heartbeat();
                }
            }
            catch (Exception e)
            {
                Logging.LogDisplay.WriteLine("Error in RTD server {0} Heartbeat: {1}", GetType().Name, e.ToString());
                return 0;
            }
        }

        void IRtdServer.ServerTerminate()
        {
            try
            {
                // The Unregister call here just tells the reg-free loading that we are gone, 
                // to ensure a fresh load with new 'fake progid' next time.
                // Also safe to call (basically a no-op) if we are not loaded via reg-free, but via real COM Server.
                RtdRegistration.UnregisterRTDServer(RegisteredProgId);

                if (_updateSync != null)
                {
                    _updateSync.DeregisterUpdateNotify(_callbackObject);
                }

                if (Excel2010RtdBugHelper.ExcelVersionHasRtdBug)
                {
                    Excel2010RtdBugHelper.RecordRtdServerTerminate();
                }
                
                using (XlCall.Suspend())
                {
                    ServerTerminate();
                }
            }
            catch (Exception e)
            {
                Logging.LogDisplay.WriteLine("Error in RTD server {0} ServerTerminate: {1}", GetType().Name, e.ToString());
            }
        }
    }

    // We try to work around the Excel 2010 RTM bug whereby DisconnectData is not called on RTD topics 
    // when the calling formula has inputs depending on another cell.
    // E.g. A1: 5, B1: =MyRtdFunc(A1) where MyRtdFunc calls some RTD server with the input as a parameter.
    //      When A1 is changed, MyRtfFunc is called, the new RTD topic created, but DisconnectData never called.

    // CONSIDER: The approach here has one limitation - the ServerTerminate is not emulated when the last topic
    //           disconnects (Excel doesn't terminate the RTD server).
    //           By changing the helper into a wrapper for the ExcelRtdServer, we could do that part properly too.
    class Excel2010RtdBugHelper
    {
        public static bool ExcelVersionHasRtdBug;
        static Excel2010RtdBugHelper()
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
            ExcelVersionHasRtdBug = false;
        }

        #region internal methods to record RTD activity
        public static void RecordRtdConnectData(ExcelRtdServer.Topic topic, IEnumerable<string> topicInfo)
        {
            // Debug.Print("ConnectData " + topic.TopicId);
            string topicKey = TopicKey(topic.Server.RegisteredProgId, topicInfo);
            _activeTopics[topicKey] = topic;
            _activeTopicCallers[topic] = new ExcelReferenceSet();
        }

        public static void RecordRtdDisconnectData(ExcelRtdServer.Topic topic)
        {
            // Debug.Print("DisconnectData " + topic.TopicId);
            foreach (KeyValuePair<string, ExcelRtdServer.Topic> activeTopic in _activeTopics)
            {
                if (activeTopic.Value != topic) continue;

                _activeTopics.Remove(activeTopic.Key);
                foreach (ExcelReference caller in _activeTopicCallers[topic])
                {
                    TopicList callerTopics;
                    if (_activeCallerTopics.TryGetValue(caller, out callerTopics))
                    {
                        callerTopics.Remove(topic);
                        if (callerTopics.Count == 0)
                        {
                            _activeCallerTopics.Remove(caller);
                        }
                    }
                }
                _activeTopicCallers.Remove(topic);
                return;
            }
        }

        // The progId we get here is the one we're loaded as (Maybe RtdSrv.XXX)
        // CONSIDER: How big an issue is duplicate calls?
        //           (Since we process and clear after every calculation, 
        //            and we don't expect many duplicates inside a calculation...?)
        public static void RecordRtdCall(string progId, string[] topicInfo)
        {
            ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            if (caller == null) return; // Can't do much in this case
            _rtdCalls.Add(new RtdCall(caller, progId, topicInfo));
        }

        public static void RecordRtdServerStart()
        {
            InstallAfterCalculateHandler();
        }

        public static void RecordRtdServerTerminate()
        {
            UninstallAfterCalculateHandler();
        }

        // Helper to build a string from the topic info. Using '\t' as the separator - which should be safe
        static string TopicKey(string progId, IEnumerable<string> topicInfo)
        {
            StringBuilder key = new StringBuilder(progId);
            foreach (string ti in topicInfo)
            {
                key.Append("\t");
                key.Append(ti);
            }
            return key.ToString();
        }
        #endregion

        // All the topics that we believe are active
        static readonly Dictionary<string, ExcelRtdServer.Topic> _activeTopics = new Dictionary<string, ExcelRtdServer.Topic>();
        // All the RTD calls (relevant to us) since the last AfterCalculate event
        static readonly List<RtdCall> _rtdCalls = new List<RtdCall>();
        // All the Callers that we've ever seen for each topic.
        static readonly Dictionary<ExcelRtdServer.Topic, ExcelReferenceSet> _activeTopicCallers = new Dictionary<ExcelRtdServer.Topic, ExcelReferenceSet>();
        static readonly Dictionary<ExcelReference, TopicList> _activeCallerTopics = new Dictionary<ExcelReference, TopicList>();

        static void ProcessRtdCalls()
        {
            if (_rtdCalls.Count == 0) return;
            // First, build a dictionary of what we saw,
            // and add any new active callers.
            Dictionary<ExcelReference, TopicList> callerTopicMap = new Dictionary<ExcelReference, TopicList>();
            foreach (RtdCall call in _rtdCalls)
            {
                // find the corresponding RtdTopicInfo
                ExcelRtdServer.Topic topic;
                if (!_activeTopics.TryGetValue(call.TopicKey, out topic))
                {
                    Debug.Print("!!! Unknown Rtd Call: " + call.TopicKey);
                    continue;
                }

                // This is a call to a topic we know
                // Check if we already have an entry for this caller in the callerTopicMap....
                TopicList callerTopics;
                if (!callerTopicMap.TryGetValue(call.Caller, out callerTopics))
                {
                    // ... no - it's a new entry.
                    // Add the caller, and the topic map
                    callerTopics = new TopicList();
                    callerTopicMap[call.Caller] = callerTopics;
                }

                // Get the known callers
                ExcelReferenceSet callers = _activeTopicCallers[topic];
                if (!callers.Contains(call.Caller))
                {
                    // Previously unknown caller for this topic - add to _activeTopicCallers
                    callers.Add(call.Caller);

                    // Add the Topic to the list of topic to watch for this caller
                    TopicList topics;
                    if (!_activeCallerTopics.TryGetValue(call.Caller, out topics))
                    {
                        // Not seen this caller before for this topic - record for future use
                        topics = new TopicList();
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

            // Now figure out what to clean up

            // For each caller and its topics that we saw in this calc ...
            TopicList orphans = new TopicList();
            foreach (KeyValuePair<ExcelReference, TopicList> callerTopics in callerTopicMap)
            {
                ExcelReference thisCalcCaller = callerTopics.Key;
                TopicList thisCalcTopics = callerTopics.Value;

                // ... Check the topics in the _activeCallerTopics list for this caller.
                TopicList activeTopics = _activeCallerTopics[thisCalcCaller];
                TopicList activeTopicsToRemove = null; // Lazy initialize
                foreach (ExcelRtdServer.Topic activeTopic in activeTopics)
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
                        activeTopicsToRemove = new TopicList();
                    activeTopicsToRemove.Add(activeTopic);
                }
                
                if (activeTopicsToRemove != null)
                {
                    foreach (ExcelRtdServer.Topic topicToRemove in activeTopicsToRemove)
                    {
                        activeTopics.Remove(topicToRemove);
                        if (activeTopics.Count == 0)
                        {
                            // Unlikely...?  (due to how the bug works - the caller should have a new topic)
                            _activeCallerTopics.Remove(thisCalcCaller);
                        }
                    }
                }
            }

            // Clear our recording and disconnect the orphans
            _rtdCalls.Clear();
            DisconnectOrphanedTopics(orphans);
        }

        static void DisconnectOrphanedTopics(IEnumerable<ExcelRtdServer.Topic> topics)
        {
            foreach (ExcelRtdServer.Topic topic in topics)
            {
                ((IRtdServer)topic.Server).DisconnectData(topic.TopicId);
            }
        }

        #region AfterCalculate Handler
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
            ProcessRtdCalls();
        }
        #endregion

        // CONSIDER: Might add equality and GetHashCode, so we can keep only unique calls.
        struct RtdCall
        {
            public ExcelReference Caller;
            public string ProgId;
            public string[] TopicInfo;

            public RtdCall(ExcelReference caller, string progId, string[] topicInfo)
            {
                Caller = caller;
                ProgId = progId;
                TopicInfo = topicInfo;
            }

            public string TopicKey
            {
                get
                {
                    return Excel2010RtdBugHelper.TopicKey(ProgId, TopicInfo);
                }
            }
        }

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
        class TopicList : List<ExcelRtdServer.Topic> { }

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
