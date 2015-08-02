//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using ExcelDna.Logging;

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
            /// Does not normally need to be called if UpdateValue(newValue) has been called,
            /// but can be used to force a recalculation of the RTD cell even if the value has not changed.
            /// </summary>
            public void UpdateNotify()
            {
                _server.SetDirty(this);
            }

            protected internal Topic(ExcelRtdServer server, int topicId)
            {
                _server = server;
                TopicId = topicId;
                _value = ExcelErrorUtil.ToComError(ExcelError.ExcelErrorNA);
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
            protected internal virtual void OnDisconnected()
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
            // By default topic.Value will be #N/A. But if the CreateTopic changed the value, we'll return the new value here.
            // Since newValues is not altered unless this methods is overridden, it means that for fresh topics (where newValues == true) this 
            // return value will be used. That seems right.
            return topic.Value;
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

        // Topic creation - allows derived Topic classes to be used.
        protected virtual Topic CreateTopic(int topicId, IList<string> topicInfo)
        {
            return new Topic(this, topicId);
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

                _callbackObject = callbackObject;
                _updateSync.RegisterUpdateNotify(_callbackObject);
                using (XlCall.Suspend())
                {
                    return ServerStart() ? 1 : 0;
                }
            }
            catch (Exception e)
            {
                Logger.RtdServer.Error("Error in RTD server {0} ServerStart: {1}", GetType().Name, e.ToString());
                return 0;
            }
        }

        // Understanding newValues
        // -----------------------
        // As input: If Excel has a cached value to display, newValues passed in will be false.
        // On return: if newValues is now false, Excel will use cached value if it has one, (else #N/A if passed in as true)
        //            if newValues is now true, Excel will use the value returned by ConnectData.
        object IRtdServer.ConnectData(int topicId, ref Array strings, ref bool newValues)
        {
            try
            {
                // Check for an active topic with the same topicId 
                // - this is unexpected, but is reported as a bug in an earlier Excel version.
                // (Thanks ngm)
                
                // (Does not address the Excel 2010 bug documented here:
                // http://social.msdn.microsoft.com/Forums/en-US/exceldev/thread/ba06ac78-7b64-449b-bce4-9a03ac91f0eb/
                // fixed by hotfix: http://support.microsoft.com/kb/2405840 and SP1. This problem is fixed by the ExcelRtd2010BugHelper. )
                if (_activeTopics.ContainsKey(topicId))
                {
                    using (XlCall.Suspend())
                    {
                        ((IRtdServer)this).DisconnectData(topicId);
                    }
                }

                List<string> topicInfo = new List<string>(strings.Length);
                for (int i = 0; i < strings.Length; i++) topicInfo.Add((string)strings.GetValue(i));

                Topic topic;
                using (XlCall.Suspend())
                {
                    // We create the topic, but what if its value is set here...?
                    topic = CreateTopic(topicId, topicInfo);
                }
                if (topic == null)
                {
                    Logger.RtdServer.Error("Error in RTD server {0} CreateTopic returned null.", GetType().Name);
                    // Not sure what to return here for error. We try the COM error version of #VALUE !?
                    return ExcelErrorUtil.ToComError(ExcelError.ExcelErrorValue);
                }
                _activeTopics[topicId] = topic;
                using (XlCall.Suspend())
                {
                    return ConnectData(topic, topicInfo, ref newValues);
                }
                
            }
            catch (Exception e)
            {
                Logger.RtdServer.Error("Error in RTD server {0} ConnectData: {1}", GetType().Name, e.ToString());
                // Not sure what to return here for error. We try the COM error version of #VALUE !?
                return ExcelErrorUtil.ToComError(ExcelError.ExcelErrorValue);
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

                _activeTopics.Remove(topicId);
                using (XlCall.Suspend())
                {
                    DisconnectData(topic);
                    topic.OnDisconnected();
                }
            }
            catch (Exception e)
            {
                Logger.RtdServer.Error("Error in RTD server {0} DisconnectData: {1}", GetType().Name, e.ToString());
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
                Logger.RtdServer.Error("Error in RTD server {0} Heartbeat: {1}", GetType().Name, e.ToString());
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
                
                using (XlCall.Suspend())
                {
                    ServerTerminate();
                }
            }
            catch (Exception e)
            {
                Logger.RtdServer.Error("Error in RTD server {0} ServerTerminate: {1}", GetType().Name, e.ToString());
            }
        }
    }
}
