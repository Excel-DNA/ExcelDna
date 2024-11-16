//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

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
            readonly ExcelRtdServer _server;
            readonly int _topicId;
            object _value;

            public ExcelRtdServer Server { get { return _server; } }
            public int TopicId { get { return _topicId; } }

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

            /// <summary>
            /// Sets the topic value and calls UpdateNotify on the RTD Server to refresh.
            /// </summary>
            /// <param name="value"></param>
            public void UpdateValue(object value)
            {
                object fixedValue = FixValue(value);
                lock (Server._updateLock)
                {
                    if (!object.Equals(_value, fixedValue))
                    {
                        _value = fixedValue;
                        _server.SetDirtyInsideLock(this);
                    }
                }
            }

            /// <summary>
            /// Calls UpdateNotify on the RTD server to refresh.
            /// Does not normally need to be called if UpdateValue(newValue) has been called,
            /// but can be used to force a recalculation of the RTD cell even if the value has not changed.
            /// NOTE: It seems around Feb 2020 this stopped working, and Excel no longer updates a cell just because RefreshData contains the topic
            ///       If this is true, we should obsolete this as a public member, since it no longer implements the intended meaning
            /// </summary>
            [Obsolete("Due to recent Excel updates, can no longer cause Topic update without changing value")]
            public void UpdateNotify()
            {
                lock (_server._updateLock)
                {
                    _server.SetDirtyInsideLock(this);
                }
            }

            protected internal Topic(ExcelRtdServer server, int topicId)
            {
                _server = server;
                _topicId = topicId;
                _value = ExcelErrorUtil.ToComError(ExcelError.ExcelErrorNA);
            }

            protected internal Topic(ExcelRtdServer server, int topicId, object initialValue)
            {
                _server = server;
                _topicId = topicId;
                _value = initialValue;
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
        HashSet<Topic> _dirtyTopics = new HashSet<Topic>();
        RtdUpdateSynchronization _updateSync;
        RTDUpdateEvent _callbackObject;
        readonly object _updateLock = new object();

        /// <summary>
        /// Performs a batch update of multiple topics, ensuring that all updates are visible to Excel at the same time.
        /// </summary>
        /// <param name="topics">List or array of topics to update</param>
        /// <param name="values">List or array of values matching the topics</param>
        public void UpdateValues(IList<Topic> topics, IList<object> values)
        {
            if (topics == null)
                throw new ArgumentNullException("topics");
            if (values == null)
                throw new ArgumentNullException("values");
            if (topics.Count != values.Count)
                throw new ArgumentException("Number of values must match number of topics");

            lock (_updateLock)
            {
                for (int i = 0; i < topics.Count; i++)
                {
                    var topic = topics[i];
                    var value = values[i];
                    // Call the real Topic.UpdateValue so that values get normalized and server notified
                    topic.UpdateValue(value);
                }
            }
        }

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

        // These three methods are added to allow derived classes to track UpdateNotify and RefreshData calls
        // Can be overridden to track posted notifications
        protected void OnUpdateNotifyPostedInsideLock(IReadOnlyCollection<Topic> dirtyTopics) { }

        // Can be overridden to track notifications called in Excel
        protected void OnUpdateNotifyInvokedInsideLock(IReadOnlyCollection<Topic> dirtyTopics) { }

        // Can be overridden to track Refresh calls from Excel
        protected void OnRefreshDataProcessedInsideLock(IReadOnlyCollection<Topic> dirtyTopics) { }

        // Called from any thread, inside the update lock
        // Add the topic to the dirty set and calls UpdateNotify()
        void SetDirtyInsideLock(Topic topic)
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
            var dirtyTopicsWasEmpty = _dirtyTopics.Count == 0;
            _dirtyTopics.Add(topic);
            if (dirtyTopicsWasEmpty)
            {
                PostUpdateNotifyInsideLock();
            }
        }

        // Called from any thread, inside the update lock
        void PostUpdateNotifyInsideLock()
        {
            _updateSync.UpdateNotify(_callbackObject);
            OnUpdateNotifyPostedInsideLock(_dirtyTopics);
        }

        // This is the private implementation of the IRtdServer interface
        // All these interface methods can be called only on the main Excel thread
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

                _callbackObject = new RTDUpdateEvent(this, callbackObject);   // Wrap the callback object to allow logging our call
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

                // NOTE: 2016-11-04
                //       Before v 0.34 the topic was added to _activeTopics before ConnectData was called
                //       The effect of moving it after (hence that topic is not in _activeTopics during the ConnectData call)
                //       is that a call to UpdateValue during the the ConnectData call will no longer cause an Update call to Excel
                //       (since SetDirty is ignored for topics not in _activeTopics)
                object value;
                using (XlCall.Suspend())
                {
                    value = ConnectData(topic, topicInfo, ref newValues);
                }
                _activeTopics[topicId] = topic;

                // Now we need to ensure that the topic value does indeed agree with the returned value
                // Otherwise we are left with an inconsistent state for future updates.
                // If there's a difference, we do force the update.
                if (!object.Equals(value, topic.Value))
                {
                    // 2020-03-03 v1.1
                    // Changing from topic.UpdateNotify, which no longer (in recent Excel) seems to work on its own.
                    // NOTE: In the unusual case that the FixValue inside Topic makes value == topic.Value, we won't get an update here
                    //       E.g. for long strings that are truncated
                    // 
                    topic.UpdateValue(value);
                }
                return value;
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
            HashSet<Topic> dirty;
            HashSet<Topic> newDirtyTopics = new HashSet<Topic>();
            lock (_updateLock)
            {
                dirty = _dirtyTopics;
                _dirtyTopics = newDirtyTopics;
                OnRefreshDataProcessedInsideLock(dirty);
            }

            // The topics in _dirtyTopics may have been Disconnected already.
            // (With another thread updating the value and setting dirty afterwards)
            // We assume Excel doesn't mind being notified of Disconnected topics.
            topicCount = dirty.Count;
            object[,] result = new object[2, topicCount];
            int i = 0;
            foreach (Topic topic in dirty)
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
                lock (_updateLock)
                {
                    _dirtyTopics.Remove(topic);
                }
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

        // Called by Excel if more than HeartbeatInterval millseconds have elapsed since the last UpdateNotify call
        // HeartbeatInterval cannot be less than 15000, but Heartbeat support can be switched off (-1)
        // So our redundant UpdateNotify call should not be happening often
        // We use the Heartbeat to retry any outstanding UpdateNotify call
        // (though this should not be needed according to the RTD interface contract)
        int IRtdServer.Heartbeat()
        {
            try
            {
                var updateNotifyPosted = false;
                // Re-post the notify in case something went wrong on the Excel side or our sync window
                // This should not be necessary but should not be harmful either
                // There has been a report of the RTD server stopping, perhaps due to sync window PostMessage failing
                // This is just an extra safety measure
                lock (_updateLock)
                {
                    if (_dirtyTopics.Count > 0)
                    {
                        // We've notified but Excel has not called us back, and Excel has not seen an UpdateNotify recently...
                        // This is unexpected
                        // We might be able to call UpdateNotify on Excel directly here, instead of scheduling it
                        // But it is convenient to allow Heartbeat to be called from any thread
                        PostUpdateNotifyInsideLock();
                        updateNotifyPosted = true;
                    }
                }

                if (updateNotifyPosted)
                {
                    Logger.RtdServer.Warn("Heartbeat callback while Notified in RTD server {0} - retrying UpdateNotify", GetType().Name);
                }

                using (XlCall.Suspend())
                {
                    // Call the derived class's override
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

        // We create a small wrapper around the IRTDUpdateEvent to allow logging of the UpdateNotify call
        class RTDUpdateEvent : IRTDUpdateEvent
        {
            readonly ExcelRtdServer _server;
            readonly IRTDUpdateEvent _inner;
            public RTDUpdateEvent(ExcelRtdServer server, IRTDUpdateEvent inner)
            {
                _server = server;
                _inner = inner;
            }

            public void UpdateNotify()
            {
                lock (_server._updateLock)
                {
                    _inner.UpdateNotify();
                    _server.OnUpdateNotifyInvokedInsideLock(_server._dirtyTopics);
                }
            }

            internal void CallInnerUpdateNotify()
            {
                _inner.UpdateNotify();
            }

            public int HeartbeatInterval
            {
                get { return _inner.HeartbeatInterval; }
                set { _inner.HeartbeatInterval = value; }
            }

            public void Disconnect()
            {
                _inner.Disconnect();
            }
        }
    }
}
