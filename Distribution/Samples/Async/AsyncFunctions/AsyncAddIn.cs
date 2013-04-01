using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;

namespace AsyncFunctions
{
    public class AsyncTestAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(ex => "!!! EXCEPTION: " + ex.ToString());
        }

        public void AutoClose()
        {
        }
    }

    public static class MyFunctions
    {
        public static object WhatTimeIsIt()
        {
            return XlCall.RTD("MyRTDServers.TimeServer", null, "NOW");
        }

        public static object WhatTimeIsItEx(string input)
        {
            return XlCall.RTD("MyRTDServers.TimeServer", null, "NOW", input);
        }
    }

    namespace MyRTDServers
    {
        // [ComVisible(true)]
        public class TimeServer : ExcelRtdServer
        {
            string _logPath;
            List<Topic> _topics;
            Timer _timer;
            public TimeServer()
            {
                Log("TimerServer created");
                _logPath = Path.ChangeExtension((string)XlCall.Excel(XlCall.xlGetName), ".log");
                _topics = new List<Topic>();
                _timer = new Timer(delegate
                    {
                        Log("Tick");
                        string now = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff");
                        foreach (Topic topic in _topics) topic.UpdateValue(now);
                    }, null, 0, 1000);
            }

            void Log(string format, params object[] args)
            {
                File.AppendAllText(_logPath, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") + " - " + string.Format(format, args));
            }

            int GetTopicId(Topic topic)
            {
                return (int)typeof(Topic)
                            .GetField("TopicId", BindingFlags.GetField | BindingFlags.Instance | BindingFlags.NonPublic)
                            .GetValue(topic);
            }

            protected override bool ServerStart()
            {
                Log("ServerStart");
                return true;
            }

            protected override void ServerTerminate()
            {
                Log("ServerTerminate");
            }

            protected override object ConnectData(Topic topic, System.Collections.Generic.IList<string> topicInfo, ref bool newValues)
            {
                Log("ConnectData: {0} - {{{1}}}", GetTopicId(topic), string.Join(", ", topicInfo));
                return true;
            }

            protected override void DisconnectData(Topic topic)
            {
                Log("DisconnectData: {0}", GetTopicId(topic));
            }
        }
    }
}
