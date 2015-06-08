using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace ExcelDna.Loader.Logging
{
    // This enum appears here and in TraceLogging in ExcelDna.Integration
    internal enum IntegrationTraceEventId
    {
        RegistrationInitialize = 1025,
        RegistrationEvent = 1026    // Everything is miscellaneous
    }

    // RegistrationLogging is a thin helper for the ExcelDna.Integration TraceSource that we get from ExcelDna.Integration.
    // If we log more types of information in ExcelDna.Loader, we should add TraceSources, write more detailed messages, or at least sort out the internal abstraction a bit better.
    internal static class RegistrationLogging
    {
        internal static TraceSource IntegrationTraceSource; // Set after Integration is initialized

        public static void Log(TraceEventType eventType, string message, params object[] args)
        {
            Debug.Write(string.Format("RegistrationLogging: {0:yyyy-MM-dd HH:mm:ss} {1} {2}\r\n", DateTime.Now, eventType, string.Format(message, args)));

            IntegrationTraceSource.TraceEvent(eventType, (int)IntegrationTraceEventId.RegistrationEvent, message, args);
        }

        public static void Info(string message, params object[] args)
        {
            Log(TraceEventType.Information, message, args);
        }

        public static void Warn(string message, params object[] args)
        {
            Log(TraceEventType.Warning, message, args);
        }

        public static void Error(string message, params object[] args)
        {
            Log(TraceEventType.Error, message, args);
        }

        public static void ErrorException(string message, Exception ex)
        {
            Log(TraceEventType.Error, "{0} : {1} - {2}", message, ex.GetType().Name.ToString(), ex.Message);
        }

    }
}
