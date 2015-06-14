using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace ExcelDna.Loader
{
    // NOTE: This enum appears here and in TraceLogging in ExcelDna.Integration
    [Flags]
    enum IntegrationTraceEventId
    {
        Registration = 1 << 5,
        RegistrationInitialize = Registration + 1,
        RegistrationEvent = Registration + 2    // Everything is miscellaneous
    }


    // NOTE: There's a similar class in ExcelDna.Integration
    // RegistrationLogger is a thin helper for the ExcelDna.Integration TraceSource that we get from ExcelDna.Integration.
    // If we log more types of information in ExcelDna.Loader, we should add TraceSources, write more detailed messages, or at least sort out the internal abstraction a bit better.
    static class RegistrationLogger
    {
        internal static TraceSource IntegrationTraceSource; // Set after Integration is initialized

        public static void Log(TraceEventType eventType, string message, params object[] args)
        {
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

        public static void Error(Exception ex, string message, params object[] args)
        {
            Log(TraceEventType.Error, "{0} : {1} - {2}", message, ex.GetType().Name.ToString(), ex.Message);
        }

    }
}
