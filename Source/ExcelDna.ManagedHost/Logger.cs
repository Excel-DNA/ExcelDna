using System;
using System.Diagnostics;
using System.Globalization;

namespace ExcelDna.ManagedHost
{
    class Logger
    {
        // NOTE: This enum appears here and in TraceLogging in ExcelDna.Integration and in the Logging in Loader?
        enum IntegrationTraceEventId
        {
            Initialization = 1,
        }

        class TraceLogger
        {
            static internal TraceSource IntegrationTraceSource; // Set after Integration is initialized
        }


        internal static void SetIntegrationTraceSource(TraceSource traceSource) => TraceLogger.IntegrationTraceSource = traceSource;

        int _eventId;
        Logger(IntegrationTraceEventId traceEventId)
        {
            _eventId = (int)traceEventId;
        }

        void Log(TraceEventType eventType, string message, params object[] args)
        {
            if (TraceLogger.IntegrationTraceSource == null)
            {
                // We are in the pre-initialization stage. Just log to Debug.
                // NOTE: Without the explcit check and short-circuit here the loading would (sometimes!) fail.
                //       This is despite the catch-all exception handler that should be suppressing all errors.
                //       Somehow the null-reference exception being thrown and caught interferes with the AppDomain_AssemblyResolve!?
                //       The problem only happened when ExcelDna.Integration had to be resolved from resources (in which case the calls here would come from AssemblyResolve.
                //       Under the debugger, the problem was inconsistent (sometimes the load worked fine), but outside the debugger it always failed.
                Debug.Print(eventType.ToString() + " : " + string.Format(message, args));
                return;
            }
            try
            {
                TraceLogger.IntegrationTraceSource.TraceEvent(eventType, _eventId, message, args);
            }
            catch (Exception e)
            {
                // We certainly want to suppress errors here, though they indicate Excel-DNA bugs.
                Debug.Print("ExcelDna.Loader - Logger.Log error: " + e.Message);
            }
        }

        public void Verbose(string message, params object[] args) => Log(TraceEventType.Verbose, message, args);
        public void Info(string message, params object[] args) => Log(TraceEventType.Information, message, args);
        public void Warn(string message, params object[] args) => Log(TraceEventType.Warning, message, args);
        public void Error(string message, params object[] args) => Log(TraceEventType.Error, message, args);

        public void Error(Exception ex, string message, params object[] args)
        {
            if (args != null)
            {
                try
                {
                    message = string.Format(CultureInfo.InvariantCulture, message, args);
                }
                catch (Exception fex)
                {
                    Debug.Print("Logger.Error formatting exception " + fex.Message);
                }
            }
            Log(TraceEventType.Error, $"{message} : {ex.GetType().Name} - {ex.Message}");
        }


        static Logger _initializationLogger = new Logger(IntegrationTraceEventId.Initialization);
        static internal Logger Initialization => _initializationLogger;
    }
}
