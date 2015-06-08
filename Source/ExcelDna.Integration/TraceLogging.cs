using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Security;
using System.Text;

namespace ExcelDna.Integration
{
    // This class supports intenal logging.
    // Internal logging is implemented with the System.Diagnostics tracing implementation.

    // Add a trace listener for the ExcelDna.Integration source which logs warnings and errors to the LogDisplay 
    // (only popping up the window for errors).
    // Verbose logging can be configured via the .config file

    // We define a TraceSource called ExcelDna.Integration (that is also exported to ExcelDna.Loader and called from there)
    // We consolidate the two assemblies against a single TraceSource, since ExcelDna.Integration is the only public contract,
    // and we expect to move more of the registration into the ExcelDna.Integration assembly in future.

    #region Microsoft License
    // The logging helper implementation here is adapted from the Logging.cs file for System.Net
    // Taken from https://github.com/Microsoft/referencesource/blob/c697a4b9782dc8c85c02344a847cb68163702aa7/System/net/System/Net/Logging.cs
    // Under the following license:
    //
    // The MIT License (MIT)

    // Copyright (c) Microsoft Corporation
       
    // Permission is hereby granted, free of charge, to any person obtaining a copy 
    // of this software and associated documentation files (the "Software"), to deal 
    // in the Software without restriction, including without limitation the rights 
    // to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
    // copies of the Software, and to permit persons to whom the Software is 
    // furnished to do so, subject to the following conditions: 
       
    // The above copyright notice and this permission notice shall be included in all 
    // copies or substantial portions of the Software. 
       
    // THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
    // IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
    // FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
    // AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
    // LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, 
    // OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE 
    // SOFTWARE.
    #endregion

    enum IntegrationTraceEventId
    {
        RegistrationInitialize = 1025,
        RegistrationEvent = 1026    // Everything is miscellaneous
    }

    class TraceLogging
    {
        static volatile bool s_LoggingEnabled = true;
        static volatile bool s_LoggingInitialized;
        static volatile bool s_AppDomainShutdown;
        const string TraceSourceName = "ExcelDna.Integration";
        internal static TraceSource IntegrationTraceSource; // Retrieved from ExcelDna.Loader through ExcelIntegration

        public static void Initialize()
        {
            if (!s_LoggingInitialized)
            {
                bool loggingEnabled = false;
                IntegrationTraceSource = new TraceSource(TraceSourceName, SourceLevels.All);
                //GlobalLog.Print("Initalizating tracing");

                try
                {
                    loggingEnabled = (IntegrationTraceSource.Switch.ShouldTrace(TraceEventType.Critical));
                }
                catch (SecurityException)
                {
                    // These may throw if the caller does not have permission to hook up trace listeners.
                    // We treat this case as though logging were disabled.
                    Close();
                    loggingEnabled = false;
                }
                if (loggingEnabled)
                {
                    AppDomain currentDomain = AppDomain.CurrentDomain;
                    //currentDomain.UnhandledException += UnhandledExceptionHandler;
                    currentDomain.DomainUnload += AppDomainUnloadEvent;
                    currentDomain.ProcessExit += ProcessExitEvent;
                }
                s_LoggingEnabled = loggingEnabled;
                s_LoggingInitialized = true;
            }
        }

        static bool ValidateSettings(TraceSource traceSource, TraceEventType traceLevel)
        {
            if (!s_LoggingEnabled)
            {
                return false;
            }
            if (!s_LoggingInitialized)
            {
                Initialize();
            }
            if (traceSource == null || !traceSource.Switch.ShouldTrace(traceLevel))
            {
                return false;
            }
            if (s_AppDomainShutdown)
            {
                return false;
            }
            return true;
        }

        static void ProcessExitEvent(object sender, EventArgs e) 
        {
            Close();
            s_AppDomainShutdown = true;
        }

        private static void AppDomainUnloadEvent(object sender, EventArgs e) 
        {
            Close();
            s_AppDomainShutdown = true;
        }

        static void Close()
        {
            if (IntegrationTraceSource != null) 
                IntegrationTraceSource.Close();
        }

    }

    class RegistrationLogging
    {
        public static void Log(TraceEventType eventType, string message, params object[] args)
        {
            Debug.Write(string.Format("RegistrationLogging: {0:yyyy-MM-dd HH:mm:ss} {1} {2}\r\n", DateTime.Now, eventType, string.Format(message, args)));

            TraceLogging.IntegrationTraceSource.TraceEvent(eventType, (int)IntegrationTraceEventId.RegistrationEvent, message, args);
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
