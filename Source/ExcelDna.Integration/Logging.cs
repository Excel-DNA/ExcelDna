//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Security;
using System.Text;

namespace ExcelDna.Logging
{
    // This class supports internal logging, implemented with the System.Diagnostics tracing implementation.

    // Add a trace listener for the ExcelDna.Integration source which logs warnings and errors to the LogDisplay 
    // (only popping up the window for errors).
    // Verbose logging can be configured via the .config file

    // We define a TraceSource called ExcelDna.Integration (that is also exported to ExcelDna.Loader and called from there)
    // We consolidate the two assemblies against a single TraceSource, since ExcelDna.Integration is the only public contract,
    // and we expect to move more of the registration into the ExcelDna.Integration assembly in future.

    // DOCUMENT: Info on custom TraceSources etc: https://msdn.microsoft.com/en-us/magazine/cc300790.aspx
    //           and http://blogs.msdn.com/b/kcwalina/archive/2005/09/20/tracingapis.aspx

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

    // NOTE: To simplify configuration (so that we provide one TraceSource per referenced assembly) and still allow some grouping
    //       we use the EventId to define a trace event classification.
    // NOTE: There's a copy of this enum in ExcelDna.Loader (in LoaderLogging.cs) too.
    enum IntegrationTraceEventId
    {
        Initialization = 1,
        DnaCompilation = 2, 
        Registration = 3,
        ComAddIn = 4,
        RtdServer = 5,
    }

    // TraceLogger manages the IntegrationTraceSource that we use for logging.
    // It deals with lifetime (particularly closing the TraceSource if the add-in is unloaded).
    // The default configuration of the TraceSource is set here, and can be overridden in the .xll.config file.
    class TraceLogger
    {
        static volatile bool s_LoggingEnabled = true;
        static volatile bool s_LoggingInitialized;
        static volatile bool s_AppDomainShutdown;
        const string TraceSourceName = "ExcelDna.Integration";
        internal static TraceSource IntegrationTraceSource; // Also retrieved by ExcelDna.Loader through ExcelIntegration.GetIntegrationTraceSource()

        public static void Initialize()
        {
            if (!s_LoggingInitialized)
            {
                bool loggingEnabled = false;
                // DOCUMENT: By default the TraceSource is configured to source only Warning, Error and Fatal.
                //           the configuration can override this.
                IntegrationTraceSource = new TraceSource(TraceSourceName, SourceLevels.Warning);

                bool logDisplayTraceListenerIsConfigured = false;
                Debug.Print("{0} TraceSource created. Listeners:", TraceSourceName);
                foreach (TraceListener tl in IntegrationTraceSource.Listeners)
                {
                    Debug.Print("    {0} - {1}", tl.Name, tl.TraceOutputOptions);
                    if (tl.Name == "LogDisplay")
                        logDisplayTraceListenerIsConfigured = true;
                }

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
                    if (!logDisplayTraceListenerIsConfigured)
                    {
                        // No explicit configuration for this default listener, so we add it
                        IntegrationTraceSource.Listeners.Add(new LogDisplayTraceListener("LogDisplay"));
                    }

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

    // NOTE: There is a similar RegistrationLogger class in ExcelDna.Loader.
    // It's easier to maintain two copies for now.
    class Logger
    {
        int _eventId;

        Logger(IntegrationTraceEventId traceEventId)
        {
            _eventId = (int)traceEventId;
        }

        void Log(TraceEventType eventType, string message, params object[] args)
        {
            try
            {
                TraceLogger.IntegrationTraceSource.TraceEvent(eventType, _eventId, message, args);
            }
            catch (Exception e)
            {
                Debug.Print("ExcelDna.Integration - Logger.Log error: " + e.Message);
            }
        }

        public void Verbose(string message, params object[] args)
        {
            Log(TraceEventType.Verbose, message, args);
        }

        public void Info(string message, params object[] args)
        {
            Log(TraceEventType.Information, message, args);
        }

        public void Warn(string message, params object[] args)
        {
            Log(TraceEventType.Warning, message, args);
        }

        public void Error(string message, params object[] args)
        {
            Log(TraceEventType.Error, message, args);
        }

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
            Log(TraceEventType.Error, "{0} : {1} - {2}", message, ex.GetType().Name, ex.Message);
        }

        static Logger _initializationLogger = new Logger(IntegrationTraceEventId.Initialization);
        static internal Logger Initialization { get { return _initializationLogger; } }
        static Logger _registrationLogger = new Logger(IntegrationTraceEventId.Registration);
        static internal Logger Registration { get { return _registrationLogger; } }
        static Logger _dnaCompilationLogger = new Logger(IntegrationTraceEventId.DnaCompilation);
        static internal Logger DnaCompilation { get { return _dnaCompilationLogger; } }
        static Logger _comAddInLogger = new Logger(IntegrationTraceEventId.ComAddIn);
        static internal Logger ComAddIn { get { return _comAddInLogger; } }
        static Logger _rtdServerLogger = new Logger(IntegrationTraceEventId.RtdServer);
        static internal Logger RtdServer { get { return _rtdServerLogger; } }
    }
}
