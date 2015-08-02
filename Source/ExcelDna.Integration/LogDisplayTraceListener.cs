//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Text;

namespace ExcelDna.Logging
{
    // CONSIDER: This TraceListener might co-operate with a more structured LogDisplay in future. 
    //           This might allow the source and EventType (and maybe a sequence) to be separate columns.
    public class LogDisplayTraceListener : TraceListener
    {
        public LogDisplayTraceListener()
        {
        }

        public LogDisplayTraceListener(string name)
            : base(name)
        {
        }

        public override void TraceEvent(TraceEventCache eventCache, string source, TraceEventType eventType, int id, string message)
        {
            TraceEvent(eventCache, source, eventType, id, message, null);
        }

        public override void TraceEvent(TraceEventCache eventCache, string source, TraceEventType eventType, int id, string format, params object[] args)
        {
            if (Filter != null && !Filter.ShouldTrace(eventCache, source, eventType, id, format, args, null, null))
                return;

            string idDescription;
            if (source == "ExcelDna.Integration")
            {
                // For this source, we interpret the event id as a grouping
                IntegrationTraceEventId traceEventId = (IntegrationTraceEventId)id;
                idDescription = traceEventId.ToString();
            }
            else
            {
                idDescription = id.ToString(CultureInfo.InvariantCulture);
            }
            string header = string.Format(CultureInfo.InvariantCulture, "{0} [{1}] ", idDescription, eventType.ToString());
            base.TraceEvent(eventCache, source, eventType, id, header + format, args);

            if (eventType == TraceEventType.Error || eventType == TraceEventType.Critical)
                LogDisplay.Show();
        }

        // Normally receives the header information
        // We just suppress for now.
        public override void Write(string message)
        {
            // CONSIDER: We might write to a buffer or special structure before displaying.
            //WriteLine(message);
        }

        public override void WriteLine(string message)
        {
            try
            {
                LogDisplay.RecordLine(message);
            }
            catch (Exception e)
            {
                Debug.Print("LogDisplayTraceListener error: " + e.Message);
            }
        }
    }
}
