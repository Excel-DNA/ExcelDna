using System.Diagnostics;

namespace ExcelDna.Logging
{
    internal class DiagnosticsFilter : TraceFilter
    {
        private TraceEventType filterLevel;

        public DiagnosticsFilter(TraceEventType filterLevel)
        {
            this.filterLevel = filterLevel;
        }

        public override bool ShouldTrace(TraceEventCache cache, string source, TraceEventType eventType, int id, string formatOrMessage, object[] args, object data1, object[] data)
        {
            return eventType <= filterLevel;
        }
    }
}
