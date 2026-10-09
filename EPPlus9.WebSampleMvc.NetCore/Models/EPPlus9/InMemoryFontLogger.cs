using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

namespace EPPlus9.WebSampleMvc.NetCore.Models.EPPlus9
{
    /// <summary>
    /// Collects font log events in memory so they can be rendered after an export.
    /// Create one instance per export/request. Thread-safe.
    /// </summary>
    public sealed class InMemoryFontLogger : IFontLogger
    {
        private readonly object _lock = new object();
        private readonly List<FontLogEvent> _events = new List<FontLogEvent>();
        private readonly FontLogSeverity _minSeverity;
        private readonly int _maxEvents;
        private int _dropped;

        public InMemoryFontLogger(FontLogSeverity minSeverity = FontLogSeverity.Information, int maxEvents = 5000)
        {
            _minSeverity = minSeverity;
            _maxEvents = maxEvents;
        }

        public bool IsEnabled(FontLogSeverity severity) => severity >= _minSeverity;

        public void Log(FontLogEvent logEvent)
        {
            if (logEvent == null) return;
            lock (_lock)
            {
                if (_events.Count >= _maxEvents)
                {
                    _dropped++;
                    return;
                }
                _events.Add(logEvent);
            }
        }

        /// <summary>Returns a copy of the events collected so far.</summary>
        public IReadOnlyList<FontLogEvent> GetSnapshot()
        {
            lock (_lock)
            {
                return _events.ToArray();
            }
        }

        /// <summary>Number of events discarded because the cap was reached.</summary>
        public int DroppedCount
        {
            get { lock (_lock) { return _dropped; } }
        }
    }

    public sealed class PdfExportTableLogModel
    {
        public PdfExportTableLogModel(IReadOnlyList<FontLogEvent> events, FontLogSeverity minSeverity, int droppedCount, long pdfSize)
        {
            Events = events;
            MinSeverity = minSeverity;
            DroppedCount = droppedCount;
            PdfSize = pdfSize;
        }

        public IReadOnlyList<FontLogEvent> Events { get; }
        public FontLogSeverity MinSeverity { get; }
        public int DroppedCount { get; }
        public long PdfSize { get; }
    }
}