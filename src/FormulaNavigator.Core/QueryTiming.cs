using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;

namespace FormulaNavigator.Core
{
    // Measures the stages of one Ctrl+Shift+Q request so the results window can
    // show how long each part of the work took. The clock starts when the command
    // handler begins, before the window is created or Excel is touched. The clock
    // is injectable so the description is testable without a real timer.
    public sealed class QueryTiming
    {
        private readonly Func<long> _elapsedMilliseconds;
        private long _windowShown = -1;
        private long _listReady = -1;
        private long _detailsDone = -1;
        private long _traceMilliseconds = -1;
        private long _searchMilliseconds = -1;

        public QueryTiming()
        {
            Stopwatch clock = Stopwatch.StartNew();
            _elapsedMilliseconds = delegate { return clock.ElapsedMilliseconds; };
        }

        public QueryTiming(Func<long> elapsedMilliseconds)
        {
            if (elapsedMilliseconds == null) throw new ArgumentNullException("elapsedMilliseconds");
            _elapsedMilliseconds = elapsedMilliseconds;
        }

        // Starts a new measurement for a repeated request in the same window. The
        // window-shown mark belongs to the first request and is kept.
        public void BeginSearch()
        {
            _listReady = -1;
            _detailsDone = -1;
            _traceMilliseconds = -1;
            _searchMilliseconds = -1;
        }

        // Called when the window is first shown, before the search starts.
        public void MarkWindowShown()
        {
            if (_windowShown < 0) _windowShown = _elapsedMilliseconds();
        }

        // Called when the address list is built and rendered for the first time.
        public void MarkListReady()
        {
            if (_listReady < 0) _listReady = _elapsedMilliseconds();
        }

        // Called when the deferred formulas and values finished loading.
        public void MarkDetailsDone()
        {
            if (_detailsDone < 0) _detailsDone = _elapsedMilliseconds();
        }

        // The tracer stage duration, measured by the gateway around the trace.
        public void RecordTrace(long milliseconds)
        {
            if (milliseconds >= 0) _traceMilliseconds = milliseconds;
        }

        // The search stage duration (trace plus cleanup), measured by the gateway.
        public void RecordSearch(long milliseconds)
        {
            if (milliseconds >= 0) _searchMilliseconds = milliseconds;
        }

        public string Describe()
        {
            var parts = new List<string>();
            if (_windowShown >= 0) parts.Add("окно " + _windowShown + " мс");
            if (_traceMilliseconds >= 0) parts.Add("трассировка " + _traceMilliseconds + " мс");
            if (_listReady >= 0) parts.Add("список " + _listReady + " мс");
            if (_detailsDone >= 0) parts.Add("детали " + _detailsDone + " мс");
            return string.Join(" · ", parts.ToArray());
        }

        public string SearchSeconds()
        {
            if (_searchMilliseconds < 0) return "";
            return (_searchMilliseconds / 1000.0).ToString("0.00", CultureInfo.InvariantCulture);
        }
    }
}
