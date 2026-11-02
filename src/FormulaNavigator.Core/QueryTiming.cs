using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace FormulaNavigator.Core
{
    // Stage timings of one Ctrl+Shift+Q request: the built-in trace, reading the
    // formulas and values, and the total time from the key press until the window
    // shows the list. The clock is injectable so the description is testable.
    public sealed class QueryTiming
    {
        private readonly Func<long> _elapsedMilliseconds;
        private long _start;
        private long _trace = -1;
        private long _details = -1;
        private long _ready = -1;

        public QueryTiming()
        {
            Stopwatch clock = Stopwatch.StartNew();
            _elapsedMilliseconds = delegate { return clock.ElapsedMilliseconds; };
        }

        public QueryTiming(Func<long> elapsedMilliseconds)
        {
            if (elapsedMilliseconds == null) throw new ArgumentNullException("elapsedMilliseconds");
            _elapsedMilliseconds = elapsedMilliseconds;
            _start = elapsedMilliseconds();
        }

        // Starts a repeated request (the Refresh button) from the current moment.
        public void Restart()
        {
            _start = _elapsedMilliseconds();
            _trace = -1;
            _details = -1;
            _ready = -1;
        }

        // Trace stage and total search time (trace plus formula reads), both
        // measured by the gateway.
        public void RecordSearch(long traceMilliseconds, long searchMilliseconds)
        {
            if (traceMilliseconds >= 0) _trace = traceMilliseconds;
            if (searchMilliseconds >= 0 && traceMilliseconds >= 0 && searchMilliseconds >= traceMilliseconds)
                _details = searchMilliseconds - traceMilliseconds;
        }

        // The list is rendered and the window is usable.
        public void MarkReady()
        {
            if (_ready < 0) _ready = _elapsedMilliseconds() - _start;
        }

        public string Describe()
        {
            var parts = new List<string>();
            if (_trace >= 0) parts.Add("трассировка " + _trace + " мс");
            if (_details >= 0) parts.Add("формулы " + _details + " мс");
            if (_ready >= 0) parts.Add("всего до окна " + _ready + " мс");
            return string.Join(" · ", parts.ToArray());
        }
    }
}
