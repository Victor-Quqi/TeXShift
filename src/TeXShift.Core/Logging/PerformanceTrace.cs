using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Threading;

namespace TeXShift.Core.Logging
{
    internal sealed class PerformanceTrace
    {
        internal sealed class Entry
        {
            public int Depth { get; set; }
            public string Name { get; set; }
            public long DurationMs { get; set; }
            public string Detail { get; set; }
        }

        private readonly object _gate = new object();
        private readonly List<Entry> _entries = new List<Entry>();
        private readonly long _startTimestamp = Stopwatch.GetTimestamp();

        public void Add(Entry entry)
        {
            if (entry == null)
            {
                return;
            }

            lock (_gate)
            {
                _entries.Add(entry);
            }
        }

        public long GetTotalMs()
        {
            var elapsed = Stopwatch.GetTimestamp() - _startTimestamp;
            return (long)(elapsed * 1000.0 / Stopwatch.Frequency);
        }

        public string FormatReport()
        {
            var builder = new StringBuilder();
            builder.AppendLine("TeXShift Performance Report");
            builder.AppendLine($"TotalMs\t{GetTotalMs()}");
            builder.AppendLine();
            builder.AppendLine("Depth\tStep\tDurationMs\tDetail");

            List<Entry> snapshot;
            lock (_gate)
            {
                snapshot = new List<Entry>(_entries);
            }

            foreach (var entry in snapshot)
            {
                builder.Append(entry.Depth);
                builder.Append('\t');
                builder.Append(entry.Name ?? string.Empty);
                builder.Append('\t');
                builder.Append(entry.DurationMs);
                builder.Append('\t');
                builder.AppendLine(entry.Detail ?? string.Empty);
            }

            return builder.ToString();
        }
    }

    internal static class PerformanceTraceContext
    {
        private sealed class NoOpScope : IDisposable
        {
            public static readonly NoOpScope Instance = new NoOpScope();
            public void Dispose() { }
        }

        private sealed class TraceSession : IDisposable
        {
            private readonly PerformanceTrace _previous;
            private readonly int _previousDepth;

            public TraceSession(PerformanceTrace previous, int previousDepth)
            {
                _previous = previous;
                _previousDepth = previousDepth;
            }

            public void Dispose()
            {
                Current.Value = _previous;
                Depth.Value = _previousDepth;
            }
        }

        private sealed class MeasureScope : IDisposable
        {
            private readonly PerformanceTrace _trace;
            private readonly string _name;
            private readonly int _depth;
            private readonly long _startTimestamp;
            private readonly string _detail;

            public MeasureScope(PerformanceTrace trace, string name, int depth, string detail)
            {
                _trace = trace;
                _name = name;
                _depth = depth;
                _detail = detail;
                _startTimestamp = Stopwatch.GetTimestamp();
            }

            public void Dispose()
            {
                var elapsed = Stopwatch.GetTimestamp() - _startTimestamp;
                var ms = (long)(elapsed * 1000.0 / Stopwatch.Frequency);
                _trace.Add(new PerformanceTrace.Entry
                {
                    Depth = _depth,
                    Name = _name,
                    DurationMs = ms,
                    Detail = _detail
                });

                Depth.Value = _depth;
            }
        }

        private static readonly AsyncLocal<PerformanceTrace> Current = new AsyncLocal<PerformanceTrace>();
        private static readonly AsyncLocal<int> Depth = new AsyncLocal<int>();

        public static PerformanceTrace GetCurrent() => Current.Value;

        public static IDisposable BeginSession(PerformanceTrace trace)
        {
            var previous = Current.Value;
            var previousDepth = Depth.Value;

            Current.Value = trace;
            Depth.Value = 0;
            return new TraceSession(previous, previousDepth);
        }

        public static IDisposable Measure(string name, string detail = null)
        {
            var trace = Current.Value;
            if (trace == null)
            {
                return NoOpScope.Instance;
            }

            var depth = Depth.Value;
            Depth.Value = depth + 1;
            return new MeasureScope(trace, name ?? string.Empty, depth, detail);
        }

        public static void AddMetric(string name, string value)
        {
            var trace = Current.Value;
            if (trace == null)
            {
                return;
            }

            trace.Add(new PerformanceTrace.Entry
            {
                Depth = Depth.Value,
                Name = name ?? string.Empty,
                DurationMs = 0,
                Detail = value ?? string.Empty
            });
        }
    }
}

