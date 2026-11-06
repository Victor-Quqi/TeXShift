using System;
using System.Diagnostics;
using System.Reflection;
using System.Xml.Linq;
using TeXShift.Core.Logging;

namespace TeXShift.AddIn
{
    /// <summary>Correlates Ribbon callbacks with their outcome without logging note contents.</summary>
    internal sealed class AddInCommandLog : IDisposable
    {
        private readonly string _id = Guid.NewGuid().ToString("N");
        private readonly string _command;
        private readonly Stopwatch _timer = Stopwatch.StartNew();
        private readonly IDisposable _performance;
        public string PageId { get; set; }
        public bool Success { get; set; }

        public AddInCommandLog(string command)
        {
            _command = command;
            Write("started");
            _performance = PerformanceTraceContext.BeginCommand(command, _id);
        }

        public void Dispose()
        {
            _performance.Dispose();
            Write(Success ? "succeeded" : "failed");
        }

        private void Write(string status)
        {
            try
            {
                var assembly = typeof(Connect).Assembly;
                RuntimeLog.Write(new XElement("AddInCommand",
                    new XAttribute("id", _id), new XAttribute("command", _command),
                    new XAttribute("status", status), new XAttribute("utc", DateTime.UtcNow.ToString("O")),
                    new XAttribute("pid", Process.GetCurrentProcess().Id),
                    new XAttribute("assembly", assembly.Location),
                    new XAttribute("version", assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? ""),
                    new XAttribute("pageId", PageId ?? ""),
                    new XAttribute("durationMs", _timer.ElapsedMilliseconds)).ToString(SaveOptions.DisableFormatting));
            }
            catch (Exception)
            {
                // Diagnostics must not escape a COM callback or change the conversion result.
            }
        }
    }
}
