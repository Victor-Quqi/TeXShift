using System;
using System.Globalization;
using System.IO;
using System.Text;
using TeXShift.Core.Utils;

namespace TeXShift.Core.Logging
{
    /// <summary>
    /// Writes small, always-on runtime diagnostics without affecting application flow.
    /// Independent of the debug settings on purpose: the failures worth diagnosing happen
    /// on the ordinary conversion path, where debug artifacts are never produced.
    /// </summary>
    public static class RuntimeLog
    {
        private const long MaxLogSizeBytes = 1024 * 1024;
        private const string LogFileName = "runtime.log";
        private const string OldLogFileName = "runtime.old.log";

        private static readonly object SyncRoot = new object();
        private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

        public static void Write(string message)
        {
            try
            {
                lock (SyncRoot)
                {
                    var outputFolder = TeXShiftPaths.RuntimeLogFolder;
                    Directory.CreateDirectory(outputFolder);

                    var logPath = Path.Combine(outputFolder, LogFileName);
                    var oldLogPath = Path.Combine(outputFolder, OldLogFileName);
                    var entry = FormatEntry(message);

                    RotateIfNeeded(logPath, oldLogPath, Utf8NoBom.GetByteCount(entry));
                    File.AppendAllText(logPath, entry, Utf8NoBom);
                }
            }
            catch (Exception)
            {
                // Runtime diagnostics must never affect application flow.
            }
        }

        private static string FormatEntry(string message)
        {
            var timestamp = DateTimeOffset.Now.ToString(
                "yyyy-MM-dd HH:mm:ss.fff zzz",
                CultureInfo.InvariantCulture);
            var prefix = "[" + timestamp + "] ";
            var normalized = (message ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace('\r', '\n');

            return prefix
                + normalized.Replace("\n", Environment.NewLine + prefix)
                + Environment.NewLine;
        }

        private static void RotateIfNeeded(string logPath, string oldLogPath, int incomingBytes)
        {
            if (!File.Exists(logPath))
            {
                return;
            }

            var currentSize = new FileInfo(logPath).Length;
            if (currentSize + incomingBytes <= MaxLogSizeBytes)
            {
                return;
            }

            if (File.Exists(oldLogPath))
            {
                File.Delete(oldLogPath);
            }

            File.Move(logPath, oldLogPath);
        }
    }
}
