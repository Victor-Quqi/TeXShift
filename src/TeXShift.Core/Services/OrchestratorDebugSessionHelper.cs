using System;
using System.Threading.Tasks;
using TeXShift.Core.Logging;
using OneNoteInterop = Microsoft.Office.Interop.OneNote;

namespace TeXShift.Core.Services
{
    internal static class OrchestratorDebugSessionHelper
    {
        internal static void InitializeDebugSession(
            ServiceContainer serviceContainer,
            bool writeDebugFiles,
            string outputDirectory,
            bool dumpFullPageXml,
            DebugSessionKind sessionKind,
            Action<string> setDebugOutputFolder,
            Action addSessionMetrics,
            out IDebugLogger logger,
            out PerformanceTrace perfTrace,
            out IDisposable perfSession)
        {
            logger = null;
            perfTrace = null;
            perfSession = null;

            if (!writeDebugFiles)
            {
                return;
            }

            logger = serviceContainer.CreateDebugLogger(outputDirectory);
            logger.StartSession(sessionKind);
            setDebugOutputFolder(logger.DebugSessionFolder);

            perfTrace = new PerformanceTrace();
            perfSession = PerformanceTraceContext.BeginSession(perfTrace);
            PerformanceTraceContext.AddMetric("SessionKind", sessionKind.ToString());
            addSessionMetrics?.Invoke();
            PerformanceTraceContext.AddMetric("DumpFullPageXml", dumpFullPageXml.ToString());
        }

        internal static async Task LogFinalPageXmlAsync(
            OneNoteInterop.Application oneNoteApp,
            bool writeDebugFiles,
            bool dumpFullPageXml,
            IDebugLogger logger,
            string pageId)
        {
            if (!writeDebugFiles || logger == null || string.IsNullOrWhiteSpace(pageId))
            {
                return;
            }

            string finalPageXml;
            using (PerformanceTraceContext.Measure("Write.GetFinalPageContent.Basic", "piBasic"))
            {
                finalPageXml = await Task.Run(() =>
                {
                    oneNoteApp.GetPageContent(pageId, out string xml, OneNoteInterop.PageInfo.piBasic, OneNoteInterop.XMLSchema.xs2013);
                    return xml;
                }).ConfigureAwait(false);
            }

            using (PerformanceTraceContext.Measure("Debug.Write.FinalPageXml.Basic"))
            {
                await logger.LogFinalPageXmlAsync(finalPageXml).ConfigureAwait(false);
            }

            if (!dumpFullPageXml)
            {
                return;
            }

            string fullPageXml;
            using (PerformanceTraceContext.Measure("Write.GetFinalPageContent.Full", "piAll"))
            {
                fullPageXml = await Task.Run(() =>
                {
                    oneNoteApp.GetPageContent(pageId, out string xml, OneNoteInterop.PageInfo.piAll, OneNoteInterop.XMLSchema.xs2013);
                    return xml;
                }).ConfigureAwait(false);
            }

            using (PerformanceTraceContext.Measure("Debug.Write.FinalPageXml.Full"))
            {
                await logger.LogFinalPageXmlFullAsync(fullPageXml).ConfigureAwait(false);
            }
        }

        internal static async Task TryLogPerformanceAsync(
            bool writeDebugFiles,
            IDebugLogger logger,
            PerformanceTrace perfTrace)
        {
            if (!writeDebugFiles || logger == null || perfTrace == null)
            {
                return;
            }

            try
            {
                await logger.LogPerformanceAsync(perfTrace.FormatReport()).ConfigureAwait(false);
            }
            catch
            {
                // Avoid surfacing performance logging failures to the user.
            }
        }
    }
}
