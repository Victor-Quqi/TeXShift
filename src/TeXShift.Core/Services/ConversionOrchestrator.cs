using System;
using System.IO;
using System.Threading.Tasks;
using System.Xml.Linq;
using TeXShift.Core.Logging;
using TeXShift.Core.Localization;
using TeXShift.Core.OneNote;
using TeXShift.Core.Services.ReverseConversion;
using OneNoteInterop = Microsoft.Office.Interop.OneNote;

namespace TeXShift.Core.Services
{
    /// <summary>
    /// Orchestrates the Read -> Convert -> Write pipeline.
    /// </summary>
    public class ConversionOrchestrator
    {
        private readonly ServiceContainer _serviceContainer;
        private readonly OneNoteInterop.Application _oneNoteApp;

        public ConversionOrchestrator(ServiceContainer serviceContainer, OneNoteInterop.Application oneNoteApp)
        {
            if (serviceContainer == null)
                throw new ArgumentNullException(nameof(serviceContainer));
            if (oneNoteApp == null)
                throw new ArgumentNullException(nameof(oneNoteApp));

            _serviceContainer = serviceContainer;
            _oneNoteApp = oneNoteApp;
        }

        /// <summary>
        /// Executes the conversion pipeline with optional debug logging and PDF export.
        /// </summary>
        public async Task<ConversionResult> ExecuteAsync(ConversionOptions options)
        {
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            IDebugLogger logger = null;
            PerformanceTrace perfTrace = null;
            IDisposable perfSession = null;
            var result = new ConversionResult();

            try
            {
                OrchestratorDebugSessionHelper.InitializeDebugSession(
                    _serviceContainer,
                    options.WriteDebugFiles,
                    options.OutputDirectory,
                    options.DumpFullPageXml,
                    DebugSessionKind.ForwardConversion,
                    folder => result.DebugOutputFolder = folder,
                    () => PerformanceTraceContext.AddMetric("ExportPdf", options.ExportPdf.ToString()),
                    out logger,
                    out perfTrace,
                    out perfSession);

                var reader = _serviceContainer.CreateContentReader(_oneNoteApp);
                ReadResult readResult;
                using (PerformanceTraceContext.Measure("Read.ExtractContentAsync"))
                {
                    readResult = await reader.ExtractContentAsync().ConfigureAwait(false);
                }
                result.ReadResult = readResult;

                if (!readResult.IsSuccess)
                {
                    // Don't wrap read errors in an exception - let Connect.cs handle
                    // ReadResult.ErrorMessage directly for user-friendly display
                    return result;
                }

                // OneNote sometimes reports a caret inside a rich object (or a single-paragraph text box) as Selection mode.
                // When the selection covers an entire outline, promote to Cursor mode so we replace the whole outline and keep TeXShift meta consistent.
                if (ReverseSelectionPromoter.TryPromoteFullOutlineSelectionToCursor(
                    readResult,
                    requireValidMeta: false,
                    preservePageHasTodoTagDef: false,
                    out var promoted))
                {
                    readResult = promoted;
                    result.ReadResult = promoted;
                }

                // Forward conversion requires actual Markdown text. The reader may report success for non-text content
                // (e.g., Image/Table) to support reverse conversion and debug tooling, so guard here.
                if (string.IsNullOrWhiteSpace(readResult.ExtractedText))
                {
                    readResult.IsSuccess = false;
                    readResult.ErrorMessage = readResult.Mode == DetectionMode.Selection
                        ? Resources.GetString("Error_NoValidTextContent")
                        : Resources.GetString("Error_EmptyTextBox");
                    result.ReadResult = readResult;
                    return result;
                }

                if (options.WriteDebugFiles)
                {
                    PerformanceTraceContext.AddMetric("DetectionMode", readResult.Mode.ToString());
                    PerformanceTraceContext.AddMetric("ExtractedTextChars", (readResult.ExtractedText?.Length ?? 0).ToString());

                    using (PerformanceTraceContext.Measure("Debug.Write.InputMarkdown"))
                    {
                        await logger.LogInputMarkdownAsync(readResult.ExtractedText).ConfigureAwait(false);
                    }

                    using (PerformanceTraceContext.Measure("Debug.Write.OriginalXml"))
                    {
                        await logger.LogOriginalXmlAsync(readResult.OriginalXmlNode).ConfigureAwait(false);
                    }
                }

                var converter = _serviceContainer.CreateMarkdownConverter(readResult.SourceOutlineWidth);
                XElement oneNoteXml;
                using (PerformanceTraceContext.Measure("Convert.MarkdownToOneNoteXmlAsync"))
                {
                    oneNoteXml = await converter.ConvertToOneNoteXmlAsync(readResult.ExtractedText).ConfigureAwait(false);
                }

                if (options.WriteDebugFiles)
                {
                    using (PerformanceTraceContext.Measure("Debug.Write.ConvertedXml"))
                    {
                        await logger.LogConvertedXmlAsync(oneNoteXml).ConfigureAwait(false);
                    }
                }

                var writer = _serviceContainer.CreateContentWriter(_oneNoteApp);
                using (PerformanceTraceContext.Measure("Write.ReplaceContentAsync"))
                {
                    await writer.ReplaceContentAsync(readResult, oneNoteXml).ConfigureAwait(false);
                }

                await OrchestratorDebugSessionHelper.LogFinalPageXmlAsync(
                    _oneNoteApp,
                    options.WriteDebugFiles,
                    options.DumpFullPageXml,
                    logger,
                    readResult.PageId).ConfigureAwait(false);

                if (options.ExportPdf)
                {
                    if (!await TryExportPdfAsync(options, result, logger, readResult.PageId).ConfigureAwait(false))
                    {
                        return result;
                    }
                }

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Error = ex;
                if (options.WriteDebugFiles && logger != null)
                {
                    await logger.LogErrorAsync(ex).ConfigureAwait(false);
                }
                return result;
            }
            finally
            {
                await OrchestratorDebugSessionHelper.TryLogPerformanceAsync(
                    options.WriteDebugFiles,
                    logger,
                    perfTrace).ConfigureAwait(false);
                perfSession?.Dispose();
            }
        }

        private async Task<bool> TryExportPdfAsync(ConversionOptions options, ConversionResult result, IDebugLogger logger, string pageId)
        {
            if (options == null || result == null || string.IsNullOrWhiteSpace(pageId))
            {
                return false;
            }

            string pdfPath = ResolvePdfOutputPath(options.OutputDirectory, result.DebugOutputFolder);
            var publisher = _serviceContainer.CreateOneNotePublisher(_oneNoteApp);
            bool exported;
            using (PerformanceTraceContext.Measure("Export.ExportToPdfAsync"))
            {
                exported = await publisher.ExportToPdfAsync(pageId, pdfPath).ConfigureAwait(false);
            }
            if (!exported)
            {
                var error = new IOException("PDF export failed.");
                result.Error = error;
                if (options.WriteDebugFiles && logger != null)
                {
                    await logger.LogErrorAsync(error).ConfigureAwait(false);
                }
                return false;
            }
            result.PdfPath = pdfPath;
            return true;
        }


        private string ResolvePdfOutputPath(string outputDirectory, string debugOutputFolder)
        {
            var baseDirectory = outputDirectory;
            if (string.IsNullOrWhiteSpace(baseDirectory))
            {
                baseDirectory = debugOutputFolder;
            }

            if (string.IsNullOrWhiteSpace(baseDirectory))
            {
                throw new InvalidOperationException("OutputDirectory is required for PDF export.");
            }

            string fileName = $"F05_Export_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
            return Path.Combine(baseDirectory, fileName);
        }
    }
}
