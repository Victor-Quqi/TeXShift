using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using TeXShift.Core.Logging;
using TeXShift.Core.Localization;
using TeXShift.Core.OneNote;
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
                InitializeDebugSession(options, result, out logger, out perfTrace, out perfSession);

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
                if (TryPromoteFullOutlineSelectionToCursor(readResult, out var promoted))
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

                await LogFinalPageXmlAsync(options, logger, readResult.PageId).ConfigureAwait(false);

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
                await TryLogPerformanceAsync(options, logger, perfTrace).ConfigureAwait(false);
                perfSession?.Dispose();
            }
        }

        private void InitializeDebugSession(
            ConversionOptions options,
            ConversionResult result,
            out IDebugLogger logger,
            out PerformanceTrace perfTrace,
            out IDisposable perfSession)
        {
            logger = null;
            perfTrace = null;
            perfSession = null;

            if (options == null || result == null || !options.WriteDebugFiles)
            {
                return;
            }

            logger = _serviceContainer.CreateDebugLogger(options.OutputDirectory);
            logger.StartSession(DebugSessionKind.ForwardConversion);
            result.DebugOutputFolder = logger.DebugSessionFolder;

            perfTrace = new PerformanceTrace();
            perfSession = PerformanceTraceContext.BeginSession(perfTrace);
            PerformanceTraceContext.AddMetric("SessionKind", "ForwardConversion");
            PerformanceTraceContext.AddMetric("ExportPdf", options.ExportPdf.ToString());
            PerformanceTraceContext.AddMetric("DumpFullPageXml", options.DumpFullPageXml.ToString());
        }

        private async Task LogFinalPageXmlAsync(ConversionOptions options, IDebugLogger logger, string pageId)
        {
            if (options == null || !options.WriteDebugFiles || logger == null || string.IsNullOrWhiteSpace(pageId))
            {
                return;
            }

            string finalPageXml;
            using (PerformanceTraceContext.Measure("Write.GetFinalPageContent.Basic", "piBasic"))
            {
                finalPageXml = await Task.Run(() =>
                {
                    _oneNoteApp.GetPageContent(pageId, out string xml, OneNoteInterop.PageInfo.piBasic, OneNoteInterop.XMLSchema.xs2013);
                    return xml;
                }).ConfigureAwait(false);
            }

            using (PerformanceTraceContext.Measure("Debug.Write.FinalPageXml.Basic"))
            {
                await logger.LogFinalPageXmlAsync(finalPageXml).ConfigureAwait(false);
            }

            if (!options.DumpFullPageXml)
            {
                return;
            }

            string fullPageXml;
            using (PerformanceTraceContext.Measure("Write.GetFinalPageContent.Full", "piAll"))
            {
                fullPageXml = await Task.Run(() =>
                {
                    _oneNoteApp.GetPageContent(pageId, out string xml, OneNoteInterop.PageInfo.piAll, OneNoteInterop.XMLSchema.xs2013);
                    return xml;
                }).ConfigureAwait(false);
            }

            using (PerformanceTraceContext.Measure("Debug.Write.FinalPageXml.Full"))
            {
                await logger.LogFinalPageXmlFullAsync(fullPageXml).ConfigureAwait(false);
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

        private static async Task TryLogPerformanceAsync(ConversionOptions options, IDebugLogger logger, PerformanceTrace perfTrace)
        {
            if (options == null || !options.WriteDebugFiles || logger == null || perfTrace == null)
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

        private static bool TryPromoteFullOutlineSelectionToCursor(ReadResult readResult, out ReadResult promoted)
        {
            promoted = null;
            if (readResult == null || !readResult.IsSuccess)
            {
                return false;
            }

            if (readResult.Mode != DetectionMode.Selection)
            {
                return false;
            }

            if (readResult.OriginalXmlNodes == null || readResult.OriginalXmlNodes.Count == 0)
            {
                return false;
            }

            var firstOe = readResult.OriginalXmlNodes.FirstOrDefault();
            if (firstOe == null)
            {
                return false;
            }

            var ns = firstOe.Name.Namespace;
            var outline = firstOe.Ancestors(ns + "Outline").FirstOrDefault();
            if (outline == null)
            {
                return false;
            }

            string outlineObjectId = outline.Attribute("objectID")?.Value;
            if (string.IsNullOrWhiteSpace(outlineObjectId))
            {
                return false;
            }

            var rootChildren = outline.Element(ns + "OEChildren");
            if (rootChildren == null)
            {
                return false;
            }

            var rootOes = rootChildren.Elements(ns + "OE").ToList();
            if (rootOes.Count == 0)
            {
                return false;
            }

            // Map selected OEs to root-level OEs under the outline.
            var selectedRootOes = readResult.OriginalXmlNodes
                .Select(oe => FindRootOe(oe, rootChildren, ns))
                .Where(oe => oe != null)
                .Distinct()
                .ToList();

            if (selectedRootOes.Count != rootOes.Count)
            {
                return false;
            }

            if (!rootOes.All(selectedRootOes.Contains))
            {
                return false;
            }

            promoted = new ReadResult
            {
                IsSuccess = true,
                Mode = DetectionMode.Cursor,
                ExtractedText = readResult.ExtractedText,
                PageId = readResult.PageId,
                OriginalXmlNode = outline,
                SourceOutlineWidth = readResult.SourceOutlineWidth
            };
            promoted.TargetObjectIds.Add(outlineObjectId);
            return true;
        }

        private static XElement FindRootOe(XElement oe, XElement rootChildren, XNamespace ns)
        {
            if (oe == null || rootChildren == null)
            {
                return null;
            }

            // Walk up: OE -> OEChildren -> OE -> ... until OE's parent is the outline's root OEChildren.
            var current = oe;
            while (current != null)
            {
                if (current.Parent == rootChildren)
                {
                    return current;
                }

                // OE's parent is OEChildren; OEChildren's parent is OE.
                var parentOe = current.Parent?.Parent;
                if (parentOe == null || parentOe.Name != ns + "OE")
                {
                    return null;
                }

                current = parentOe;
            }

            return null;
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
