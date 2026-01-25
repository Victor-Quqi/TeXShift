using System;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using TeXShift.Core.Logging;
using TeXShift.Core.OneNote;
using TeXShift.Core.OneNoteToMarkdown;
using TeXShift.Core.Services.ReverseConversion;
using TeXShift.Core.Utils;
using OneNoteInterop = Microsoft.Office.Interop.OneNote;

namespace TeXShift.Core.Services
{
    /// <summary>
    /// Orchestrates the Read (selection/cursor) -> Reverse Convert -> Write (as plain Markdown text) pipeline.
    /// </summary>
    public sealed class ReverseConversionOrchestrator
    {
        private readonly ServiceContainer _serviceContainer;
        private readonly OneNoteInterop.Application _oneNoteApp;

        public ReverseConversionOrchestrator(ServiceContainer serviceContainer, OneNoteInterop.Application oneNoteApp)
        {
            _serviceContainer = serviceContainer ?? throw new ArgumentNullException(nameof(serviceContainer));
            _oneNoteApp = oneNoteApp ?? throw new ArgumentNullException(nameof(oneNoteApp));
        }

        public async Task<ReverseConversionPipelineResult> ExecuteAsync(ReverseConversionOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            IDebugLogger logger = null;
            PerformanceTrace perfTrace = null;
            IDisposable perfSession = null;
            var result = new ReverseConversionPipelineResult();

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
                    // Let the caller handle ReadResult.ErrorMessage (no exception).
                    return result;
                }

                // If the selection covers the entire outline and TeXShift meta is valid,
                // promote to Cursor mode so reverse conversion can restore the stored source verbatim.
                if (ReverseSelectionPromoter.TryPromoteFullOutlineSelectionToCursorWithValidMeta(readResult, out var promotedFullOutline))
                {
                    readResult = promotedFullOutline;
                    result.ReadResult = promotedFullOutline;
                }

                // If the selection is inside a OneNote table, OneNote usually reports the selected *cell OEs*.
                // Writing back by deleting those OEs can leave <OEChildren> empty inside <Cell>, which OneNote rejects.
                // Promote to the table container OE so we replace the whole table atomically.
                if (ReverseSelectionPromoter.TryPromoteSelectionToContainingTable(readResult, out var promotedTable))
                {
                    readResult = promotedTable;
                    result.ReadResult = promotedTable;
                }

                // If the user clicks into an embedded object (e.g., Math) without selecting a range,
                // OneNote may mark that object as "selected", which is detected as Selection mode.
                // For TeXShift-managed outlines, prefer restoring the stored source for the entire outline.
                if (ReverseSelectionPromoter.TryPromoteSelectionToOutlineWithValidMeta(readResult, out var promoted))
                {
                    readResult = promoted;
                    result.ReadResult = promoted;
                }

                PerformanceTraceContext.AddMetric("DetectionMode", readResult.Mode.ToString());

                XElement toConvert;
                using (PerformanceTraceContext.Measure("Reverse.BuildInputFragment"))
                {
                    toConvert = BuildReverseInputFragment(readResult);
                }

                if (options.WriteDebugFiles)
                {
                    PerformanceTraceContext.AddMetric("InputFragmentOeCount", toConvert.Descendants().Count(e => e.Name.LocalName == "OE").ToString());

                    // Log what we actually feed into the reverse converter (Cursor: full Outline, Selection: an Outline fragment).
                    using (PerformanceTraceContext.Measure("Debug.Write.OriginalXml"))
                    {
                        await logger.LogOriginalXmlAsync(toConvert).ConfigureAwait(false);
                    }

                    // Also log selection nodes for diagnosing OneNote's selection/cursor quirks.
                    if (readResult.OriginalXmlNode != null)
                    {
                        using (PerformanceTraceContext.Measure("Debug.Write.SelectionXml"))
                        {
                            await logger.LogSelectionXmlAsync(readResult.OriginalXmlNode).ConfigureAwait(false);
                        }
                    }
                    else if (readResult.OriginalXmlNodes != null && readResult.OriginalXmlNodes.Count > 0)
                    {
                        using (PerformanceTraceContext.Measure("Debug.Write.SelectionXml"))
                        {
                            await logger.LogSelectionXmlAsync(readResult.OriginalXmlNodes).ConfigureAwait(false);
                        }
                    }
                }

                var converter = _serviceContainer.CreateOneNoteToMarkdownConverter();
                ReverseConversionResult reverseResult;
                using (PerformanceTraceContext.Measure("Reverse.ConvertToMarkdownAsync"))
                {
                    reverseResult = await converter.ConvertToMarkdownAsync(toConvert).ConfigureAwait(false);
                }
                result.ReverseResult = reverseResult;

                string markdown = reverseResult?.Markdown ?? string.Empty;

                // When reversing a table selection inside a larger TeXShift-managed outline, the generic table renderer
                // normalizes the separator row (e.g., "-----" -> "---"). That is valid Markdown, but it breaks strict
                // round-trip expectations. If outline meta is valid, prefer the original table snippet from meta.
                if (TeXShiftMetaTableSnippetRestorer.TryRestoreSelectedTableFromTeXShiftMeta(readResult, markdown, out var restoredTableMarkdown))
                {
                    markdown = restoredTableMarkdown ?? string.Empty;
                    if (reverseResult != null)
                    {
                        reverseResult.Markdown = markdown;
                    }
                    PerformanceTraceContext.AddMetric("Reverse.MetaTableSnippetUsed", "True");
                }
                else
                {
                    PerformanceTraceContext.AddMetric("Reverse.MetaTableSnippetUsed", "False");
                }

                if (options.WriteDebugFiles)
                {
                    PerformanceTraceContext.AddMetric("ReversedMarkdownChars", markdown.Length.ToString());

                    using (PerformanceTraceContext.Measure("Debug.Write.ReversedMarkdown"))
                    {
                        await logger.LogReversedMarkdownAsync(markdown).ConfigureAwait(false);
                    }
                }

                // Write back as plain text (Markdown) by replacing the selection/cursor target.
                // IMPORTANT: Do not store the whole Markdown as a single OE with <br/> inside <one:T>.
                // OneNote will re-serialize those <br/> with extra newlines in CDATA, which breaks the next forward conversion
                // (tables/quotes/code fences rely on contiguous lines). Splitting into one OE per line avoids that.
                var ns = XNamespace.Get("http://schemas.microsoft.com/office/onenote/2013/onenote");
                XElement outline;
                using (PerformanceTraceContext.Measure("Reverse.BuildWriteBackOutline"))
                {
                    outline = BuildPlainMarkdownOutline(ns, markdown ?? string.Empty);
                }
                if (options.WriteDebugFiles)
                {
                    using (PerformanceTraceContext.Measure("Debug.Write.WriteBackXml"))
                    {
                        await logger.LogConvertedXmlAsync(outline).ConfigureAwait(false);
                    }
                }

                var writer = _serviceContainer.CreateContentWriter(_oneNoteApp);
                using (PerformanceTraceContext.Measure("Write.ReplaceContentAsync"))
                {
                    await writer.ReplaceContentAsync(readResult, outline).ConfigureAwait(false);
                }

                await LogFinalPageXmlAsync(options, logger, readResult.PageId).ConfigureAwait(false);

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
            ReverseConversionOptions options,
            ReverseConversionPipelineResult result,
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
            logger.StartSession(DebugSessionKind.ReverseConversion);
            result.DebugOutputFolder = logger.DebugSessionFolder;

            perfTrace = new PerformanceTrace();
            perfSession = PerformanceTraceContext.BeginSession(perfTrace);
            PerformanceTraceContext.AddMetric("SessionKind", "ReverseConversion");
            PerformanceTraceContext.AddMetric("DumpFullPageXml", options.DumpFullPageXml.ToString());
        }

        private async Task LogFinalPageXmlAsync(ReverseConversionOptions options, IDebugLogger logger, string pageId)
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

        private static async Task TryLogPerformanceAsync(ReverseConversionOptions options, IDebugLogger logger, PerformanceTrace perfTrace)
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

        private static XElement BuildPlainMarkdownOutline(XNamespace ns, string markdown)
        {
            // Normalize line endings so we can preserve blank lines accurately.
            string normalized = (markdown ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace("\r", "\n");

            var lines = normalized.Split(new[] { '\n' }, StringSplitOptions.None);
            var oeChildren = new XElement(ns + "OEChildren");

            foreach (var line in lines)
            {
                oeChildren.Add(new XElement(ns + "OE",
                    new XElement(ns + "T", new XCData(OneNoteHtmlTextEscaper.Escape(line ?? string.Empty)))));
            }

            if (!oeChildren.Elements().Any())
            {
                oeChildren.Add(new XElement(ns + "OE",
                    new XElement(ns + "T", new XCData(string.Empty))));
            }

            return new XElement(ns + "Outline", oeChildren);
        }

        private static XElement BuildReverseInputFragment(ReadResult readResult)
        {
            if (readResult == null)
            {
                throw new ArgumentNullException(nameof(readResult));
            }

            if (readResult.Mode == DetectionMode.Cursor)
            {
                return readResult.OriginalXmlNode; // Outline
            }

            if (readResult.OriginalXmlNodes == null || readResult.OriginalXmlNodes.Count == 0)
            {
                throw new InvalidOperationException("Reverse conversion: selection mode requires OriginalXmlNodes.");
            }

            var ns = readResult.OriginalXmlNodes.First().Name.Namespace;
            return new XElement(ns + "Outline",
                new XElement(ns + "OEChildren", readResult.OriginalXmlNodes.Select(n => new XElement(n))));
        }
    }
}
