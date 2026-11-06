using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Xml.Linq;
using TeXShift.Core.Errors;
using TeXShift.Core.Localization;
using TeXShift.Core.Logging;
using TeXShift.Core.OneNote;
using OneNoteInterop = Microsoft.Office.Interop.OneNote;

namespace TeXShift.Core.Services
{
    public sealed class PasteConversionOrchestrator
    {
        private readonly ServiceContainer _services;
        private readonly OneNoteComDispatcher _dispatcher;

        internal PasteConversionOrchestrator(ServiceContainer services, OneNoteComDispatcher dispatcher)
        {
            _services = services;
            _dispatcher = dispatcher;
        }

        public async Task<ConversionResult> ExecuteAsync(string markdown)
        {
            var result = new ConversionResult();
            try
            {
                if (string.IsNullOrWhiteSpace(markdown)) throw new ArgumentException("Markdown is required.", nameof(markdown));
                PerformanceTraceContext.AddMetric("Paste.MarkdownChars", markdown.Length.ToString());
                // Capture the target immediately while rendering the independent clipboard input.
                var readPage = ReadPageAsync();
                var convert = ConvertAsync(markdown);
                await Task.WhenAll(readPage, convert).ConfigureAwait(false);
                var page = readPage.Result;
                result.ReadResult = new ReadResult { PageId = (string)page.Attribute("ID"), IsSuccess = true };
                using (PerformanceTraceContext.Measure("Write.InsertContentAsync"))
                    await new OneNotePageWriter(_dispatcher).InsertContentAsync(page, convert.Result).ConfigureAwait(false);
                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Error = ex;
                RuntimeLog.Write($"Paste conversion failed. {ex}");
            }
            return result;
        }

        private async Task<XElement> ConvertAsync(string markdown)
        {
            using (PerformanceTraceContext.Measure("Convert.MarkdownToOneNoteXmlAsync"))
                return await _services.CreateMarkdownConverter().ConvertToOneNoteXmlAsync(markdown).ConfigureAwait(false);
        }

        private async Task<XElement> ReadPageAsync()
        {
            using (PerformanceTraceContext.Measure("Paste.ReadPage"))
                return await _dispatcher.InvokeAsync(app =>
                {
                    OneNoteInterop.Windows windows = null;
                    OneNoteInterop.Window window = null;
                    try
                    {
                        string pageId;
                        using (PerformanceTraceContext.Measure("Read.OneNote.GetCurrentPageId"))
                        {
                            windows = app.Windows;
                            window = windows.CurrentWindow;
                            pageId = window?.CurrentPageId;
                        }
                        if (string.IsNullOrEmpty(pageId))
                            throw new ContentReadException(Resources.GetString("Error_NoPageId"), "No current page for paste.");
                        // Ink-heavy pages are expensive to export with binary data.
                        string xml;
                        using (PerformanceTraceContext.Measure("Read.OneNote.GetPageContent", "piSelection"))
                            app.GetPageContent(pageId, out xml, OneNoteInterop.PageInfo.piSelection, OneNoteInterop.XMLSchema.xs2013);
                        PerformanceTraceContext.AddMetric("Read.PageXmlChars", xml.Length.ToString());
                        using (PerformanceTraceContext.Measure("Read.ParseXmlContent"))
                            return XElement.Parse(xml);
                    }
                    finally
                    {
                        if (window != null) Marshal.ReleaseComObject(window);
                        if (windows != null) Marshal.ReleaseComObject(windows);
                    }
                }).ConfigureAwait(false);
        }
    }
}
