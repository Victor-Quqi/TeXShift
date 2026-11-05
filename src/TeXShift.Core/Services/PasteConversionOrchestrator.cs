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
                var page = await _dispatcher.InvokeAsync(app =>
                {
                    OneNoteInterop.Windows windows = null;
                    OneNoteInterop.Window window = null;
                    try
                    {
                        windows = app.Windows;
                        window = windows.CurrentWindow;
                        string pageId = window?.CurrentPageId;
                        if (string.IsNullOrEmpty(pageId))
                            throw new ContentReadException(Resources.GetString("Error_NoPageId"), "No current page for paste.");
                        // Ink-heavy pages are expensive to export with binary data.
                        app.GetPageContent(pageId, out string xml, OneNoteInterop.PageInfo.piSelection, OneNoteInterop.XMLSchema.xs2013);
                        return XElement.Parse(xml);
                    }
                    finally
                    {
                        if (window != null) Marshal.ReleaseComObject(window);
                        if (windows != null) Marshal.ReleaseComObject(windows);
                    }
                }).ConfigureAwait(false);
                result.ReadResult = new ReadResult { PageId = (string)page.Attribute("ID"), IsSuccess = true };
                var converted = await _services.CreateMarkdownConverter().ConvertToOneNoteXmlAsync(markdown).ConfigureAwait(false);
                await new OneNotePageWriter(_dispatcher).InsertContentAsync(page, converted).ConfigureAwait(false);
                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Error = ex;
                RuntimeLog.Write($"Paste conversion failed. {ex}");
            }
            return result;
        }
    }
}
