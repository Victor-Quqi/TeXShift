using System;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Core;
using TeXShift.AddIn.Localization;
using TeXShift.Core.Errors;

namespace TeXShift.AddIn
{
    public partial class Connect
    {
        private bool _pasteConversionInProgress;

        public async void OnPasteConvertButtonClick(IRibbonControl control)
        {
            if (!GetPasteConvertEnabled(control) || _pasteConversionInProgress) return;

            using (var activity = new AddInCommandLog("paste-convert"))
            {
                try
                {
                    if (TryShowInitializationError()) return;
                    _pasteConversionInProgress = true;
                    string markdown = ReadClipboardText();
                    if (string.IsNullOrWhiteSpace(markdown))
                    {
                        throw new ContentReadException(UIResources.GetString("Message_PasteConvert_EmptyClipboard"),
                            "Clipboard contains no Markdown text.");
                    }

                    var result = await _serviceContainer.CreatePasteConversionOrchestrator(_oneNoteApp).ExecuteAsync(markdown);
                    activity.PageId = result.ReadResult?.PageId;
                    activity.Success = result.Success;
                    if (result.Error != null) HandleConversionError(result.Error, null);
                }
                catch (Exception ex)
                {
                    activity.Success = false;
                    HandleConversionError(ex, null);
                }
                finally
                {
                    _pasteConversionInProgress = false;
                }
            }
        }

        private static string ReadClipboardText()
        {
            string text = null;
            Exception error = null;
            // COM Ribbon callbacks can arrive on MTA threads; OLE clipboard access needs STA.
            var thread = new Thread(() =>
            {
                try { text = Clipboard.GetText(TextDataFormat.UnicodeText); }
                catch (Exception ex) { error = ex; }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (error != null) throw error;
            return text;
        }
    }
}
