using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Interop;
using Microsoft.Office.Core;
using TeXShift.AddIn.Interop;
using TeXShift.AddIn.Localization;
using TeXShift.AddIn.UI.WPF;
using TeXShift.AddIn.UI.WPF.ViewModels;
using TeXShift.Core.Errors;
using TeXShift.Core.Localization;
using TeXShift.Core.Logging;
using TeXShift.Core.OneNote;
using TeXShift.Core.Services;

namespace TeXShift.AddIn
{
    /// <summary>
    /// Partial class containing Markdown conversion logic.
    /// </summary>
    public partial class Connect
    {
        #region Conversion Button Handlers

        /// <summary>
        /// Ribbon button click handler for conversion.
        /// Uses async void pattern for event handlers.
        /// </summary>
        public async void OnConvertButtonClick(IRibbonControl control)
        {
            try
            {
                await PerformConversionAsync(showSuccessDialog: false, writeDebugFiles: false);
            }
            catch (Exception ex)
            {
                await HandleConversionErrorAsync(ex, null);
            }
        }

        public async void OnDebugConvertButtonClick(IRibbonControl control)
        {
            try
            {
                await PerformConversionAsync(showSuccessDialog: true, writeDebugFiles: true);
            }
            catch (Exception ex)
            {
                await HandleConversionErrorAsync(ex, null);
            }
        }

        #endregion

        #region Conversion Logic

        /// <summary>
        /// Core conversion logic. Reads from OneNote, converts, and writes back.
        /// </summary>
        /// <param name="showSuccessDialog">If true, shows a detailed message box on success.</param>
        /// <param name="writeDebugFiles">If true, saves conversion artifacts to the DebugOutput folder.</param>
        private async Task PerformConversionAsync(bool showSuccessDialog, bool writeDebugFiles)
        {
            using (var activity = new AddInCommandLog("convert"))
            {
                try
                {
                    if (TryShowInitializationError()) return;
                    var orchestrator = _serviceContainer.CreateConversionOrchestrator(_oneNoteApp);
                    var result = await orchestrator.ExecuteAsync(new ConversionOptions
                    {
                        WriteDebugFiles = writeDebugFiles,
                        ExportPdf = writeDebugFiles && (_appSettings?.Debug?.ExportPdf ?? true),
                        DumpFullPageXml = writeDebugFiles && (_appSettings?.Debug?.DumpFullPageXml ?? false),
                        OutputDirectory = _appSettings?.Debug?.DebugOutputPath
                    });

                    activity.PageId = result.ReadResult?.PageId;
                    activity.Success = result.Success;

                    if (!result.Success)
                    {
                        if (result.Error != null)
                        {
                            await HandleConversionErrorAsync(result.Error, result.DebugOutputFolder);
                        }
                        else if (result.ReadResult != null && !result.ReadResult.IsSuccess)
                        {
                            // Show read errors (like "no selection") using the new error dialog
                            await ShowWpfErrorDialogAsync(
                                result.ReadResult.ErrorMessage,
                                string.Empty,
                                ResolveDebugFolderPath(result.DebugOutputFolder));
                        }
                        return;
                    }

                    if (showSuccessDialog)
                    {
                        ShowSuccessMessageFromResult(result);
                    }
                }
                catch (Exception ex)
                {
                    activity.Success = false;
                    await HandleConversionErrorAsync(ex, null);
                }
            }
        }

        #endregion

        #region Result Handling

        /// <summary>
        /// Shows a success message box with conversion details.
        /// </summary>
        private void ShowSuccessMessageFromResult(ConversionResult result)
        {
            ShowTopMostMessageBox(
                string.Format(
                    UIResources.GetString("Message_Success_ConversionComplete"),
                    GetDetectionModeLabel(result.ReadResult?.Mode ?? DetectionMode.None),
                    result.ReadResult?.ExtractedText?.Length ?? 0,
                    result.DebugOutputFolder),
                UIResources.GetString("Message_Title_ConversionComplete"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }

        private static string GetDetectionModeLabel(DetectionMode mode)
        {
            switch (mode)
            {
                case DetectionMode.Cursor:
                    return Resources.GetString("Mode_Cursor");
                case DetectionMode.Selection:
                    return Resources.GetString("Mode_Selection");
                default:
                    return Resources.GetString("Mode_Unknown");
            }
        }

        #endregion

        #region Error Handling

        /// <summary>
        /// Handles and displays conversion errors.
        /// </summary>
        private async Task HandleConversionErrorAsync(Exception ex, string debugFolderPath)
        {
            var userMessage = ErrorMessages.GetUserFriendlyMessage(ex);
            var resolvedDebugFolder = ResolveDebugFolderPath(debugFolderPath);
            var technicalDetails = BuildTechnicalDetails(ex, resolvedDebugFolder);

            try
            {
                await ShowWpfErrorDialogAsync(userMessage, technicalDetails, resolvedDebugFolder);
            }
            catch
            {
                ShowTopMostMessageBox(
                    $"{userMessage}\n\n{technicalDetails}",
                    Resources.GetString("Dialog_ErrorTitle"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private string ResolveDebugFolderPath(string debugFolderPath)
        {
            if (!string.IsNullOrWhiteSpace(debugFolderPath))
            {
                return debugFolderPath;
            }

            try
            {
                return DebugLogger.ResolveDebugOutputFolder(_appSettings?.Debug?.DebugOutputPath);
            }
            catch
            {
                // Folder resolution may fail (e.g., UnauthorizedAccessException in Program Files).
                // Return null so error dialogs can still be shown without cascading failures.
                return null;
            }
        }

        private string BuildTechnicalDetails(Exception ex, string debugFolderPath)
        {
            if (ex == null)
            {
                return string.Empty;
            }

            var builder = new StringBuilder();

            if (ex is TeXShiftException texShiftException)
            {
                builder.AppendLine($"{Resources.GetString("Dialog_ErrorCodeLabel")}: {texShiftException.ErrorCode}");
                if (!string.IsNullOrWhiteSpace(texShiftException.UserMessage))
                {
                    builder.AppendLine($"{Resources.GetString("Dialog_UserMessageLabel")}: {texShiftException.UserMessage}");
                }
            }

            if (ex is COMException comEx)
            {
                builder.AppendLine($"{Resources.GetString("Dialog_ComHResultLabel")}: 0x{comEx.HResult:X}");
            }
            else if (ex.HResult != 0)
            {
                builder.AppendLine($"{Resources.GetString("Dialog_HResultLabel")}: 0x{ex.HResult:X}");
            }

            builder.AppendLine();
            builder.AppendLine(ex.ToString());

            if (!string.IsNullOrWhiteSpace(debugFolderPath))
            {
                builder.AppendLine();
                builder.AppendLine($"{Resources.GetString("Dialog_DebugFolderLabel")}: {debugFolderPath}");
            }

            return builder.ToString();
        }

        private async Task ShowWpfErrorDialogAsync(string userMessage, string technicalDetails, string debugFolderPath)
        {
            IntPtr parentHwnd = IntPtr.Zero;
            try
            {
                // Release the Ribbon callback while OneNote services the COM request.
                using (PerformanceTraceContext.Measure("Error.ResolveOwner"))
                    parentHwnd = await _serviceContainer.GetOneNoteWindowHandleAsync(_oneNoteApp).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                // OneNote itself may have disconnected. Keep the error visible as an unowned window.
                RuntimeLog.Write($"Could not resolve the OneNote error dialog owner: {ex}");
            }
            var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

            var thread = new System.Threading.Thread(() =>
            {
                Exception threadException = null;
                IDisposable presentation = PerformanceTraceContext.Measure("Error.ShowDialog");
                IDisposable dismissal = null;
                try
                {
                    var viewModel = new ErrorDialogViewModel(userMessage, technicalDetails, debugFolderPath);
                    var dialog = new ErrorDialog(viewModel);

                    var helper = new WindowInteropHelper(dialog);
                    helper.Owner = parentHwnd;

                    dialog.ContentRendered += (sender, args) =>
                    {
                        presentation?.Dispose();
                        presentation = null;
                        dismissal = dismissal ?? PerformanceTraceContext.Measure("Error.WaitForDismissal");
                    };

                    dialog.ShowDialog();
                }
                catch (Exception ex)
                {
                    RuntimeLog.Write($"Error dialog failed: {ex}");
                    threadException = ex;
                }
                finally
                {
                    dismissal?.Dispose();
                    presentation?.Dispose();
                    var dispatcher = System.Windows.Threading.Dispatcher.CurrentDispatcher;
                    dispatcher.InvokeShutdown();
                }
                if (parentHwnd != IntPtr.Zero) NativeMethods.SetForegroundWindow(parentHwnd);
                if (threadException != null) completion.TrySetException(threadException);
                else completion.TrySetResult(true);
            }) { IsBackground = true, Name = "TeXShift_ErrorDialog_STA" };

            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            await completion.Task.ConfigureAwait(false);
        }

        #endregion
    }
}
