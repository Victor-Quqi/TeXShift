using System;
using System.Windows.Forms;
using Microsoft.Office.Core;
using TeXShift.AddIn.Localization;
using TeXShift.Core.OneNote;
using TeXShift.Core.Services;

namespace TeXShift.AddIn
{
    /// <summary>
    /// Partial class containing OneNote -> Markdown reverse conversion logic.
    /// </summary>
    public partial class Connect
    {
        public void OnReverseConvertButtonClick(IRibbonControl control)
        {
            try
            {
                if (TryShowInitializationError())
                {
                    return;
                }
                PerformReverseConversionAsync(showSuccessDialog: false, writeDebugFiles: false);
            }
            catch (Exception ex)
            {
                HandleConversionError(ex, null);
            }
        }

        public void OnDebugReverseConvertButtonClick(IRibbonControl control)
        {
            try
            {
                if (TryShowInitializationError())
                {
                    return;
                }
                PerformReverseConversionAsync(showSuccessDialog: true, writeDebugFiles: true);
            }
            catch (Exception ex)
            {
                HandleConversionError(ex, null);
            }
        }

        private async void PerformReverseConversionAsync(bool showSuccessDialog, bool writeDebugFiles)
        {
            try
            {
                var orchestrator = _serviceContainer.CreateReverseConversionOrchestrator(_oneNoteApp);
                var result = await orchestrator.ExecuteAsync(new ReverseConversionOptions
                {
                    WriteDebugFiles = writeDebugFiles,
                    DumpFullPageXml = writeDebugFiles && (_appSettings?.Debug?.DumpFullPageXml ?? false),
                    OutputDirectory = _appSettings?.Debug?.DebugOutputPath
                });

                if (!result.Success)
                {
                    if (result.Error != null)
                    {
                        HandleConversionError(result.Error, result.DebugOutputFolder);
                    }
                    else if (result.ReadResult != null && !result.ReadResult.IsSuccess)
                    {
                        ShowWpfErrorDialog(
                            result.ReadResult.ErrorMessage,
                            string.Empty,
                            ResolveDebugFolderPath(result.DebugOutputFolder));
                    }
                    return;
                }

                if (showSuccessDialog)
                {
                    ShowTopMostMessageBox(
                        string.Format(
                            UIResources.GetString("Message_Success_ReverseConversionComplete"),
                            GetDetectionModeLabel(result.ReadResult?.Mode ?? DetectionMode.None),
                            result.ReverseResult?.Markdown?.Length ?? 0,
                            result.DebugOutputFolder),
                        UIResources.GetString("Message_Title_ReverseConversionComplete"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                HandleConversionError(ex, null);
            }
        }
    }
}
