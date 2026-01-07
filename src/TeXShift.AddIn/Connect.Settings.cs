using System;
using System.Windows.Forms;
using System.Windows.Interop;
using Microsoft.Office.Core;
using TeXShift.AddIn.Localization;
using TeXShift.AddIn.UI;
using TeXShift.AddIn.UI.WPF;
using TeXShift.Core.Localization;

namespace TeXShift.AddIn
{
    /// <summary>
    /// Partial class containing settings dialog handling.
    /// </summary>
    public partial class Connect
    {
        #region Settings Button Handler

        /// <summary>
        /// Settings button click handler. Opens the settings dialog.
        /// </summary>
        public void OnSettingsButtonClick(IRibbonControl control)
        {
            try
            {
                ShowWpfSettingsDialog();
            }
            catch (Exception ex)
            {
                // Temporarily show WPF error for debugging
                ShowTopMostMessageBox(
                    string.Format(
                        UIResources.GetString("Message_Error_WpfFallback"),
                        ex.Message,
                        ex.StackTrace ?? string.Empty),
                    UIResources.GetString("Message_Title_WpfError"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                try
                {
                    ShowWinFormsSettingsDialog();
                }
                catch (Exception ex2)
                {
                    ShowTopMostMessageBox(
                        string.Format(UIResources.GetString("Message_Error_OpenSettings"), ex2.Message),
                        Resources.GetString("Dialog_ErrorTitle"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        #endregion

        #region Settings Dialogs

        /// <summary>
        /// Shows the WPF Material Design settings dialog.
        /// </summary>
        private void ShowWpfSettingsDialog()
        {
            var parentHwnd = GetForegroundWindow();
            Core.Configuration.AppSettings updatedSettings = null;
            bool dialogResult = false;
            Exception threadException = null;

            // WPF requires STA thread
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    var dialog = new SettingsWindow(_appSettings);

                    // Set WPF window's Owner to OneNote window
                    var helper = new WindowInteropHelper(dialog);
                    helper.Owner = parentHwnd;

                    if (dialog.ShowDialog() == true)
                    {
                        updatedSettings = dialog.GetUpdatedSettings();
                        dialogResult = true;
                    }
                }
                catch (Exception ex)
                {
                    threadException = ex;
                }
                finally
                {
                    // Ensure WPF Dispatcher is closed correctly to release resources
                    var dispatcher = System.Windows.Threading.Dispatcher.CurrentDispatcher;
                    dispatcher.InvokeShutdown();
                }
            });

            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();

            // Force GC to clean up WPF resources
            GC.Collect();
            GC.WaitForPendingFinalizers();

            // Restore focus to OneNote window
            SetForegroundWindow(parentHwnd);

            if (threadException != null)
                throw threadException;

            if (dialogResult && updatedSettings != null)
            {
                _appSettings = updatedSettings;
                _settingsManager.Save(_appSettings);
                LocalizationManager.Initialize(_appSettings?.Language);
                ApplySettingsToStyleConfig();

                // Refresh Ribbon to update button visibility
                _ribbon?.Invalidate();
            }
        }

        /// <summary>
        /// Fallback: Shows the WinForms settings dialog.
        /// </summary>
        private void ShowWinFormsSettingsDialog()
        {
            var owner = new Win32Window(GetForegroundWindow());
            using (var dialog = new SettingsDialog(_appSettings))
            {
                if (dialog.ShowDialog(owner) == DialogResult.OK)
                {
                    _appSettings = dialog.GetUpdatedSettings();
                    _settingsManager.Save(_appSettings);
                    LocalizationManager.Initialize(_appSettings?.Language);
                    ApplySettingsToStyleConfig();

                    // Refresh Ribbon to update button visibility
                    _ribbon?.Invalidate();
                }
            }
        }

        #endregion

        #region Settings Application

        /// <summary>
        /// Applies the current AppSettings to the OneNoteStyleConfig singleton.
        /// </summary>
        private void ApplySettingsToStyleConfig()
        {
            if (_appSettings == null || _serviceContainer == null)
                return;

            var styleConfig = _serviceContainer.StyleConfig;

            // Apply code block settings
            var codeBlock = _appSettings.CodeBlock;
            styleConfig.SetCodeBlockStyle(
                codeBlock.BackgroundColor,
                codeBlock.TextColor,
                codeBlock.FontFamily,
                codeBlock.FontSize,
                codeBlock.SpaceBetween,
                codeBlock.EnableSyntaxHighlight);

            // Apply inline code settings
            var inlineCode = _appSettings.InlineCode;
            styleConfig.SetInlineCodeStyle(
                inlineCode.FontFamily,
                inlineCode.BackgroundColor);

            // Apply quote block settings
            var quoteBlock = _appSettings.QuoteBlock;
            styleConfig.SetQuoteBlockStyle(quoteBlock.BackgroundColor);

            // Apply heading settings
            var headings = _appSettings.Headings;
            for (int i = 1; i <= 6; i++)
            {
                styleConfig.SetHeadingFont(i, headings.GetFontSize(i));
            }
        }

        #endregion
    }
}
