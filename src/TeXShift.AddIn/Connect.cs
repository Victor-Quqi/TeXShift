using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Interop;
using Extensibility;
using Microsoft.Office.Core;
using TeXShift.Core.Configuration;
using TeXShift.Core.Errors;
using TeXShift.Core.Localization;
using TeXShift.Core.Logging;
using TeXShift.Core.OneNote;
using TeXShift.Core.Services;
using TeXShift.AddIn.UI;
using TeXShift.AddIn.UI.WPF;
using TeXShift.AddIn.UI.WPF.ViewModels;
using OneNote = Microsoft.Office.Interop.OneNote;

 namespace TeXShift.AddIn
 {
     /// <summary>
     /// Helper class to wrap a window handle for use with WinForms dialogs.
     /// </summary>
     internal class Win32Window : System.Windows.Forms.IWin32Window
     {
         public IntPtr Handle { get; }
         public Win32Window(IntPtr handle) => Handle = handle;
     }

     /// <summary>
     /// OneNote COM Add-in entry point.
     /// </summary>
     [ComVisible(true)]
     [Guid("1EE8F914-ECBD-4709-92C0-E770851C4966")]
     [ProgId("TeXShift.AddIn.Connect")]
     public class Connect : IDTExtensibility2, IRibbonExtensibility
     {
         [DllImport("user32.dll")]
         private static extern IntPtr GetForegroundWindow();

         [DllImport("user32.dll")]
         private static extern bool EnableWindow(IntPtr hWnd, bool bEnable);

         [DllImport("user32.dll")]
         private static extern bool SetForegroundWindow(IntPtr hWnd);

         private static readonly string _addInDirectory;

         /// <summary>
         /// Static constructor to set up assembly resolution for COM Add-in.
         /// </summary>
         static Connect()
         {
             _addInDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
             AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;
         }

         /// <summary>
         /// Handles assembly resolution for dependencies that can't be found in the default probe paths.
         /// </summary>
         private static Assembly OnAssemblyResolve(object sender, ResolveEventArgs args)
         {
             var assemblyName = new AssemblyName(args.Name);
             var assemblyPath = Path.Combine(_addInDirectory, assemblyName.Name + ".dll");

             if (File.Exists(assemblyPath))
             {
                 return Assembly.LoadFrom(assemblyPath);
             }

             return null;
         }

         private OneNote.Application _oneNoteApp;
         private IRibbonUI ribbon;
         private ServiceContainer _serviceContainer;
         private AppSettings _appSettings;
         private SettingsManager _settingsManager;
 
         /// <summary>
         /// Called when the add-in is connected to OneNote.
         /// </summary>
         public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
         {
             _oneNoteApp = (OneNote.Application)Application;

             // Load settings from JSON file
            _settingsManager = new SettingsManager();
            _appSettings = _settingsManager.Load();

            // Initialize localization based on settings or system culture
            LocalizationManager.Initialize(_appSettings?.Language);

            // Initialize dependency injection container
            _serviceContainer = new ServiceContainer();

             // Apply loaded settings to style configuration
             ApplySettingsToStyleConfig();
         }
 
         /// <summary>
         /// Called when the add-in is disconnected from OneNote.
         /// </summary>
         public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
         {
             // Dispose ServiceContainer to clean up WebView2 and other resources
             _serviceContainer?.Dispose();
             _serviceContainer = null;

             // Explicitly release the COM object to ensure OneNote can shut down cleanly.
             SafeReleaseComObject(_oneNoteApp);
 
             // While explicit release is key, garbage collection can help clean up any other managed wrappers.
             GC.Collect();
             GC.WaitForPendingFinalizers();
             GC.Collect();
             GC.WaitForPendingFinalizers();
         }

        /// <summary>
        /// Called when the add-in is loaded on startup.
        /// </summary>
        public void OnStartupComplete(ref Array custom)
        {
        }

        /// <summary>
        /// Called when OneNote is shutting down.
        /// </summary>
        public void OnBeginShutdown(ref Array custom)
        {
        }

        /// <summary>
        /// Called when add-ins are updated.
        /// </summary>
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        /// <summary>
        /// Returns the XML for the custom Ribbon UI.
        /// </summary>
        public string GetCustomUI(string RibbonID)
        {
            return GetResourceText("TeXShift.AddIn.Ribbon.xml");
        }

        public void OnRibbonLoad(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #region Ribbon Visibility Callbacks

        /// <summary>
        /// Ribbon callback: Returns whether the debug tools group should be visible.
        /// </summary>
        public bool GetDebugGroupVisible(IRibbonControl control)
        {
            return _appSettings?.Debug?.ShowDebugButtons ?? false;
        }

        #endregion

        #region Settings

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
                    $"WPF 对话框失败，回退到 WinForms:\n\n{ex.Message}\n\n{ex.StackTrace}",
                    "TeXShift - WPF 错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                try
                {
                    ShowWinFormsSettingsDialog();
                }
                catch (Exception ex2)
                {
                    ShowTopMostMessageBox(
                        $"打开设置失败:\n\n{ex2.Message}",
                        "TeXShift - 错误",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// Shows the WPF Material Design settings dialog.
        /// </summary>
        private void ShowWpfSettingsDialog()
        {
            var parentHwnd = GetForegroundWindow();
            AppSettings updatedSettings = null;
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
                ribbon?.Invalidate();
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
                    ribbon?.Invalidate();
                }
            }
        }

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

        /// <summary>
        /// Ribbon button click handler for conversion.
        /// Uses async void pattern for event handlers.
        /// </summary>
        public void OnConvertButtonClick(IRibbonControl control)
        {
            // This is the new "Silent Convert" button.
            // It does not show a success message box to avoid interrupting the user's workflow.
            // It does NOT write debug files.
            // Errors will still be displayed.
            PerformConversionAsync(showSuccessDialog: false, writeDebugFiles: false);
        }

        public void OnDebugConvertButtonClick(IRibbonControl control)
        {
            // This is the original "Convert" button, now repurposed for debugging.
            // It shows detailed success information and saves debug files.
            PerformConversionAsync(showSuccessDialog: true, writeDebugFiles: true);
        }

        /// <summary>
        /// Core conversion logic. Reads from OneNote, converts, and writes back.
        /// </summary>
        /// <param name="showSuccessDialog">If true, shows a detailed message box on success.</param>
        /// <param name="writeDebugFiles">If true, saves conversion artifacts to the DebugOutput folder.</param>
        private async void PerformConversionAsync(bool showSuccessDialog, bool writeDebugFiles)
        {
            try
            {
                var orchestrator = _serviceContainer.CreateConversionOrchestrator(_oneNoteApp);
                var result = await orchestrator.ExecuteAsync(new ConversionOptions
                {
                    WriteDebugFiles = writeDebugFiles,
                    ExportPdf = writeDebugFiles && (_appSettings?.Debug?.ExportPdf ?? true),
                    OutputDirectory = _appSettings?.Debug?.DebugOutputPath
                });

                if (!result.Success)
                {
                    if (result.ReadResult != null && !result.ReadResult.IsSuccess)
                    {
                        ShowTopMostMessageBox(result.ReadResult.ErrorMessage, "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        HandleConversionError(result.Error, result.DebugOutputFolder);
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
                HandleConversionError(ex, null);
            }
        }

        /// <summary>
        /// Shows a success message box with conversion details.
        /// </summary>
        private void ShowSuccessMessageFromResult(ConversionResult result)
        {
            ShowTopMostMessageBox(
                $"转换成功!\n\n" +
                $"模式: {result.ReadResult?.ModeAsString()}\n" +
                $"处理了 {result.ReadResult?.ExtractedText?.Length ?? 0} 个字符\n\n" +
                $"调试文件已保存至:\n{result.DebugOutputFolder}",
                "TeXShift - 转换完成",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }

        /// <summary>
        /// Handles and displays conversion errors.
        /// </summary>
        private void HandleConversionError(Exception ex, string debugFolderPath)
        {
            var userMessage = ErrorMessages.GetUserFriendlyMessage(ex);
            var resolvedDebugFolder = ResolveDebugFolderPath(debugFolderPath);
            var technicalDetails = BuildTechnicalDetails(ex, resolvedDebugFolder);

            try
            {
                ShowWpfErrorDialog(userMessage, technicalDetails, resolvedDebugFolder);
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

            return DebugLogger.ResolveDebugOutputFolder(_appSettings?.Debug?.DebugOutputPath);
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

        private void ShowWpfErrorDialog(string userMessage, string technicalDetails, string debugFolderPath)
        {
            var parentHwnd = GetForegroundWindow();
            Exception threadException = null;

            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    var viewModel = new ErrorDialogViewModel(userMessage, technicalDetails, debugFolderPath);
                    var dialog = new ErrorDialog(viewModel);

                    var helper = new WindowInteropHelper(dialog);
                    helper.Owner = parentHwnd;

                    dialog.ShowDialog();
                }
                catch (Exception ex)
                {
                    threadException = ex;
                }
                finally
                {
                    var dispatcher = System.Windows.Threading.Dispatcher.CurrentDispatcher;
                    dispatcher.InvokeShutdown();
                }
            });

            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            SetForegroundWindow(parentHwnd);

            if (threadException != null)
            {
                throw threadException;
            }
        }

        /// <summary>
        /// Safely releases a COM object and sets its reference to null.
        /// </summary>
        /// <param name="obj">The COM object to release.</param>
        private void SafeReleaseComObject(object obj)
        {
            if (obj != null)
            {
                try
                {
                    Marshal.ReleaseComObject(obj);
                }
                catch (Exception ex)
                {
                    // Log the exception but don't throw - object might already be released
                    System.Diagnostics.Debug.WriteLine($"Warning: Failed to release COM object: {ex.Message}");
                }
            }
        }
 
        /// <summary>
        /// Debug button: Shows and saves the selected content's XML structure only.
        /// </summary>
        public async void OnDebugSelectionXmlButtonClick(IRibbonControl control)
        {
            try
            {
                var reader = _serviceContainer.CreateContentReader(_oneNoteApp);
                var result = await reader.ExtractContentAsync();

                if (!result.IsSuccess)
                {
                    ShowTopMostMessageBox(result.ErrorMessage, "调试工具", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (result.OriginalXmlNode == null)
                {
                    ShowTopMostMessageBox("未能获取选中内容的XML节点。", "调试工具", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var logger = _serviceContainer.CreateDebugLogger(_appSettings?.Debug?.DebugOutputPath);
                logger.StartSession();
                string savedPath = await logger.LogSelectionXmlAsync(result.OriginalXmlNode);
                string formattedXml = System.Xml.Linq.XDocument.Parse(result.OriginalXmlNode.ToString()).ToString();


                // Show in dialog
                string caption = $"选中内容 XML 结构 - {result.ModeAsString()} (已保存至: {Path.GetFileName(savedPath)})";
                ShowTextInScrollableMessageBox(formattedXml, caption);

                // Show success message
                ShowTopMostMessageBox(
                    $"选中内容的XML已保存至：\n{savedPath}\n\n" +
                    $"模式: {result.ModeAsString()}\n" +
                    $"节点类型: {result.OriginalXmlNode.Name.LocalName}\n" +
                    $"ObjectIDs: {string.Join(", ", result.TargetObjectIds)}",
                    "调试工具 - 保存成功",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowTopMostMessageBox("调试功能发生错误：\n" + ex.ToString(), "调试工具异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Debug button: Shows and saves the raw OneNote XML structure for entire page.
        /// </summary>
        public async void OnDebugXmlButtonClick(IRibbonControl control)
        {
            try
            {
                var (pageId, xmlContent) = await Task.Run(() =>
                {
                    OneNote.Windows windows = null;
                    OneNote.Window window = null;
                    try
                    {
                        windows = _oneNoteApp.Windows;
                        window = windows.CurrentWindow;
                        string id = window?.CurrentPageId;
                        if (string.IsNullOrEmpty(id)) return (null, null);
                        _oneNoteApp.GetPageContent(id, out string xml, OneNote.PageInfo.piAll);
                        return (id, xml);
                    }
                    finally
                    {
                        SafeReleaseComObject(window);
                        SafeReleaseComObject(windows);
                    }
                });

                if (string.IsNullOrEmpty(pageId) || string.IsNullOrEmpty(xmlContent))
                {
                    ShowTopMostMessageBox("无法获取当前页面内容。", "调试工具", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var logger = _serviceContainer.CreateDebugLogger(_appSettings?.Debug?.DebugOutputPath);
                logger.StartSession();
                string savedPath = await logger.LogPageXmlAsync(xmlContent);
                string formattedXml = System.Xml.Linq.XDocument.Parse(xmlContent).ToString();

                // Show in dialog
                string caption = $"OneNote XML 结构 (已保存至: {Path.GetFileName(savedPath)})";
                ShowTextInScrollableMessageBox(formattedXml, caption);

                // Show success message
                ShowTopMostMessageBox(
                    $"XML已保存至：\n{savedPath}\n\n文件大小: {new FileInfo(savedPath).Length / 1024.0:F2} KB",
                    "调试工具 - 保存成功",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowTopMostMessageBox("调试功能发生错误：\n" + ex.ToString(), "调试工具异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        /// <summary>
        /// Helper function: Shows a MessageBox that appears on top of all windows.
        /// </summary>
        private DialogResult ShowTopMostMessageBox(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            var owner = new Win32Window(GetForegroundWindow());
            return MessageBox.Show(owner, text, caption, buttons, icon);
        }

        /// <summary>
        /// Helper function: Creates a form with a scrollbar to display a large amount of text.
        /// </summary>
        private void ShowTextInScrollableMessageBox(string text, string caption)
        {
            var owner = new Win32Window(GetForegroundWindow());
            using (Form form = new Form
            {
                Text = caption,
                Size = new System.Drawing.Size(600, 400),
                StartPosition = FormStartPosition.CenterParent
            })
            {
                TextBox textBox = new TextBox
                {
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    Dock = DockStyle.Fill,
                    ReadOnly = true,
                    Font = new System.Drawing.Font("Consolas", 10),
                    Text = text
                };
                form.Controls.Add(textBox);
                form.ShowDialog(owner);
            }
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            foreach (string name in resourceNames)
            {
                if (string.Compare(resourceName, name, System.StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(name)))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
    }
}
