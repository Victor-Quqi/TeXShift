using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using TeXShift.Core;
using TeXShift.Core.Logging;
using OneNote = Microsoft.Office.Interop.OneNote;

 namespace TeXShift
 {
     /// <summary>
     /// OneNote COM Add-in entry point.
     /// </summary>
     [ComVisible(true)]
     [Guid("1EE8F914-ECBD-4709-92C0-E770851C4966")]
     [ProgId("TeXShift.Connect")]
     public class Connect : IDTExtensibility2, IRibbonExtensibility
     {
         private OneNote.Application _oneNoteApp;
         private IRibbonUI ribbon;
         private ServiceContainer _serviceContainer;
 
         /// <summary>
         /// Called when the add-in is connected to OneNote.
         /// </summary>
         public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
         {
             _oneNoteApp = (OneNote.Application)Application;

             // Initialize dependency injection container
             _serviceContainer = new ServiceContainer();
         }
 
         /// <summary>
         /// Called when the add-in is disconnected from OneNote.
         /// </summary>
         public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
         {
             // Explicitly release the COM object to ensure OneNote can shut down cleanly.
             SafeReleaseComObject(_oneNoteApp);
             _serviceContainer = null;
 
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
            return GetResourceText("TeXShift.Ribbon.xml");
        }

        public void OnRibbonLoad(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

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
            IDebugLogger logger = null;

            try
            {
                if (writeDebugFiles)
                {
                    logger = _serviceContainer.CreateDebugLogger();
                    logger.StartSession();
                }

                // Step 1: Read content from OneNote
                var reader = _serviceContainer.CreateContentReader(_oneNoteApp);
                var readResult = await reader.ExtractContentAsync();

                if (!readResult.IsSuccess)
                {
                    MessageBox.Show(readResult.ErrorMessage, "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (writeDebugFiles)
                {
                    await logger.LogInputMarkdownAsync(readResult.ExtractedText);
                    await logger.LogOriginalXmlAsync(readResult.OriginalXmlNode);
                }

                // Step 2: Convert Markdown to OneNote XML
                var converter = _serviceContainer.CreateMarkdownConverter(readResult.SourceOutlineWidth);
                var oneNoteXml = await converter.ConvertToOneNoteXmlAsync(readResult.ExtractedText);

                if (writeDebugFiles)
                {
                    await logger.LogConvertedXmlAsync(oneNoteXml);
                }

                // Step 3: Write back to OneNote
                var writer = _serviceContainer.CreateContentWriter(_oneNoteApp);
                await writer.ReplaceContentAsync(readResult, oneNoteXml);

                if (writeDebugFiles)
                {
                    string finalPageXml = await Task.Run(() =>
                    {
                        _oneNoteApp.GetPageContent(readResult.PageId, out string xml, OneNote.PageInfo.piAll);
                        return xml;
                    });
                    await logger.LogFinalPageXmlAsync(finalPageXml);
                }

                if (showSuccessDialog)
                {
                    MessageBox.Show(
                        $"转换成功!\n\n" +
                        $"模式: {readResult.ModeAsString()}\n" +
                        $"处理了 {readResult.ExtractedText.Length} 个字符\n\n" +
                        $"调试文件已保存至:\n{logger?.DebugSessionFolder}",
                        "TeXShift - 转换完成",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
            catch (Exception ex)
            {
                // Use fire-and-forget for logging to avoid secondary exceptions blocking the UI
                _ = logger?.LogErrorAsync(ex);

                MessageBox.Show(
                    $"转换失败:\n\n{ex.Message}\n\n详细信息:\n{ex.StackTrace}",
                    "TeXShift - 错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
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
                catch (Exception)
                {
                    // Ignore exceptions during release, as the object might already be gone.
                }
                finally
                {
                    obj = null;
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
                    MessageBox.Show(result.ErrorMessage, "调试工具", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (result.OriginalXmlNode == null)
                {
                    MessageBox.Show("未能获取选中内容的XML节点。", "调试工具", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var logger = _serviceContainer.CreateDebugLogger();
                logger.StartSession();
                string savedPath = await logger.LogSelectionXmlAsync(result.OriginalXmlNode);
                string formattedXml = System.Xml.Linq.XDocument.Parse(result.OriginalXmlNode.ToString()).ToString();


                // Show in dialog
                string caption = $"选中内容 XML 结构 - {result.ModeAsString()} (已保存至: {Path.GetFileName(savedPath)})";
                ShowTextInScrollableMessageBox(formattedXml, caption);

                // Show success message
                MessageBox.Show(
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
                MessageBox.Show("调试功能发生错误：\n" + ex.ToString(), "调试工具异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    MessageBox.Show("无法获取当前页面内容。", "调试工具", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var logger = _serviceContainer.CreateDebugLogger();
                logger.StartSession();
                string savedPath = await logger.LogPageXmlAsync(xmlContent);
                string formattedXml = System.Xml.Linq.XDocument.Parse(xmlContent).ToString();

                // Show in dialog
                string caption = $"OneNote XML 结构 (已保存至: {Path.GetFileName(savedPath)})";
                ShowTextInScrollableMessageBox(formattedXml, caption);

                // Show success message
                MessageBox.Show(
                    $"XML已保存至：\n{savedPath}\n\n文件大小: {new FileInfo(savedPath).Length / 1024.0:F2} KB",
                    "调试工具 - 保存成功",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("调试功能发生错误：\n" + ex.ToString(), "调试工具异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        /// <summary>
        /// Helper function: Creates a form with a scrollbar to display a large amount of text.
        /// </summary>
        private void ShowTextInScrollableMessageBox(string text, string caption)
        {
            Form form = new Form
            {
                Text = caption,
                Size = new System.Drawing.Size(600, 400),
                StartPosition = FormStartPosition.CenterParent
            };
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
            form.ShowDialog();
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
