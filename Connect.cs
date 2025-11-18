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
             _serviceContainer = null;
             _oneNoteApp = null;
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
        public async void OnConvertButtonClick(IRibbonControl control)
        {
            string debugFolder = null;
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            try
            {
                // Prepare debug output folder
                debugFolder = PrepareDebugFolder();

                // Step 1: Read content from OneNote (async - non-blocking)
                var reader = _serviceContainer.CreateContentReader(_oneNoteApp);
                var readResult = await reader.ExtractContentAsync();

                if (!readResult.IsSuccess)
                {
                    MessageBox.Show(readResult.ErrorMessage, "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Save input Markdown (async file I/O)
                string inputFile = Path.Combine(debugFolder, $"01_Input_Markdown_{timestamp}.md");
                await Task.Run(() => File.WriteAllText(inputFile, readResult.ExtractedText, Encoding.UTF8));

                // Save original XML node (async file I/O)
                if (readResult.OriginalXmlNode != null)
                {
                    string originalXmlFile = Path.Combine(debugFolder, $"02_Original_XML_{timestamp}.xml");
                    await Task.Run(() => File.WriteAllText(originalXmlFile, readResult.OriginalXmlNode.ToString(), Encoding.UTF8));
                }

                // Step 2: Convert Markdown to OneNote XML (async - non-blocking)
                var converter = _serviceContainer.CreateMarkdownConverter();
                var oneNoteXml = await converter.ConvertToOneNoteXmlAsync(readResult.ExtractedText);

                // Save converted XML (async file I/O)
                string convertedXmlFile = Path.Combine(debugFolder, $"03_Converted_XML_{timestamp}.xml");
                await Task.Run(() => File.WriteAllText(convertedXmlFile, oneNoteXml.ToString(), Encoding.UTF8));

                // Step 3: Write back to OneNote (async - non-blocking)
                var writer = _serviceContainer.CreateContentWriter(_oneNoteApp);
                await writer.ReplaceContentAsync(readResult, oneNoteXml);

                // Save final page XML (async)
                string finalPageXml = await Task.Run(() =>
                {
                    string xml;
                    _oneNoteApp.GetPageContent(readResult.PageId, out xml, OneNote.PageInfo.piAll);
                    return xml;
                });
                string finalXmlFile = Path.Combine(debugFolder, $"04_Final_Page_XML_{timestamp}.xml");
                await Task.Run(() => File.WriteAllText(finalXmlFile, FormatXml(finalPageXml), Encoding.UTF8));

                // Success message with debug info
                MessageBox.Show(
                    $"转换成功!\n\n" +
                    $"模式: {readResult.ModeAsString()}\n" +
                    $"处理了 {readResult.ExtractedText.Length} 个字符\n\n" +
                    $"调试文件已保存至:\n{debugFolder}",
                    "TeXShift - 转换完成",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                // Save error log (fire-and-forget async)
                if (debugFolder != null)
                {
                    try
                    {
                        string errorLogFile = Path.Combine(debugFolder, $"ERROR_{timestamp}.txt");
                        await Task.Run(() => File.WriteAllText(errorLogFile,
                            $"转换失败\n\n时间: {DateTime.Now}\n\n错误消息:\n{ex.Message}\n\n堆栈跟踪:\n{ex.StackTrace}",
                            Encoding.UTF8));
                    }
                    catch { }
                }

                MessageBox.Show(
                    $"转换失败:\n\n{ex.Message}\n\n详细信息:\n{ex.StackTrace}",
                    "TeXShift - 错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        /// <summary>
        /// Debug button: Shows and saves the selected content's XML structure only.
        /// </summary>
        public async void OnDebugSelectionXmlButtonClick(IRibbonControl control)
        {
            try
            {
                // Use ContentReader to get selected content (async)
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

                // Format XML for better readability
                string formattedXml = FormatXml(result.OriginalXmlNode.ToString());

                // Save to file (async)
                string debugFolder = PrepareDebugFolder();
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string filename = $"Selection_XML_{timestamp}.xml";
                string fullPath = Path.Combine(debugFolder, filename);
                await Task.Run(() => File.WriteAllText(fullPath, formattedXml, Encoding.UTF8));

                // Show in dialog
                string caption = $"选中内容 XML 结构 - {result.ModeAsString()} (已保存至: {filename})";
                ShowTextInScrollableMessageBox(formattedXml, caption);

                // Show success message
                MessageBox.Show(
                    $"选中内容的XML已保存至：\n{fullPath}\n\n" +
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
                // Get page ID and XML content (async)
                var (pageId, xmlContent) = await Task.Run(() =>
                {
                    string id = _oneNoteApp.Windows.CurrentWindow?.CurrentPageId;
                    if (string.IsNullOrEmpty(id))
                        return (null, null);

                    string xml;
                    _oneNoteApp.GetPageContent(id, out xml, OneNote.PageInfo.piAll);
                    return (id, xml);
                });

                if (string.IsNullOrEmpty(pageId))
                {
                    MessageBox.Show("无法获取当前页面ID。", "调试工具", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (string.IsNullOrEmpty(xmlContent))
                {
                    MessageBox.Show("获取页面XML失败。", "调试工具", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Format XML for better readability
                string formattedXml = FormatXml(xmlContent);

                // Save to file (async)
                string savedPath = await SaveDebugXmlAsync(formattedXml);

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
        /// Formats XML string with proper indentation for readability.
        /// </summary>
        private string FormatXml(string xml)
        {
            try
            {
                var doc = System.Xml.Linq.XDocument.Parse(xml);
                return doc.ToString();
            }
            catch
            {
                // If parsing fails, return original
                return xml;
            }
        }

        /// <summary>
        /// Asynchronously saves debug XML to DebugOutput folder with timestamp.
        /// </summary>
        private async Task<string> SaveDebugXmlAsync(string xml)
        {
            return await Task.Run(() =>
            {
                // Find project root directory (where .sln file is located)
                string assemblyLocation = Assembly.GetExecutingAssembly().Location;
                string currentDir = Path.GetDirectoryName(assemblyLocation);
                string projectRoot = FindProjectRoot(currentDir);

                if (projectRoot == null)
                {
                    // Fallback to assembly location if project root not found
                    projectRoot = currentDir;
                }

                string debugFolder = Path.Combine(projectRoot, "DebugOutput");

                if (!Directory.Exists(debugFolder))
                {
                    Directory.CreateDirectory(debugFolder);
                }

                // Generate filename with timestamp
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string filename = $"OneNote_XML_{timestamp}.xml";
                string fullPath = Path.Combine(debugFolder, filename);

                // Save file
                File.WriteAllText(fullPath, xml, Encoding.UTF8);

                return fullPath;
            });
        }

        /// <summary>
        /// Finds the project root directory by looking for .sln file.
        /// </summary>
        private string FindProjectRoot(string startPath)
        {
            DirectoryInfo dir = new DirectoryInfo(startPath);

            while (dir != null)
            {
                // Check if this directory contains a .sln file
                if (dir.GetFiles("*.sln").Length > 0)
                {
                    return dir.FullName;
                }

                // Move up to parent directory
                dir = dir.Parent;
            }

            return null;
        }

        /// <summary>
        /// Prepares the debug output folder for saving conversion artifacts.
        /// </summary>
        private string PrepareDebugFolder()
        {
            string assemblyLocation = Assembly.GetExecutingAssembly().Location;
            string currentDir = Path.GetDirectoryName(assemblyLocation);
            string projectRoot = FindProjectRoot(currentDir);

            if (projectRoot == null)
            {
                projectRoot = currentDir;
            }

            string debugFolder = Path.Combine(projectRoot, "DebugOutput");

            if (!Directory.Exists(debugFolder))
            {
                Directory.CreateDirectory(debugFolder);
            }

            return debugFolder;
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
