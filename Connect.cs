using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
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
 
         /// <summary>
         /// Called when the add-in is connected to OneNote.
         /// </summary>
         public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
         {
             _oneNoteApp = (OneNote.Application)Application;
         }
 
         /// <summary>
         /// Called when the add-in is disconnected from OneNote.
         /// </summary>
         public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
         {
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

        public void OnConvertButtonClick(IRibbonControl control)
        {
            try
            {
                var reader = new ContentReader(_oneNoteApp);
                var result = reader.ExtractContent();

                if (result.IsSuccess)
                {
                    ShowTextInScrollableMessageBox(result.ExtractedText, $"检测到: {result.ModeAsString()}");
                }
                else
                {
                    MessageBox.Show(result.ErrorMessage, "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("插件发生未知错误：\n" + ex.ToString(), "插件异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Debug button: Shows and saves the raw OneNote XML structure for current selection.
        /// </summary>
        public void OnDebugXmlButtonClick(IRibbonControl control)
        {
            try
            {
                string pageId = _oneNoteApp.Windows.CurrentWindow?.CurrentPageId;
                if (string.IsNullOrEmpty(pageId))
                {
                    MessageBox.Show("无法获取当前页面ID。", "调试工具", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Get full page XML
                string xmlContent;
                _oneNoteApp.GetPageContent(pageId, out xmlContent, OneNote.PageInfo.piAll);

                if (string.IsNullOrEmpty(xmlContent))
                {
                    MessageBox.Show("获取页面XML失败。", "调试工具", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Format XML for better readability
                string formattedXml = FormatXml(xmlContent);

                // Save to file
                string savedPath = SaveDebugXml(formattedXml);

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
        /// Saves debug XML to DebugOutput folder with timestamp.
        /// </summary>
        private string SaveDebugXml(string xml)
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
