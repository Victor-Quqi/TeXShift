using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace TeXShift.Core.Logging
{
    /// <summary>
    /// Handles writing debug artifacts to the file system.
    /// </summary>
    public class DebugLogger : IDebugLogger
    {
        private string _sessionTimestamp;
        private readonly string _customOutputPath;
        public string DebugSessionFolder { get; private set; }

        /// <summary>
        /// Creates a new DebugLogger instance.
        /// </summary>
        /// <param name="customOutputPath">Custom output directory path. If null or empty, uses default location.</param>
        public DebugLogger(string customOutputPath = null)
        {
            _customOutputPath = customOutputPath;
        }

        public void StartSession()
        {
            _sessionTimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            DebugSessionFolder = PrepareDebugFolder();
        }

        public async Task LogInputMarkdownAsync(string markdown)
        {
            if (DebugSessionFolder == null) return;
            string inputFile = Path.Combine(DebugSessionFolder, $"01_Input_Markdown_{_sessionTimestamp}.md");
            await Task.Run(() => File.WriteAllText(inputFile, markdown, Encoding.UTF8));
        }

        public async Task LogOriginalXmlAsync(XNode xmlNode)
        {
            if (DebugSessionFolder == null || xmlNode == null) return;
            string originalXmlFile = Path.Combine(DebugSessionFolder, $"02_Original_XML_{_sessionTimestamp}.xml");
            await Task.Run(() => File.WriteAllText(originalXmlFile, xmlNode.ToString(), Encoding.UTF8));
        }

        public async Task LogConvertedXmlAsync(XElement convertedXml)
        {
            if (DebugSessionFolder == null || convertedXml == null) return;
            string convertedXmlFile = Path.Combine(DebugSessionFolder, $"03_Converted_XML_{_sessionTimestamp}.xml");
            await Task.Run(() => File.WriteAllText(convertedXmlFile, convertedXml.ToString(), Encoding.UTF8));
        }

        public async Task LogFinalPageXmlAsync(string finalXml)
        {
            if (DebugSessionFolder == null) return;
            string finalXmlFile = Path.Combine(DebugSessionFolder, $"04_Final_Page_XML_{_sessionTimestamp}.xml");
            await Task.Run(() => File.WriteAllText(finalXmlFile, FormatXml(finalXml), Encoding.UTF8));
        }

        public async Task LogErrorAsync(Exception ex)
        {
            if (DebugSessionFolder == null) return;
            try
            {
                string errorLogFile = Path.Combine(DebugSessionFolder, $"ERROR_{_sessionTimestamp}.txt");
                string errorContent = $"转换失败\n\n时间: {DateTime.Now}\n\n错误消息:\n{ex.Message}\n\n堆栈跟踪:\n{ex.StackTrace}";
                await Task.Run(() => File.WriteAllText(errorLogFile, errorContent, Encoding.UTF8));
            }
            catch (Exception logEx)
            {
                // Prevent logging errors from causing a crash
                System.Diagnostics.Debug.WriteLine($"Warning: Failed to log error: {logEx.Message}");
            }
        }

        public async Task<string> LogSelectionXmlAsync(XNode selectionXml)
        {
            if (DebugSessionFolder == null || selectionXml == null) return null;

            string filename = $"Selection_XML_{_sessionTimestamp}.xml";
            string fullPath = Path.Combine(DebugSessionFolder, filename);
            string formattedXml = FormatXml(selectionXml.ToString());

            await Task.Run(() => File.WriteAllText(fullPath, formattedXml, Encoding.UTF8));
            return fullPath;
        }

        public async Task<string> LogPageXmlAsync(string pageXml)
        {
            if (DebugSessionFolder == null) return null;

            string filename = $"OneNote_XML_{_sessionTimestamp}.xml";
            string fullPath = Path.Combine(DebugSessionFolder, filename);
            string formattedXml = FormatXml(pageXml);

            await Task.Run(() => File.WriteAllText(fullPath, formattedXml, Encoding.UTF8));
            return fullPath;
        }

        private string PrepareDebugFolder()
        {
            string debugFolder;

            // Use custom path if provided and valid
            if (!string.IsNullOrWhiteSpace(_customOutputPath))
            {
                debugFolder = _customOutputPath;
            }
            else
            {
                // Fall back to default: DebugOutput in project root
                string assemblyLocation = Assembly.GetExecutingAssembly().Location;
                string currentDir = Path.GetDirectoryName(assemblyLocation);
                string projectRoot = FindProjectRoot(currentDir) ?? currentDir;
                debugFolder = Path.Combine(projectRoot, "DebugOutput");
            }

            if (!Directory.Exists(debugFolder))
            {
                Directory.CreateDirectory(debugFolder);
            }
            return debugFolder;
        }

        private string FindProjectRoot(string startPath)
        {
            DirectoryInfo dir = new DirectoryInfo(startPath);
            while (dir != null)
            {
                if (dir.GetFiles("*.sln").Length > 0)
                {
                    return dir.FullName;
                }
                dir = dir.Parent;
            }
            return null;
        }

        private string FormatXml(string xml)
        {
            try
            {
                return XDocument.Parse(xml).ToString();
            }
            catch (Exception ex)
            {
                // If parsing fails, return original
                System.Diagnostics.Debug.WriteLine($"Warning: Failed to format XML: {ex.Message}");
                return xml;
            }
        }
    }
}