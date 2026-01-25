using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using TeXShift.Core.Localization;

namespace TeXShift.Core.Logging
{
    /// <summary>
    /// Handles writing debug artifacts to the file system.
    /// </summary>
    public class DebugLogger : IDebugLogger
    {
        private string _sessionTimestamp;
        private DebugSessionKind _sessionKind = DebugSessionKind.ForwardConversion;
        private readonly string _customOutputPath;
        public string DebugSessionFolder { get; private set; }

        private static readonly Encoding Utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

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
            StartSession(DebugSessionKind.ForwardConversion);
        }

        public void StartSession(DebugSessionKind kind)
        {
            _sessionKind = kind;
            _sessionTimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            DebugSessionFolder = PrepareDebugFolder();
        }

        public async Task LogInputMarkdownAsync(string markdown)
        {
            if (DebugSessionFolder == null) return;
            string inputFile = Path.Combine(DebugSessionFolder, GetArtifactFileName(DebugArtifact.InputMarkdown));
            await Task.Run(() => File.WriteAllText(inputFile, markdown, Utf8NoBom));
        }

        public async Task LogReversedMarkdownAsync(string markdown)
        {
            if (DebugSessionFolder == null) return;
            string outputFile = Path.Combine(DebugSessionFolder, GetArtifactFileName(DebugArtifact.ReversedMarkdown));
            await Task.Run(() => File.WriteAllText(outputFile, markdown ?? string.Empty, Utf8NoBom));
        }

        public async Task LogOriginalXmlAsync(XNode xmlNode)
        {
            if (DebugSessionFolder == null || xmlNode == null) return;
            string originalXmlFile = Path.Combine(DebugSessionFolder, GetArtifactFileName(DebugArtifact.OriginalXml));
            await Task.Run(() => File.WriteAllText(originalXmlFile, xmlNode.ToString(), Utf8NoBom));
        }

        public async Task LogConvertedXmlAsync(XElement convertedXml)
        {
            if (DebugSessionFolder == null || convertedXml == null) return;
            string convertedXmlFile = Path.Combine(DebugSessionFolder, GetArtifactFileName(DebugArtifact.ConvertedXml));
            await Task.Run(() => File.WriteAllText(convertedXmlFile, convertedXml.ToString(), Utf8NoBom));
        }

        public async Task LogFinalPageXmlAsync(string finalXml)
        {
            if (DebugSessionFolder == null) return;
            string finalXmlFile = Path.Combine(DebugSessionFolder, GetArtifactFileName(DebugArtifact.FinalPageXml));
            await Task.Run(() => File.WriteAllText(finalXmlFile, FormatXml(finalXml), Utf8NoBom));
        }

        public async Task LogFinalPageXmlFullAsync(string finalXml)
        {
            if (DebugSessionFolder == null) return;
            string finalXmlFile = Path.Combine(DebugSessionFolder, GetArtifactFileName(DebugArtifact.FinalPageXmlFull));
            await Task.Run(() => File.WriteAllText(finalXmlFile, FormatXml(finalXml), Utf8NoBom));
        }

        public async Task LogPerformanceAsync(string report)
        {
            if (DebugSessionFolder == null) return;
            string perfFile = Path.Combine(DebugSessionFolder, GetArtifactFileName(DebugArtifact.Performance));
            await Task.Run(() => File.WriteAllText(perfFile, report ?? string.Empty, Utf8NoBom));
        }

        public async Task LogErrorAsync(Exception ex)
        {
            if (DebugSessionFolder == null) return;
            try
            {
                string errorLogFile = Path.Combine(DebugSessionFolder, GetArtifactFileName(DebugArtifact.Error));
                string errorContent = string.Format(Resources.GetString("Debug_ConversionFailed"), DateTime.Now, ex.Message, ex.StackTrace);
                await Task.Run(() => File.WriteAllText(errorLogFile, errorContent, Utf8NoBom));
            }
            catch (Exception logEx)
            {
                // Prevent logging errors from causing a crash
                System.Diagnostics.Debug.WriteLine($"Warning: Failed to log error. {logEx.GetType().Name}: {logEx.Message}");
                System.Diagnostics.Debug.WriteLine(logEx);
                System.Diagnostics.Debug.WriteLine(ex);
            }
        }

        public async Task<(string Path, string FormattedXml)> LogSelectionXmlAsync(XNode selectionXml)
        {
            if (DebugSessionFolder == null || selectionXml == null) return (null, null);

            string filename = GetArtifactFileName(DebugArtifact.SelectionXml);
            string fullPath = Path.Combine(DebugSessionFolder, filename);
            string formattedXml = FormatXml(selectionXml.ToString());

            await Task.Run(() => File.WriteAllText(fullPath, formattedXml, Utf8NoBom));
            return (fullPath, formattedXml);
        }

        public async Task<(string Path, string FormattedXml)> LogSelectionXmlAsync(System.Collections.Generic.IEnumerable<XElement> selectionXmlNodes)
        {
            if (DebugSessionFolder == null || selectionXmlNodes == null) return (null, null);

            var nodesList = selectionXmlNodes as System.Collections.Generic.IList<XElement>
                ?? new System.Collections.Generic.List<XElement>(selectionXmlNodes);

            if (nodesList.Count == 0) return (null, null);

            string filename = GetArtifactFileName(DebugArtifact.SelectionXml);
            string fullPath = Path.Combine(DebugSessionFolder, filename);

            // Wrap multiple nodes in a container element for valid XML
            var container = new XElement("SelectedNodes",
                new XAttribute("count", nodesList.Count),
                nodesList);
            string formattedXml = container.ToString();

            await Task.Run(() => File.WriteAllText(fullPath, formattedXml, Utf8NoBom));
            return (fullPath, formattedXml);
        }

        public async Task<string> LogPageXmlAsync(string pageXml)
        {
            if (DebugSessionFolder == null) return null;

            string filename = GetArtifactFileName(DebugArtifact.PageXml);
            string fullPath = Path.Combine(DebugSessionFolder, filename);
            string formattedXml = FormatXml(pageXml);

            await Task.Run(() => File.WriteAllText(fullPath, formattedXml, Utf8NoBom));
            return fullPath;
        }

        private enum DebugArtifact
        {
            InputMarkdown,
            SelectionXml,
            OriginalXml,
            ReversedMarkdown,
            ConvertedXml,
            FinalPageXml,
            FinalPageXmlFull,
            PageXml,
            Performance,
            Error
        }

        private string GetArtifactFileName(DebugArtifact artifact)
        {
            string prefix = GetSessionPrefix(_sessionKind);

            switch (_sessionKind)
            {
                case DebugSessionKind.ReverseConversion:
                    return GetReverseFileName(prefix, artifact);
                case DebugSessionKind.SelectionXml:
                    return GetSelectionXmlFileName(prefix, artifact);
                case DebugSessionKind.PageXml:
                    return GetPageXmlFileName(prefix, artifact);
                default:
                    return GetForwardFileName(prefix, artifact);
            }
        }

        private string GetForwardFileName(string prefix, DebugArtifact artifact)
        {
            switch (artifact)
            {
                case DebugArtifact.InputMarkdown:
                    return $"{prefix}01_Input_Markdown_{_sessionTimestamp}.md";
                case DebugArtifact.SelectionXml:
                    return $"{prefix}00_Selection_XML_{_sessionTimestamp}.xml";
                case DebugArtifact.OriginalXml:
                    return $"{prefix}02_Original_XML_{_sessionTimestamp}.xml";
                case DebugArtifact.ConvertedXml:
                    return $"{prefix}03_Converted_XML_{_sessionTimestamp}.xml";
                case DebugArtifact.FinalPageXml:
                    return $"{prefix}04_Final_Page_XML_Basic_{_sessionTimestamp}.xml";
                case DebugArtifact.FinalPageXmlFull:
                    return $"{prefix}04_Final_Page_XML_Full_{_sessionTimestamp}.xml";
                case DebugArtifact.ReversedMarkdown:
                    return $"{prefix}05_Reversed_Markdown_{_sessionTimestamp}.md";
                case DebugArtifact.Performance:
                    return $"{prefix}06_Perf_{_sessionTimestamp}.txt";
                case DebugArtifact.PageXml:
                    return $"{prefix}00_Page_XML_{_sessionTimestamp}.xml";
                default:
                    return $"{prefix}99_ERROR_{_sessionTimestamp}.txt";
            }
        }

        private string GetReverseFileName(string prefix, DebugArtifact artifact)
        {
            switch (artifact)
            {
                case DebugArtifact.SelectionXml:
                    return $"{prefix}00_Selection_XML_{_sessionTimestamp}.xml";
                case DebugArtifact.OriginalXml:
                    return $"{prefix}01_Original_XML_{_sessionTimestamp}.xml";
                case DebugArtifact.ReversedMarkdown:
                    return $"{prefix}02_Reversed_Markdown_{_sessionTimestamp}.md";
                case DebugArtifact.ConvertedXml:
                    // Reverse conversion writes back a plain text outline; name it explicitly to avoid confusion with forward conversion output.
                    return $"{prefix}03_WriteBack_XML_{_sessionTimestamp}.xml";
                case DebugArtifact.FinalPageXml:
                    return $"{prefix}04_Final_Page_XML_Basic_{_sessionTimestamp}.xml";
                case DebugArtifact.FinalPageXmlFull:
                    return $"{prefix}04_Final_Page_XML_Full_{_sessionTimestamp}.xml";
                case DebugArtifact.Performance:
                    return $"{prefix}05_Perf_{_sessionTimestamp}.txt";
                case DebugArtifact.InputMarkdown:
                    return $"{prefix}00_Input_Markdown_{_sessionTimestamp}.md";
                case DebugArtifact.PageXml:
                    return $"{prefix}00_Page_XML_{_sessionTimestamp}.xml";
                default:
                    return $"{prefix}99_ERROR_{_sessionTimestamp}.txt";
            }
        }

        private string GetSelectionXmlFileName(string prefix, DebugArtifact artifact)
        {
            switch (artifact)
            {
                case DebugArtifact.SelectionXml:
                    return $"{prefix}01_Selection_XML_{_sessionTimestamp}.xml";
                default:
                    return $"{prefix}99_ERROR_{_sessionTimestamp}.txt";
            }
        }

        private string GetPageXmlFileName(string prefix, DebugArtifact artifact)
        {
            switch (artifact)
            {
                case DebugArtifact.PageXml:
                    return $"{prefix}01_OneNote_XML_{_sessionTimestamp}.xml";
                default:
                    return $"{prefix}99_ERROR_{_sessionTimestamp}.txt";
            }
        }

        private static string GetSessionPrefix(DebugSessionKind kind)
        {
            switch (kind)
            {
                case DebugSessionKind.ReverseConversion:
                    return "R";
                case DebugSessionKind.SelectionXml:
                    return "S";
                case DebugSessionKind.PageXml:
                    return "P";
                default:
                    return "F";
            }
        }

        private string PrepareDebugFolder()
        {
            return ResolveDebugOutputFolder(_customOutputPath);
        }

        public static string ResolveDebugOutputFolder(string customOutputPath)
        {
            string debugFolder;

            // Use custom path if provided and valid
            if (!string.IsNullOrWhiteSpace(customOutputPath))
            {
                debugFolder = customOutputPath;
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

        private static string FindProjectRoot(string startPath)
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
