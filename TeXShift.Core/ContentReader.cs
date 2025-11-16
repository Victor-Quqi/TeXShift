using System.Linq;
using System.Text;
using System.Xml.Linq;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace TeXShift.Core
{
    /// <summary>
    /// Handles reading and parsing content from a OneNote page.
    /// </summary>
    public class ContentReader
    {
        private readonly OneNote.Application _oneNoteApp;

        public ContentReader(OneNote.Application oneNoteApp)
        {
            _oneNoteApp = oneNoteApp;
        }

        /// <summary>
        /// Extracts text content based on the user's current selection or cursor position.
        /// </summary>
        public ReadResult ExtractContent()
        {
            string pageId = _oneNoteApp.Windows.CurrentWindow?.CurrentPageId;
            if (string.IsNullOrEmpty(pageId))
            {
                return new ReadResult { IsSuccess = false, ErrorMessage = "无法获取当前页面的ID。" };
            }

            string xmlContent;
            _oneNoteApp.GetPageContent(pageId, out xmlContent, OneNote.PageInfo.piAll);

            if (string.IsNullOrEmpty(xmlContent))
            {
                return new ReadResult { IsSuccess = false, ErrorMessage = "获取页面内容失败。" };
            }

            return ParseXmlContent(xmlContent);
        }

        private ReadResult ParseXmlContent(string xmlContent)
        {
            var doc = XDocument.Parse(xmlContent);
            var ns = doc.Root.Name.Namespace;

            var deepestSelectedNodes = doc.Descendants()
                .Where(e => e.Attribute("selected") != null && !e.Elements().Any(child => child.Attribute("selected") != null))
                .ToList();

            if (!deepestSelectedNodes.Any())
            {
                return new ReadResult { IsSuccess = false, Mode = DetectionMode.None, ErrorMessage = "未检测到选中的内容。\n\n请先用鼠标选中一些文字，或将光标点入一个文本框中。" };
            }

            bool isCursorMode = deepestSelectedNodes.Count == 1 &&
                                deepestSelectedNodes.First().Name == ns + "T" &&
                                string.IsNullOrEmpty(deepestSelectedNodes.First().Value);

            if (isCursorMode)
            {
                return HandleCursorMode(deepestSelectedNodes.First(), ns);
            }
            else
            {
                return HandleSelectionMode(deepestSelectedNodes, ns);
            }
        }

        private ReadResult HandleCursorMode(XElement cursorNode, XNamespace ns)
        {
            var outlineContainer = cursorNode.Ancestors(ns + "Outline").FirstOrDefault();
            if (outlineContainer == null)
            {
                return new ReadResult { IsSuccess = false, Mode = DetectionMode.Error, ErrorMessage = "错误：未能找到光标所在的文本框容器。" };
            }

            var sb = new StringBuilder();
            foreach (var oeNode in outlineContainer.Descendants(ns + "OE"))
            {
                foreach (var textNode in oeNode.Elements(ns + "T"))
                {
                    sb.Append(textNode.Value);
                }
                sb.AppendLine();
            }

            string extractedText = sb.ToString().TrimEnd('\r', '\n');
            return new ReadResult { IsSuccess = true, Mode = DetectionMode.Cursor, ExtractedText = extractedText };
        }

        private ReadResult HandleSelectionMode(System.Collections.Generic.List<XElement> selectedNodes, XNamespace ns)
        {
            var sb = new StringBuilder();
            foreach (var node in selectedNodes.Where(n => n.Name == ns + "T"))
            {
                sb.Append(node.Value);
            }

            string extractedText = sb.ToString();
            if (string.IsNullOrEmpty(extractedText))
            {
                 return new ReadResult { IsSuccess = false, Mode = DetectionMode.Selection, ErrorMessage = "成功定位到选区，但未能提取出有效文本内容。" };
            }

            return new ReadResult { IsSuccess = true, Mode = DetectionMode.Selection, ExtractedText = extractedText };
        }
    }
}