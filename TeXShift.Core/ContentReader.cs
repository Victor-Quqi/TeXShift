using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace TeXShift.Core
{
    /// <summary>
    /// Handles reading and parsing content from a OneNote page.
    /// </summary>
    public class ContentReader : IContentReader
    {
        private readonly OneNote.Application _oneNoteApp;

        public ContentReader(OneNote.Application oneNoteApp)
        {
            _oneNoteApp = oneNoteApp;
        }

        /// <summary>
        /// Asynchronously extracts text content based on the user's current selection or cursor position.
        /// </summary>
        public async Task<ReadResult> ExtractContentAsync()
        {
            // Wrap COM calls in Task.Run to avoid blocking UI thread
            // OneNote COM objects must be accessed on STA thread
            return await Task.Run(() => ExtractContent()).ConfigureAwait(false);
        }

        /// <summary>
        /// Extracts text content based on the user's current selection or cursor position.
        /// (Synchronous version - kept for internal use)
        /// </summary>
        private ReadResult ExtractContent()
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

            return ParseXmlContent(xmlContent, pageId);
        }

        private ReadResult ParseXmlContent(string xmlContent, string pageId)
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
                return HandleCursorMode(deepestSelectedNodes.First(), ns, pageId);
            }
            else
            {
                return HandleSelectionMode(deepestSelectedNodes, ns, pageId);
            }
        }

        private ReadResult HandleCursorMode(XElement cursorNode, XNamespace ns, string pageId)
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
            string objectId = outlineContainer.Attribute("objectID")?.Value;

            return new ReadResult
            {
                IsSuccess = true,
                Mode = DetectionMode.Cursor,
                ExtractedText = extractedText,
                PageId = pageId,
                TargetObjectId = objectId,
                OriginalXmlNode = outlineContainer
            };
        }

        private ReadResult HandleSelectionMode(System.Collections.Generic.List<XElement> selectedNodes, XNamespace ns, string pageId)
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

            // Find the common parent OE or Outline node
            var firstTextNode = selectedNodes.FirstOrDefault(n => n.Name == ns + "T");
            var parentOE = firstTextNode?.Ancestors(ns + "OE").FirstOrDefault();
            var parentOutline = firstTextNode?.Ancestors(ns + "Outline").FirstOrDefault();

            // Use OE for selection mode, fallback to Outline
            var targetNode = parentOE ?? parentOutline;
            string objectId = targetNode?.Attribute("objectID")?.Value;

            return new ReadResult
            {
                IsSuccess = true,
                Mode = DetectionMode.Selection,
                ExtractedText = extractedText,
                PageId = pageId,
                TargetObjectId = objectId,
                OriginalXmlNode = targetNode
            };
        }
    }
}