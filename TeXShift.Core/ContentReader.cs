using System;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
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
            OneNote.Windows windows = null;
            OneNote.Window window = null;
            try
            {
                windows = _oneNoteApp.Windows;
                window = windows.CurrentWindow;
                string pageId = window?.CurrentPageId;

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
            finally
            {
                // Release COM objects in the reverse order of creation.
                SafeReleaseComObject(window);
                SafeReleaseComObject(windows);
            }
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
            var rootChildren = outlineContainer.Element(ns + "OEChildren");
            if (rootChildren != null)
            {
                ProcessOEChildren(rootChildren, ns, sb, 0);
            }

            string extractedText = sb.ToString().TrimEnd('\r', '\n');
            string objectId = outlineContainer.Attribute("objectID")?.Value;

            var result = new ReadResult
            {
                IsSuccess = true,
                Mode = DetectionMode.Cursor,
                ExtractedText = extractedText,
                PageId = pageId,
                OriginalXmlNode = outlineContainer,
                SourceOutlineWidth = ExtractOutlineWidth(outlineContainer, ns)
            };
            if (objectId != null)
            {
                result.TargetObjectIds.Add(objectId);
            }
            return result;
        }

        private ReadResult HandleSelectionMode(System.Collections.Generic.List<XElement> selectedNodes, XNamespace ns, string pageId)
        {
            var parentOEs = FindParentOENodes(selectedNodes, ns);
            if (!parentOEs.Any())
            {
                return new ReadResult { IsSuccess = false, Mode = DetectionMode.Selection, ErrorMessage = "成功定位到选区，但未能找到有效的文本容器。" };
            }

            string extractedText = BuildTextFromOENodes(parentOEs, ns);
            if (string.IsNullOrEmpty(extractedText))
            {
                return new ReadResult { IsSuccess = false, Mode = DetectionMode.Selection, ErrorMessage = "成功定位到选区，但未能提取出有效文本内容。" };
            }

            var outlineContainer = parentOEs.FirstOrDefault()?.Ancestors(ns + "Outline").FirstOrDefault();

            var result = new ReadResult
            {
                IsSuccess = true,
                Mode = DetectionMode.Selection,
                ExtractedText = extractedText,
                PageId = pageId,
                OriginalXmlNode = parentOEs.FirstOrDefault(), // For attribute preservation
                OriginalXmlNodes = parentOEs.Cast<XElement>().ToList(),
                SourceOutlineWidth = outlineContainer != null ? ExtractOutlineWidth(outlineContainer, ns) : null
            };

            CollectObjectIds(parentOEs, result);
            return result;
        }

        /// <summary>
        /// Finds all unique parent OE nodes from the selected nodes.
        /// </summary>
        private System.Collections.Generic.List<XElement> FindParentOENodes(System.Collections.Generic.List<XElement> selectedNodes, XNamespace ns)
        {
            return selectedNodes
                .Select(n => n.Name == ns + "OE" ? n : n.Ancestors(ns + "OE").FirstOrDefault())
                .Where(oe => oe != null)
                .Distinct()
                .ToList();
        }

        /// <summary>
        /// Builds text from a hierarchical structure of OE nodes.
        /// </summary>
        private string BuildTextFromOENodes(System.Collections.Generic.List<XElement> parentOEs, XNamespace ns)
        {
            // Reconstruct hierarchical relationship: top-level OEs are those whose parent isn't also selected
            var sb = new StringBuilder();
            var topLevelOEs = parentOEs.Where(oe => oe.Parent?.Parent != null && !parentOEs.Contains(oe.Parent.Parent)).ToList();

            for (int i = 0; i < topLevelOEs.Count; i++)
            {
                ProcessOE(topLevelOEs[i], ns, sb, 0);
                if (i < topLevelOEs.Count - 1)
                {
                    sb.AppendLine();
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// Collects object IDs from OE nodes and adds them to the result.
        /// </summary>
        private void CollectObjectIds(System.Collections.Generic.List<XElement> parentOEs, ReadResult result)
        {
            foreach (var oe in parentOEs)
            {
                string objectId = oe.Attribute("objectID")?.Value;
                if (objectId != null)
                {
                    result.TargetObjectIds.Add(objectId);
                }
            }
        }

        /// <summary>
        /// Recursively processes an OEChildren element, building a string with Markdown-compliant indentation.
        /// </summary>
        private void ProcessOEChildren(XElement oeChildren, XNamespace ns, StringBuilder sb, int indentLevel)
        {
            foreach (var oe in oeChildren.Elements(ns + "OE"))
            {
                ProcessOE(oe, ns, sb, indentLevel);
            }
        }

        /// <summary>
        /// Processes a single OE element, appends its text, and recursively handles its children.
        /// </summary>
        private void ProcessOE(XElement oe, XNamespace ns, StringBuilder sb, int indentLevel)
        {
            sb.Append(new string(' ', indentLevel * 4));

            var oeText = string.Concat(oe.Elements(ns + "T").Select(t => t.Value));
            oeText = WebUtility.HtmlDecode(oeText);
            sb.AppendLine(oeText);

            var nestedChildren = oe.Element(ns + "OEChildren");
            if (nestedChildren != null)
            {
                ProcessOEChildren(nestedChildren, ns, sb, indentLevel + 1);
            }
        }


        /// <summary>
        /// Extracts the width from an Outline element's Size child element.
        /// </summary>
        private double? ExtractOutlineWidth(XElement outlineElement, XNamespace ns)
        {
            var sizeElement = outlineElement.Element(ns + "Size");
            if (sizeElement != null)
            {
                var widthAttr = sizeElement.Attribute("width");
                if (widthAttr != null && double.TryParse(widthAttr.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double width))
                {
                    return width;
                }
            }
            return null;
        }

        /// <summary>
        /// Safely releases a COM object.
        /// </summary>
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
    }
}