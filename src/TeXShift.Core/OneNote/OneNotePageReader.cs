using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using TeXShift.Core.Abstractions;
using TeXShift.Core.Logging;
using TeXShift.Core.Localization;
using TeXShift.Core.Utils;
using OneNoteInterop = Microsoft.Office.Interop.OneNote;

namespace TeXShift.Core.OneNote
{
    /// <summary>
    /// Handles reading and parsing content from a OneNote page.
    /// </summary>
    public class OneNotePageReader : IContentReader
    {
        private readonly OneNoteInterop.Application _oneNoteApp;
        private static readonly Regex BrTagWithFollowingNewlineRegex = new Regex(
            "<br\\b[^>]*>(?:\\r\\n|\\r|\\n)?",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public OneNotePageReader(OneNoteInterop.Application oneNoteApp)
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
            OneNoteInterop.Windows windows = null;
            OneNoteInterop.Window window = null;
            try
            {
                string pageId;
                using (PerformanceTraceContext.Measure("Read.OneNote.GetCurrentPageId"))
                {
                    windows = _oneNoteApp.Windows;
                    window = windows.CurrentWindow;
                    pageId = window?.CurrentPageId;
                }

                if (string.IsNullOrEmpty(pageId))
                {
                    return new ReadResult { IsSuccess = false, ErrorMessage = Resources.GetString("Error_NoPageId") };
                }

                string xmlContent;
                using (PerformanceTraceContext.Measure("Read.OneNote.GetPageContent", "piSelection"))
                {
                    // piSelection includes selection markers but excludes binary payloads (e.g., InkDrawing/Data base64).
                    _oneNoteApp.GetPageContent(pageId, out xmlContent, OneNoteInterop.PageInfo.piSelection, OneNoteInterop.XMLSchema.xs2013);
                }

                if (PerformanceTraceContext.GetCurrent() != null)
                {
                    PerformanceTraceContext.AddMetric("Read.PageXmlChars", (xmlContent?.Length ?? 0).ToString());
                    PerformanceTraceContext.AddMetric("Read.PageXml.InkDrawingCount", CountOccurrences(xmlContent, "<one:InkDrawing").ToString());
                    PerformanceTraceContext.AddMetric("Read.PageXml.DataCount", CountOccurrences(xmlContent, "<one:Data").ToString());
                }

                if (string.IsNullOrEmpty(xmlContent))
                {
                    return new ReadResult { IsSuccess = false, ErrorMessage = Resources.GetString("Error_GetPageContentFailed") };
                }

                using (PerformanceTraceContext.Measure("Read.ParseXmlContent"))
                {
                    return ParseXmlContent(xmlContent, pageId);
                }
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
            XDocument doc;
            using (PerformanceTraceContext.Measure("Read.Parse.XDocument.Parse"))
            {
                doc = XDocument.Parse(xmlContent);
            }

            var ns = doc.Root.Name.Namespace;

            bool hasTodoTagDef = false;
            using (PerformanceTraceContext.Measure("Read.Parse.FindTodoTagDef"))
            {
                hasTodoTagDef = doc.Root?.Elements(ns + "TagDef")
                    .Any(e => string.Equals((string)e.Attribute("index"), "0", StringComparison.Ordinal)) == true;
            }

            var quickStyleDefinitions = doc.Root?
                .Elements(ns + "QuickStyleDef")
                .Select(element => new XElement(element))
                .ToList() ?? new List<XElement>();

            List<XElement> deepestSelectedNodes;
            using (PerformanceTraceContext.Measure("Read.Parse.FindSelectedNodes"))
            {
                deepestSelectedNodes = doc.Descendants()
                    .Where(e => e.Attribute("selected") != null && !e.Elements().Any(child => child.Attribute("selected") != null))
                    .ToList();
            }

            PerformanceTraceContext.AddMetric("Read.SelectedNodeCount", deepestSelectedNodes.Count.ToString());

            if (!deepestSelectedNodes.Any())
            {
                return new ReadResult { IsSuccess = false, Mode = DetectionMode.None, ErrorMessage = Resources.GetString("Error_NoSelection") };
            }

            bool isCursorMode = deepestSelectedNodes.Count == 1 &&
                                  deepestSelectedNodes.First().Name == ns + "T" &&
                                  string.IsNullOrEmpty(deepestSelectedNodes.First().Value);

            if (isCursorMode)
            {
                using (PerformanceTraceContext.Measure("Read.HandleCursorMode"))
                {
                    var result = HandleCursorMode(deepestSelectedNodes.First(), ns, pageId);
                    if (result != null)
                    {
                        result.PageHasTodoTagDef = hasTodoTagDef;
                        result.PageQuickStyleDefinitions = quickStyleDefinitions;
                    }
                    return result;
                }
            }
            else
            {
                using (PerformanceTraceContext.Measure("Read.HandleSelectionMode"))
                {
                    var result = HandleSelectionMode(deepestSelectedNodes, ns, pageId);
                    if (result != null)
                    {
                        result.PageHasTodoTagDef = hasTodoTagDef;
                        result.PageQuickStyleDefinitions = quickStyleDefinitions;
                    }
                    return result;
                }
            }
        }

        private ReadResult HandleCursorMode(XElement cursorNode, XNamespace ns, string pageId)
        {
            var outlineContainer = cursorNode.Ancestors(ns + "Outline").FirstOrDefault();
            if (outlineContainer == null)
            {
                return new ReadResult { IsSuccess = false, Mode = DetectionMode.Error, ErrorMessage = Resources.GetString("Error_NoTextBoxContainer") };
            }

            var sb = new StringBuilder();
            var rootChildren = outlineContainer.Element(ns + "OEChildren");
            if (rootChildren != null)
            {
                PerformanceTraceContext.AddMetric("Read.Cursor.RootOeCount", rootChildren.Elements(ns + "OE").Count().ToString());
                var cursorOe = cursorNode.Ancestors(ns + "OE").FirstOrDefault();
                using (PerformanceTraceContext.Measure("Read.Cursor.BuildTextFromOEs"))
                {
                    if (ShouldSkipCursorOe(cursorOe, rootChildren, ns))
                    {
                        foreach (var oe in rootChildren.Elements(ns + "OE"))
                        {
                            if (oe == cursorOe)
                            {
                                continue;
                            }
                            ProcessOE(oe, ns, sb, 0);
                        }
                    }
                    else
                    {
                        ProcessOEChildren(rootChildren, ns, sb, 0);
                    }
                }
            }

            // ProcessOE appends a line terminator after each OE. Remove only the final terminator
            // (artifact of the join), while preserving intentional trailing blank lines.
            string extractedText = TrimSingleTrailingNewline(sb.ToString());
            if (string.IsNullOrWhiteSpace(extractedText))
            {
                // Some OneNote objects (e.g., Table/Image) are not stored as text in <one:T>.
                // In those cases, extractedText can be empty even though there is convertible content.
                // Keep Cursor mode usable for reverse conversion/debug tooling.
                if (!ContainsNonTextConvertibleContent(outlineContainer, ns))
                {
                    return new ReadResult
                    {
                        IsSuccess = false,
                        Mode = DetectionMode.Cursor,
                        ErrorMessage = Resources.GetString("Error_EmptyTextBox")
                    };
                }

                extractedText = string.Empty;
            }

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

        private ReadResult HandleSelectionMode(List<XElement> selectedNodes, XNamespace ns, string pageId)
        {
            List<XElement> parentOEs;
            using (PerformanceTraceContext.Measure("Read.Selection.FindParentOes"))
            {
                parentOEs = FindParentOENodes(selectedNodes, ns);
            }

            PerformanceTraceContext.AddMetric("Read.Selection.ParentOeCount", parentOEs.Count.ToString());
            if (!parentOEs.Any())
            {
                return new ReadResult { IsSuccess = false, Mode = DetectionMode.Selection, ErrorMessage = Resources.GetString("Error_NoValidTextContainer") };
            }

            string extractedText;
            using (PerformanceTraceContext.Measure("Read.Selection.BuildTextFromOEs"))
            {
                extractedText = TrimSingleTrailingNewline(BuildTextFromOENodes(parentOEs, ns));
            }
            if (string.IsNullOrWhiteSpace(extractedText))
            {
                // Selection may target a non-text OE (e.g., selecting an embedded Image or a Table).
                // Treat it as valid content for reverse conversion so we can emit placeholders/Markdown.
                if (!ContainsNonTextConvertibleContent(parentOEs, ns))
                {
                    return new ReadResult { IsSuccess = false, Mode = DetectionMode.Selection, ErrorMessage = Resources.GetString("Error_NoValidTextContent") };
                }

                extractedText = string.Empty;
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
        private List<XElement> FindParentOENodes(List<XElement> selectedNodes, XNamespace ns)
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
        private string BuildTextFromOENodes(List<XElement> parentOEs, XNamespace ns)
        {
            // Reconstruct hierarchical relationship: top-level OEs are those whose parent isn't also selected
            var sb = new StringBuilder();
            var topLevelOEs = parentOEs.Where(oe => oe.Parent?.Parent != null && !parentOEs.Contains(oe.Parent.Parent)).ToList();

            foreach (var oe in topLevelOEs)
            {
                ProcessOE(oe, ns, sb, 0);
            }

            return sb.ToString();
        }

        private static string TrimSingleTrailingNewline(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            if (text.EndsWith("\r\n", StringComparison.Ordinal))
            {
                return text.Substring(0, text.Length - 2);
            }

            char last = text[text.Length - 1];
            if (last == '\n' || last == '\r')
            {
                return text.Substring(0, text.Length - 1);
            }

            return text;
        }

        private static int CountOccurrences(string haystack, string needle)
        {
            if (string.IsNullOrEmpty(haystack) || string.IsNullOrEmpty(needle))
            {
                return 0;
            }

            int count = 0;
            int index = 0;
            while (true)
            {
                index = haystack.IndexOf(needle, index, StringComparison.OrdinalIgnoreCase);
                if (index < 0)
                {
                    return count;
                }

                count++;
                index += needle.Length;
            }
        }

        /// <summary>
        /// Collects object IDs from OE nodes and adds them to the result.
        /// </summary>
        private void CollectObjectIds(List<XElement> parentOEs, ReadResult result)
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

            // XElement.Value automatically decodes XML entities from CDATA sections:
            //   &gt; → >  (for Markdown syntax recognition)
            //   &amp;lt; → &lt; (preserves user's HTML entities)
            // Additional HtmlDecode would cause double-decoding and content loss.
            var oeText = string.Concat(oe.Elements(ns + "T").Select(t => t.Value));
            // OneNote may wrap link-like text with a simple <a href="...">...</a> inside <one:T>.
            // Strip that wrapper (keep displayed text) so the extracted Markdown stays pure.
            oeText = OneNoteAutoLinkNormalizer.Normalize(oeText);
            oeText = NormalizeBrToNewlines(oeText);

            // OneNote stores HTML entities inside the <one:T> CDATA. Decode once so user-authored Markdown
            // syntax (e.g., '>' for quotes, '"' for strings) is restored.
            // User-typed entities are typically stored as "&amp;..." and remain entities after a single decode.
            oeText = WebUtility.HtmlDecode(oeText ?? string.Empty);
            sb.AppendLine(oeText);

            var nestedChildren = oe.Element(ns + "OEChildren");
            if (nestedChildren != null)
            {
                ProcessOEChildren(nestedChildren, ns, sb, indentLevel + 1);
            }
        }

        private bool ShouldSkipCursorOe(XElement cursorOe, XElement rootChildren, XNamespace ns)
        {
            if (cursorOe == null || rootChildren == null)
            {
                return false;
            }

            if (cursorOe.Parent != rootChildren)
            {
                return false;
            }

            var oeList = rootChildren.Elements(ns + "OE").ToList();
            if (oeList.Count <= 1)
            {
                return false;
            }

            if (oeList.First() != cursorOe)
            {
                return false;
            }

            var nested = cursorOe.Element(ns + "OEChildren");
            if (nested != null && nested.Elements(ns + "OE").Any())
            {
                return false;
            }

            foreach (var child in cursorOe.Elements())
            {
                var localName = child.Name.LocalName;
                if (!string.Equals(localName, "T", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(localName, "Meta", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            var oeText = string.Concat(cursorOe.Elements(ns + "T").Select(t => t.Value));
            return string.IsNullOrWhiteSpace(oeText);
        }

        private static string NormalizeBrToNewlines(string htmlOrText)
        {
            if (string.IsNullOrEmpty(htmlOrText) || htmlOrText.IndexOf("<br", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return htmlOrText;
            }

            // OneNote may serialize explicit line breaks as <br /> tags inside <one:T> HTML.
            // Normalize them to plain newlines so Markdown parsing stays correct.
            // OneNote's CDATA often contains a literal newline after the <br /> (XML pretty printing), which would
            // otherwise become a doubled blank line after normalization.
            return BrTagWithFollowingNewlineRegex.Replace(htmlOrText, "\n");
        }

        private static bool ContainsNonTextConvertibleContent(XElement element, XNamespace ns)
        {
            if (element == null)
            {
                return false;
            }

            // Keep this conservative: only allow elements we explicitly support in reverse conversion.
            return element.Descendants(ns + "Table").Any() || element.Descendants(ns + "Image").Any();
        }

        private static bool ContainsNonTextConvertibleContent(IEnumerable<XElement> elements, XNamespace ns)
        {
            if (elements == null)
            {
                return false;
            }

            foreach (var e in elements)
            {
                if (ContainsNonTextConvertibleContent(e, ns))
                {
                    return true;
                }
            }

            return false;
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
                    RuntimeLog.Write($"Failed to release OneNote COM object: {ex}");
                }
            }
        }
    }
}
