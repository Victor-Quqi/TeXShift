using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using TeXShift.Core.Abstractions;
using TeXShift.Core.Errors;
using TeXShift.Core.Logging;
using TeXShift.Core.Localization;
using TeXShift.Core.Markdown.Processing;
using TeXShift.Core.Utils;
using OneNoteInterop = Microsoft.Office.Interop.OneNote;

namespace TeXShift.Core.OneNote
{
    /// <summary>
    /// Handles writing converted content back to OneNote pages.
    /// </summary>
    public class OneNotePageWriter : IContentWriter
    {
        private readonly OneNoteComDispatcher _comDispatcher;
        private readonly IDebugLogger _logger;
        private readonly XNamespace _ns = OneNoteXml.Namespace;

        internal OneNotePageWriter(OneNoteComDispatcher comDispatcher, IDebugLogger logger = null)
        {
            _comDispatcher = comDispatcher ?? throw new ArgumentNullException(nameof(comDispatcher));
            _logger = logger;
        }

        /// <summary>
        /// Asynchronously replaces content in OneNote based on the read result and converted XML.
        /// </summary>
        /// <param name="readResult">The original read result containing metadata</param>
        /// <param name="newOutlineXml">The new Outline XML element to insert</param>
        public Task ReplaceContentAsync(ReadResult readResult, XElement newOutlineXml)
        {
            if (readResult == null)
                throw new ArgumentNullException(nameof(readResult));
            if (newOutlineXml == null)
                throw new ArgumentNullException(nameof(newOutlineXml));
            if (string.IsNullOrEmpty(readResult.PageId))
                throw new ArgumentException("PageId is required", nameof(readResult));
            if (!readResult.TargetObjectIds.Any())
                throw new ArgumentException("TargetObjectIds is required", nameof(readResult));

            return WriteOutlinesAsync(readResult, new[] { BuildUpdatedOutline(readResult, newOutlineXml) });
        }

        public async Task WriteOutlinesAsync(ReadResult readResult, IReadOnlyList<XElement> outlines)
        {
            string xml = BuildPageChanges(readResult, outlines);
            if (_logger != null) await _logger.LogPageChangesAsync(xml).ConfigureAwait(false);
            await _comDispatcher.InvokeAsync(app => UpdatePageContent(app, readResult.PageId,
                (string)outlines[0].Attribute("objectID"), xml)).ConfigureAwait(false);
        }

        internal Task InsertContentAsync(XElement page, XElement converted)
        {
            var outline = OneNotePasteInsertion.CreateOutline(page, converted);
            var read = new ReadResult
            {
                PageId = (string)page.Attribute("ID"),
                PageQuickStyleDefinitions = page.Elements(_ns + "QuickStyleDef").ToList()
            };
            return _comDispatcher.InvokeAsync(app =>
            {
                OneNotePasteInsertion.MaterializeBinaryContent(outline, callbackId =>
                {
                    app.GetBinaryPageContent(read.PageId, callbackId, out string data);
                    return data;
                });
                ValidatePasteTarget(app, page, outline);
                string xml = BuildPageChanges(read, new[] { outline }, insertion: true);
                UpdatePageContent(app, read.PageId, (string)outline.Attribute("objectID"), xml, selectOutline: false);
            });
        }

        private void ValidatePasteTarget(OneNoteInterop.Application app, XElement originalPage, XElement outline)
        {
            string pageId = (string)originalPage.Attribute("ID");
            string outlineId = (string)outline.Attribute("objectID");
            if (string.IsNullOrEmpty(outlineId)) return;
            var original = originalPage.Elements(_ns + "Outline").First(e => (string)e.Attribute("objectID") == outlineId);
            app.GetPageContent(pageId, out string currentXml, OneNoteInterop.PageInfo.piSelection, OneNoteInterop.XMLSchema.xs2013);
            var current = XElement.Parse(currentXml).Elements(_ns + "Outline").FirstOrDefault(e => (string)e.Attribute("objectID") == outlineId);
            if (!OneNotePasteInsertion.ContentEquals(original, current))
                throw new System.Runtime.InteropServices.COMException("Paste target changed during rendering.", unchecked((int)0x80042010));
        }

        private string BuildPageChanges(ReadResult readResult, IReadOnlyList<XElement> outlines, bool insertion = false)
        {
            if (readResult == null || string.IsNullOrEmpty(readResult.PageId))
                throw new ArgumentException("A page ID is required.", nameof(readResult));
            if (outlines == null || outlines.Count == 0)
                throw new ArgumentException("At least one outline is required.", nameof(outlines));
            foreach (var outline in outlines)
                if (outline.Name != _ns + "Outline" || (!insertion && string.IsNullOrEmpty((string)outline.Attribute("objectID"))))
                    throw new ArgumentException("Every outline must have its original identity.", nameof(outlines));

            bool needsTodoTagDef = outlines.Any(ContainsTodoTagIndexZero);
            var hyperlinkQuickStyles = new System.Collections.Generic.List<XElement>();
            foreach (var outline in outlines)
            {
                hyperlinkQuickStyles.AddRange(PrepareHyperlinkQuickStyles(outline,
                    readResult.PageQuickStyleDefinitions.Concat(hyperlinkQuickStyles).ToList()));
            }
            var page = new XElement(_ns + "Page",
                new XAttribute("ID", readResult.PageId),
                // Emit explicit prefix to match OneNote's typical serialization.
                new XAttribute(XNamespace.Xmlns + "one", _ns));

            if (needsTodoTagDef)
            {
                // OneNote validates Tag nodes against TagDef definitions during UpdatePageContent.
                // Empirically, some pages fail updates if TagDef(index="0") is not present in the *changes XML*,
                // even when the page already contains the definition. Always include it when we emit <Tag index="0">.
                using (PerformanceTraceContext.Measure("Write.AddTodoTagDef"))
                {
                    var pageDoc = new XDocument(page);
                    EnsureTagDefExists(pageDoc, _ns);
                }
            }

            page.Add(hyperlinkQuickStyles);

            if (!insertion) OneNoteWriteSelection.PlaceCaret(outlines);
            page.Add(outlines);

            string updatedXml;
            using (PerformanceTraceContext.Measure("Write.Serialize.PageChangesXml"))
            {
                updatedXml = new XDocument(page).ToString();
            }

            PerformanceTraceContext.AddMetric("Write.PageChangesXmlChars", (updatedXml?.Length ?? 0).ToString());
            return updatedXml;
        }

        private void UpdatePageContent(OneNoteInterop.Application app, string pageId, string outlineId, string updatedXml,
            bool selectOutline = true)
        {
            try
            {
                if (selectOutline)
                {
                    using (PerformanceTraceContext.Measure("Write.OneNote.SelectOutline"))
                        if (!string.IsNullOrEmpty(outlineId)) app.NavigateTo(pageId, outlineId, false);
                }
                using (PerformanceTraceContext.Measure("Write.OneNote.UpdatePageContent", "xs2013/force=true"))
                {
                    app.UpdatePageContent(updatedXml, DateTime.MinValue, OneNoteInterop.XMLSchema.xs2013, true);
                }
            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                var userMessage = ErrorMessages.GetUserFriendlyMessage(comEx);
                var technicalMessage = $"UpdatePageContent failed. HResult=0x{comEx.HResult:X}. {comEx.Message}";
                throw new ContentWriteException(userMessage, technicalMessage, comEx);
            }
            catch (Exception ex)
            {
                var userMessage = string.Format(Resources.GetString("Error_UpdatePageFailed"), ex.Message);
                var technicalMessage = $"UpdatePageContent failed unexpectedly. {ex.GetType().Name}: {ex.Message}";
                throw new ContentWriteException(userMessage, technicalMessage, ex);
            }
        }

        private XElement BuildUpdatedOutline(ReadResult readResult, XElement newOutlineXml)
        {
            if (readResult == null)
            {
                throw new ArgumentNullException(nameof(readResult));
            }

            if (newOutlineXml == null)
            {
                throw new ArgumentNullException(nameof(newOutlineXml));
            }

            if (readResult.Mode == DetectionMode.Cursor)
            {
                return BuildUpdatedOutlineForCursor(readResult, newOutlineXml);
            }

            return BuildUpdatedOutlineForSelection(readResult, newOutlineXml);
        }

        private XElement BuildUpdatedOutlineForCursor(ReadResult readResult, XElement newOutlineXml)
        {
            // Cursor mode: replace the entire Outline identified by TargetObjectIds[0].
            if (readResult.OriginalXmlNode != null)
            {
                PreserveAttributes(newOutlineXml, readResult.OriginalXmlNode);
            }
            newOutlineXml.SetAttributeValue("objectID", readResult.TargetObjectIds.First());
            RemoveSelectedAttributes(newOutlineXml);
            return newOutlineXml;
        }

        private XElement BuildUpdatedOutlineForSelection(ReadResult readResult, XElement newOutlineXml)
        {
            // Selection mode: merge the converted selection back into the original Outline
            // so non-selected siblings remain untouched.
            var reference = readResult.OriginalXmlNode ?? readResult.OriginalXmlNodes.FirstOrDefault();
            var sourceOutline = reference?.Ancestors(_ns + "Outline").FirstOrDefault();
            if (sourceOutline == null)
            {
                throw new InvalidOperationException("Cannot locate source Outline for selection update.");
            }

            var updatedOutline = new XElement(sourceOutline);
            RemoveSelectedAttributes(updatedOutline);

            bool isNewContentOutline = newOutlineXml.Name.LocalName == "Outline";
            var newOEChildren = isNewContentOutline
                ? newOutlineXml.Element(_ns + "OEChildren")?.Elements(_ns + "OE").Select(oe => new XElement(oe)).ToList()
                    ?? new System.Collections.Generic.List<XElement>()
                : new System.Collections.Generic.List<XElement> { new XElement(newOutlineXml) };

            var firstId = readResult.TargetObjectIds.First();
            var firstTarget = updatedOutline.Descendants(_ns + "OE")
                .FirstOrDefault(e => string.Equals((string)e.Attribute("objectID"), firstId, StringComparison.Ordinal));
            if (firstTarget == null)
            {
                throw new InvalidOperationException($"Cannot find selection target OE with ObjectID: {firstId}");
            }

            using (PerformanceTraceContext.Measure("Write.Selection.ApplyReplacement", $"targets={readResult.TargetObjectIds.Count}"))
            {
                if (newOEChildren.Any())
                {
                    firstTarget.ReplaceWith(newOEChildren);
                }
                else
                {
                    firstTarget.Remove();
                }

                foreach (var id in readResult.TargetObjectIds.Skip(1))
                {
                    var nodeToRemove = updatedOutline.Descendants(_ns + "OE")
                        .FirstOrDefault(e => string.Equals((string)e.Attribute("objectID"), id, StringComparison.Ordinal));
                    nodeToRemove?.Remove();
                }
            }

            return updatedOutline;
        }

        private static void RemoveSelectedAttributes(XElement element)
        {
            if (element == null)
            {
                return;
            }

            foreach (var node in element.DescendantsAndSelf())
            {
                node.Attribute("selected")?.Remove();
            }
        }

        private bool ContainsTodoTagIndexZero(XElement outline)
        {
            if (outline == null)
            {
                return false;
            }

            return outline.Descendants(_ns + "Tag")
                .Any(tag => string.Equals((string)tag.Attribute("index"), "0", StringComparison.Ordinal));
        }

        private IReadOnlyList<XElement> PrepareHyperlinkQuickStyles(
            XElement outline,
            IReadOnlyCollection<XElement> pageQuickStyles)
        {
            var availableStyles = (pageQuickStyles ?? Array.Empty<XElement>())
                .Select(element => new XElement(element))
                .ToList();
            var usedStyles = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);
            int nextIndex = GetNextQuickStyleIndex(availableStyles, outline);

            foreach (var oe in outline.Descendants(_ns + "OE"))
            {
                if (!OneNoteHyperlinkColorWorkaround.TryGetExplicitLinkColor(
                        oe,
                        _ns,
                        out string linkColor))
                {
                    continue;
                }

                if (!usedStyles.TryGetValue(linkColor, out XElement quickStyle))
                {
                    quickStyle = FindReusableColorStyle(availableStyles, linkColor) ??
                        CreateColorQuickStyle(availableStyles, nextIndex++, linkColor);
                    usedStyles.Add(linkColor, quickStyle);
                }

                oe.SetAttributeValue("quickStyleIndex", (string)quickStyle.Attribute("index"));
                var firstText = oe.Elements(_ns + "T").FirstOrDefault();
                if (firstText != null)
                {
                    firstText.AddBeforeSelf(new XElement(_ns + "T", new XCData(string.Empty)));
                }
            }

            return usedStyles.Values.Select(element => new XElement(element)).ToList();
        }

        private int GetNextQuickStyleIndex(IEnumerable<XElement> styles, XElement outline)
        {
            var indices = styles
                .Select(style => (string)style.Attribute("index"))
                .Concat(outline.Descendants(_ns + "OE")
                    .Select(oe => (string)oe.Attribute("quickStyleIndex")))
                .Select(value => int.TryParse(value, out int index) ? (int?)index : null)
                .Where(index => index.HasValue)
                .Select(index => index.Value)
                .ToList();

            return indices.Count == 0 ? 6 : System.Math.Max(6, indices.Max() + 1);
        }

        private static XElement FindReusableColorStyle(
            IEnumerable<XElement> styles,
            string color)
        {
            return styles.FirstOrDefault(style =>
                ((string)style.Attribute("name") ?? string.Empty)
                    .StartsWith("TeXShiftColor", StringComparison.OrdinalIgnoreCase) &&
                CssColorParser.TryNormalize((string)style.Attribute("fontColor"), out string styleColor) &&
                string.Equals(styleColor, color, StringComparison.OrdinalIgnoreCase));
        }

        private XElement CreateColorQuickStyle(
            IReadOnlyCollection<XElement> availableStyles,
            int index,
            string color)
        {
            var paragraphStyle = availableStyles.FirstOrDefault(style =>
                string.Equals(
                    (string)style.Attribute("name"),
                    "p",
                    StringComparison.OrdinalIgnoreCase));

            XElement quickStyle = paragraphStyle != null
                ? new XElement(paragraphStyle)
                : new XElement(
                    _ns + "QuickStyleDef",
                    new XAttribute("highlightColor", "automatic"),
                    new XAttribute("font", "微软雅黑"),
                    new XAttribute("fontSize", "11.0"),
                    new XAttribute("spaceBefore", "0.0"),
                    new XAttribute("spaceAfter", "0.0"));

            quickStyle.SetAttributeValue("index", index.ToString());
            quickStyle.SetAttributeValue("name", "TeXShiftColor" + index);
            quickStyle.SetAttributeValue("fontColor", color);
            return quickStyle;
        }

        /// <summary>
        /// Preserves important attributes from the original node.
        /// </summary>
        private void PreserveAttributes(XElement newNode, XElement originalNode)
        {
            // Preserve position attributes if they exist
            var positionAttributes = new[] { "lastModifiedTime", "author", "authorInitials", "authorResolutionID" };

            foreach (var attrName in positionAttributes)
            {
                var attr = originalNode.Attribute(attrName);
                if (attr != null && newNode.Attribute(attrName) == null)
                {
                    newNode.SetAttributeValue(attrName, attr.Value);
                }
            }

            // Preserve position and size for Outline nodes
            if (originalNode.Name.LocalName == "Outline")
            {
                var positionAttr = originalNode.Elements()
                    .FirstOrDefault(e => e.Name.LocalName == "Position");
                var sizeAttr = originalNode.Elements()
                    .FirstOrDefault(e => e.Name.LocalName == "Size");

                if (positionAttr != null)
                {
                    var existingPosition = newNode.Elements()
                        .FirstOrDefault(e => e.Name.LocalName == "Position");
                    if (existingPosition != null)
                        existingPosition.Remove();
                    newNode.AddFirst(new XElement(positionAttr));
                }

                if (sizeAttr != null)
                {
                    var existingSize = newNode.Elements()
                        .FirstOrDefault(e => e.Name.LocalName == "Size");
                    if (existingSize != null)
                        existingSize.Remove();

                    // Insert Size after Position if it exists, otherwise at the beginning
                    var position = newNode.Elements().FirstOrDefault(e => e.Name.LocalName == "Position");
                    if (position != null)
                        position.AddAfterSelf(new XElement(sizeAttr));
                    else
                        newNode.AddFirst(new XElement(sizeAttr));
                }
            }
        }

        /// <summary>
        /// Ensures that the page has a TagDef for task list checkboxes (index="0").
        /// If it doesn't exist, adds it to the page root.
        /// </summary>
        private void EnsureTagDefExists(XDocument doc, XNamespace ns)
        {
            var pageRoot = doc.Root;
            if (pageRoot == null) return;

            // Check if TagDef with index="0" already exists
            var existingTagDef = pageRoot.Elements(ns + "TagDef")
                .FirstOrDefault(e => e.Attribute("index")?.Value == "0");

            if (existingTagDef == null)
            {
                // Create a new TagDef for task list checkboxes
                // type="0" means checkbox, symbol="3" is the checkbox icon
                var tagDef = new XElement(ns + "TagDef",
                    new XAttribute("index", "0"),
                    new XAttribute("type", "0"),
                    new XAttribute("symbol", "3"),
                    new XAttribute("fontColor", "automatic"),
                    new XAttribute("highlightColor", "none"),
                    new XAttribute("name", Resources.GetString("OneNote_TodoTag")));

                // Insert TagDef at the beginning of the page (after xmlns declarations)
                // It should come before QuickStyleDef, PageSettings, and other page-level elements
                pageRoot.AddFirst(tagDef);
            }
        }
    }
}
