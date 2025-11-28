using System;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using TeXShift.Core.Abstractions;
using OneNoteInterop = Microsoft.Office.Interop.OneNote;

namespace TeXShift.Core.OneNote
{
    /// <summary>
    /// Handles writing converted content back to OneNote pages.
    /// </summary>
    public class OneNotePageWriter : IContentWriter
    {
        private readonly OneNoteInterop.Application _oneNoteApp;
        private readonly XNamespace _ns = "http://schemas.microsoft.com/office/onenote/2013/onenote";

        public OneNotePageWriter(OneNoteInterop.Application oneNoteApp)
        {
            _oneNoteApp = oneNoteApp;
        }

        /// <summary>
        /// Asynchronously replaces content in OneNote based on the read result and converted XML.
        /// </summary>
        /// <param name="readResult">The original read result containing metadata</param>
        /// <param name="newOutlineXml">The new Outline XML element to insert</param>
        public async Task ReplaceContentAsync(ReadResult readResult, XElement newOutlineXml)
        {
            // Wrap COM calls in Task.Run to avoid blocking UI thread
            await Task.Run(() => ReplaceContent(readResult, newOutlineXml)).ConfigureAwait(false);
        }

        /// <summary>
        /// Replaces content in OneNote based on the read result and converted XML.
        /// (Synchronous version - kept for internal use)
        /// </summary>
        /// <param name="readResult">The original read result containing metadata</param>
        /// <param name="newOutlineXml">The new Outline XML element to insert</param>
        private void ReplaceContent(ReadResult readResult, XElement newOutlineXml)
        {
            if (readResult == null)
                throw new ArgumentNullException(nameof(readResult));
            if (newOutlineXml == null)
                throw new ArgumentNullException(nameof(newOutlineXml));
            if (string.IsNullOrEmpty(readResult.PageId))
                throw new ArgumentException("PageId is required", nameof(readResult));
            if (!readResult.TargetObjectIds.Any())
                throw new ArgumentException("TargetObjectIds is required", nameof(readResult));

            string pageXml;
            _oneNoteApp.GetPageContent(readResult.PageId, out pageXml, OneNoteInterop.PageInfo.piAll, OneNoteInterop.XMLSchema.xs2013);

            var doc = XDocument.Parse(pageXml);
            var ns = doc.Root.Name.Namespace;

            // Ensure TagDef exists for task lists (Tag elements with index="0")
            EnsureTagDefExists(doc, ns);

            var targetNodes = readResult.TargetObjectIds
                .Select(id => FindNodeByObjectId(doc, id, ns))
                .Where(node => node != null)
                .ToList();

            if (!targetNodes.Any())
            {
                throw new InvalidOperationException($"Cannot find any target nodes with ObjectIDs: {string.Join(", ", readResult.TargetObjectIds)}");
            }

            var firstTargetNode = targetNodes.First();

            bool isNewContentOutline = newOutlineXml.Name.LocalName == "Outline";

            if (readResult.Mode == DetectionMode.Cursor)
            {
                // Cursor mode: Replace the entire Outline element
                PreserveAttributes(newOutlineXml, firstTargetNode);
                newOutlineXml.SetAttributeValue("objectID", readResult.TargetObjectIds.First());
                firstTargetNode.ReplaceWith(newOutlineXml);
            }
            else // Selection mode: Replace OEs, not the entire Outline
            {
                var newOEChildren = isNewContentOutline
                    ? newOutlineXml.Descendants(ns + "OE").ToList()
                    : new System.Collections.Generic.List<XElement> { newOutlineXml };

                if (newOEChildren.Any())
                {
                    firstTargetNode.ReplaceWith(newOEChildren);
                }
                else
                {
                    firstTargetNode.Remove();
                }

                foreach (var nodeToRemove in targetNodes.Skip(1))
                {
                    nodeToRemove.Remove();
                }
            }

            string updatedXml = doc.ToString();
            try
            {
                _oneNoteApp.UpdatePageContent(updatedXml, DateTime.MinValue, OneNoteInterop.XMLSchema.xs2013, true);
            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                throw new InvalidOperationException(
                    $"无法更新 OneNote 页面内容。可能的原因：页面被锁定、权限不足或 OneNote 未响应。\n\nCOM 错误代码: 0x{comEx.HResult:X}",
                    comEx);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"更新 OneNote 页面内容时发生错误: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Finds a node in the XML document by its objectID attribute.
        /// </summary>
        private XElement FindNodeByObjectId(XDocument doc, string objectId, XNamespace ns)
        {
            return doc.Descendants()
                .FirstOrDefault(e => e.Attribute("objectID")?.Value == objectId);
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
                    new XAttribute("name", "待办事项"));

                // Insert TagDef at the beginning of the page (after xmlns declarations)
                // It should come before QuickStyleDef, PageSettings, and other page-level elements
                pageRoot.AddFirst(tagDef);
            }
        }
    }
}
