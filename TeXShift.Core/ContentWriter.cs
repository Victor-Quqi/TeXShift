using System;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace TeXShift.Core
{
    /// <summary>
    /// Handles writing converted content back to OneNote pages.
    /// </summary>
    public class ContentWriter : IContentWriter
    {
        private readonly OneNote.Application _oneNoteApp;
        private readonly XNamespace _ns = "http://schemas.microsoft.com/office/onenote/2013/onenote";

        public ContentWriter(OneNote.Application oneNoteApp)
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

            // Get current page XML
            string pageXml;
            _oneNoteApp.GetPageContent(readResult.PageId, out pageXml, OneNote.PageInfo.piAll, OneNote.XMLSchema.xs2013);

            var doc = XDocument.Parse(pageXml);
            var ns = doc.Root.Name.Namespace;

            // Find all target nodes to be removed/replaced
            var targetNodes = readResult.TargetObjectIds
                .Select(id => FindNodeByObjectId(doc, id, ns))
                .Where(node => node != null)
                .ToList();

            if (!targetNodes.Any())
            {
                throw new InvalidOperationException($"Cannot find any target nodes with ObjectIDs: {string.Join(", ", readResult.TargetObjectIds)}");
            }

            var firstTargetNode = targetNodes.First();

            // Smart replacement logic
            bool isNewContentOutline = newOutlineXml.Name.LocalName == "Outline";

            if (readResult.Mode == DetectionMode.Cursor)
            {
                // Case 1: Cursor mode, replacing a whole Outline.
                PreserveAttributes(newOutlineXml, firstTargetNode);
                newOutlineXml.SetAttributeValue("objectID", readResult.TargetObjectIds.First());
                firstTargetNode.ReplaceWith(newOutlineXml);
            }
            else // Selection mode
            {
                // In selection mode, we always replace OEs.
                var newOEChildren = isNewContentOutline
                    ? newOutlineXml.Descendants(ns + "OE").ToList()
                    : new System.Collections.Generic.List<XElement> { newOutlineXml };

                // Replace the first OE with all the new content
                if (newOEChildren.Any())
                {
                    firstTargetNode.ReplaceWith(newOEChildren);
                }
                else
                {
                    // If new content is empty, just remove the first node
                    firstTargetNode.Remove();
                }

                // Remove all other old OEs involved in the multi-line selection
                foreach (var nodeToRemove in targetNodes.Skip(1))
                {
                    nodeToRemove.Remove();
                }
            }

            // Update the page content
            string updatedXml = doc.ToString();
            _oneNoteApp.UpdatePageContent(updatedXml, DateTime.MinValue, OneNote.XMLSchema.xs2013, true);
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
    }
}
