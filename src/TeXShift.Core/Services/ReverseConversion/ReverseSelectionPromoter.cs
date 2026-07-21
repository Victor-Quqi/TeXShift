using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TeXShift.Core.OneNote;
using TeXShift.Core.OneNoteMeta;

namespace TeXShift.Core.Services.ReverseConversion
{
    internal static class ReverseSelectionPromoter
    {
        internal static bool TryPromoteSelectionToOutlineWithValidMeta(ReadResult readResult, out ReadResult promoted)
        {
            promoted = null;
            if (readResult == null || !readResult.IsSuccess)
            {
                return false;
            }

            if (readResult.Mode != DetectionMode.Selection)
            {
                return false;
            }

            // Respect real selections. Only promote when OneNote reports a "selection" due to the caret being inside
            // an embedded object (e.g., Math/Image), where selection XML does not include the outline-level TeXShift meta.
            List<XElement> selectionRoots;
            if (readResult.OriginalXmlNodes != null && readResult.OriginalXmlNodes.Count > 0)
            {
                selectionRoots = readResult.OriginalXmlNodes;
            }
            else if (readResult.OriginalXmlNode != null)
            {
                selectionRoots = new List<XElement> { readResult.OriginalXmlNode };
            }
            else
            {
                return false;
            }

            if (selectionRoots.Count != 1)
            {
                return false;
            }

            if (!IsLikelyEmbeddedObjectCaretSelection(selectionRoots[0]))
            {
                return false;
            }

            var ns = selectionRoots[0].Name.Namespace;
            var outlineCandidates = selectionRoots
                .Select(e => e.Ancestors(ns + "Outline").FirstOrDefault())
                .Where(o => o != null)
                .GroupBy(o => (string)o.Attribute("objectID") ?? string.Empty)
                .ToList();

            // Avoid accidentally converting only one outline when selection spans multiple outlines.
            if (outlineCandidates.Count != 1)
            {
                return false;
            }

            var outline = outlineCandidates[0].FirstOrDefault();
            if (outline == null)
            {
                return false;
            }

            string outlineObjectId = outline.Attribute("objectID")?.Value;
            if (string.IsNullOrWhiteSpace(outlineObjectId))
            {
                return false;
            }

            var meta = TeXShiftMetaReader.ReadOutline(outline);
            if (meta == null || !meta.HasTeXShiftMeta || !meta.IsValid)
            {
                return false;
            }

            promoted = new ReadResult
            {
                IsSuccess = true,
                Mode = DetectionMode.Cursor,
                PageId = readResult.PageId,
                OriginalXmlNode = outline,
                SourceOutlineWidth = readResult.SourceOutlineWidth
            };
            promoted.TargetObjectIds.Add(outlineObjectId);
            return true;
        }

        internal static bool TryPromoteSelectionToContainingTable(ReadResult readResult, out ReadResult promoted)
        {
            promoted = null;
            if (readResult == null || !readResult.IsSuccess)
            {
                return false;
            }

            if (readResult.Mode != DetectionMode.Selection)
            {
                return false;
            }

            if (readResult.OriginalXmlNodes == null || readResult.OriginalXmlNodes.Count == 0)
            {
                return false;
            }

            var ns = readResult.OriginalXmlNodes[0].Name.Namespace;

            // Replace any table-internal OEs with the table container OE to avoid leaving
            // <Cell><OEChildren/></Cell> empty after write-back (OneNote rejects empty OEChildren).
            var mapped = readResult.OriginalXmlNodes
                .Select(oe => FindTableContainerOE(oe, ns) ?? oe)
                .ToList();

            if (!mapped.Any(e => e != null && e.Element(ns + "Table") != null))
            {
                return false;
            }

            // Keep order stable (selection order) while removing duplicates.
            var distinct = new List<XElement>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (var oe in mapped)
            {
                var id = (string)oe?.Attribute("objectID");
                if (string.IsNullOrWhiteSpace(id))
                {
                    continue;
                }

                if (seen.Add(id))
                {
                    distinct.Add(oe);
                }
            }

            if (distinct.Count == 0)
            {
                return false;
            }

            promoted = new ReadResult
            {
                IsSuccess = true,
                Mode = DetectionMode.Selection,
                ExtractedText = readResult.ExtractedText,
                PageId = readResult.PageId,
                OriginalXmlNode = distinct[0],
                OriginalXmlNodes = distinct,
                SourceOutlineWidth = readResult.SourceOutlineWidth,
                PageHasTodoTagDef = readResult.PageHasTodoTagDef
            };

            foreach (var oe in distinct)
            {
                var id = (string)oe.Attribute("objectID");
                if (!string.IsNullOrWhiteSpace(id))
                {
                    promoted.TargetObjectIds.Add(id);
                }
            }

            return promoted.TargetObjectIds.Count > 0;
        }

        internal static bool TryPromoteFullOutlineSelectionToCursor(
            ReadResult readResult,
            bool requireValidMeta,
            bool preservePageHasTodoTagDef,
            out ReadResult promoted)
        {
            promoted = null;
            if (readResult == null || !readResult.IsSuccess)
            {
                return false;
            }

            if (readResult.Mode != DetectionMode.Selection)
            {
                return false;
            }

            if (readResult.OriginalXmlNodes == null || readResult.OriginalXmlNodes.Count == 0)
            {
                return false;
            }

            var firstOe = readResult.OriginalXmlNodes.FirstOrDefault();
            if (firstOe == null)
            {
                return false;
            }

            var ns = firstOe.Name.Namespace;
            var outline = firstOe.Ancestors(ns + "Outline").FirstOrDefault();
            if (outline == null)
            {
                return false;
            }

            string outlineObjectId = outline.Attribute("objectID")?.Value;
            if (string.IsNullOrWhiteSpace(outlineObjectId))
            {
                return false;
            }

            var rootChildren = outline.Element(ns + "OEChildren");
            if (rootChildren == null)
            {
                return false;
            }

            var rootOes = rootChildren.Elements(ns + "OE").ToList();
            if (rootOes.Count == 0)
            {
                return false;
            }

            var selectedRootOes = readResult.OriginalXmlNodes
                .Select(oe => FindRootOeUnderOutline(oe, rootChildren, ns))
                .Where(oe => oe != null)
                .Distinct()
                .ToList();

            if (selectedRootOes.Count != rootOes.Count)
            {
                return false;
            }

            if (!rootOes.All(selectedRootOes.Contains))
            {
                return false;
            }

            if (requireValidMeta)
            {
                var meta = TeXShiftMetaReader.ReadOutline(outline);
                if (meta == null || !meta.HasTeXShiftMeta || !meta.IsValid)
                {
                    return false;
                }
            }

            promoted = new ReadResult
            {
                IsSuccess = true,
                Mode = DetectionMode.Cursor,
                ExtractedText = readResult.ExtractedText,
                PageId = readResult.PageId,
                OriginalXmlNode = outline,
                SourceOutlineWidth = readResult.SourceOutlineWidth,
                PageHasTodoTagDef = preservePageHasTodoTagDef && readResult.PageHasTodoTagDef
            };
            promoted.TargetObjectIds.Add(outlineObjectId);
            return true;
        }

        private static bool IsLikelyEmbeddedObjectCaretSelection(XElement oe)
        {
            if (oe == null)
            {
                return false;
            }

            var ns = oe.Name.Namespace;

            // Only treat it as "object click" when the selected content is a mathML payload or an embedded image.
            // This avoids promoting normal text selections to outline scope.
            var selectedT = oe.Descendants(ns + "T").Where(t => t.Attribute("selected") != null).ToList();
            if (selectedT.Count == 1)
            {
                var text = selectedT[0].Value ?? string.Empty;
                if (text.IndexOf("<!--[if mathML]", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return oe.Descendants().Any(e =>
                e.Attribute("selected") != null
                && string.Equals(e.Name.LocalName, "Image", StringComparison.OrdinalIgnoreCase));
        }

        private static XElement FindRootOeUnderOutline(XElement oe, XElement rootChildren, XNamespace ns)
        {
            if (oe == null || rootChildren == null)
            {
                return null;
            }

            // Find the root-level OE under the outline's root OEChildren (works for tables and nested OEs).
            return oe.AncestorsAndSelf(ns + "OE").FirstOrDefault(e => e.Parent == rootChildren);
        }

        private static XElement FindTableContainerOE(XElement oe, XNamespace ns)
        {
            if (oe == null)
            {
                return null;
            }

            // Already at table-container level.
            if (oe.Element(ns + "Table") != null)
            {
                return oe;
            }

            var table = oe.Ancestors(ns + "Table").FirstOrDefault();
            if (table == null)
            {
                return null;
            }

            // OneNote tables are stored as <OE><Table>...</Table></OE>.
            var container = table.Parent;
            if (container == null || container.Name != ns + "OE")
            {
                return null;
            }

            return container;
        }
    }
}
