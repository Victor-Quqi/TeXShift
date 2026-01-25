using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TeXShift.Core.OneNote;
using TeXShift.Core.OneNoteMeta;
using TeXShift.Core.Utils;

namespace TeXShift.Core.Services.ReverseConversion
{
    internal static class TeXShiftMetaTableSnippetRestorer
    {
        internal static bool TryRestoreSelectedTableFromTeXShiftMeta(ReadResult readResult, string renderedMarkdown, out string tableMarkdown)
        {
            tableMarkdown = null;
            if (readResult == null || !readResult.IsSuccess || string.IsNullOrWhiteSpace(renderedMarkdown))
            {
                return false;
            }

            if (readResult.Mode != DetectionMode.Selection)
            {
                return false;
            }

            // Only handle the common case: selection maps to a single table container OE (see TryPromoteSelectionToContainingTable).
            var selectionOes = readResult.OriginalXmlNodes != null && readResult.OriginalXmlNodes.Count > 0
                ? readResult.OriginalXmlNodes
                : (readResult.OriginalXmlNode != null ? new List<XElement> { readResult.OriginalXmlNode } : new List<XElement>());

            if (selectionOes.Count != 1)
            {
                return false;
            }

            var oe = selectionOes[0];
            if (oe == null)
            {
                return false;
            }

            var ns = oe.Name.Namespace;
            var table = oe.Element(ns + "Table");
            if (table == null)
            {
                return false;
            }

            var outline = oe.Ancestors(ns + "Outline").FirstOrDefault();
            if (outline == null)
            {
                return false;
            }

            var meta = TeXShiftMetaReader.ReadOutline(outline);
            if (meta == null || !meta.HasTeXShiftMeta || !meta.IsValid || string.IsNullOrEmpty(meta.Source))
            {
                return false;
            }

            if (!MarkdownTableExtractor.TryExtractSingleMarkdownTableKey(renderedMarkdown, out var renderedKey))
            {
                return false;
            }

            var blocks = MarkdownTableExtractor.ExtractMarkdownTables(meta.Source);
            if (blocks.Count == 0)
            {
                return false;
            }

            // Prefer positional matching (table order) to avoid ambiguity when multiple tables have identical contents.
            int index = GetTableIndexInOutline(outline, table, ns);
            if (index >= 0 && index < blocks.Count && string.Equals(blocks[index].Key, renderedKey, StringComparison.Ordinal))
            {
                tableMarkdown = blocks[index].Text;
                return true;
            }

            // Fallback: unique key match.
            var matches = blocks.Where(b => string.Equals(b.Key, renderedKey, StringComparison.Ordinal)).ToList();
            if (matches.Count == 1)
            {
                tableMarkdown = matches[0].Text;
                return true;
            }

            return false;
        }

        private static int GetTableIndexInOutline(XElement outline, XElement table, XNamespace ns)
        {
            if (outline == null || table == null)
            {
                return -1;
            }

            var targetId = (string)table.Attribute("objectID");
            var tables = outline.Descendants(ns + "Table").ToList();
            if (!string.IsNullOrWhiteSpace(targetId))
            {
                for (int i = 0; i < tables.Count; i++)
                {
                    if (string.Equals((string)tables[i].Attribute("objectID"), targetId, StringComparison.Ordinal))
                    {
                        return i;
                    }
                }
            }

            // Fallback: reference equality if objectID is missing.
            for (int i = 0; i < tables.Count; i++)
            {
                if (ReferenceEquals(tables[i], table))
                {
                    return i;
                }
            }

            return -1;
        }
    }
}

