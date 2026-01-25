using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using TeXShift.Core.OneNoteToMarkdown.Abstractions;
using TeXShift.Core.Utils;

namespace TeXShift.Core.OneNoteToMarkdown.Handlers
{
    /// <summary>
    /// Converts OneNote tables to Markdown tables.
    /// </summary>
    internal sealed partial class TableElementHandler : IElementHandler
    {
        public bool CanHandle(XElement element, IOneNoteConverterContext context)
        {
            var table = OneNoteTableHelpers.GetTable(element, context);
            if (table == null)
            {
                return false;
            }

            if (!OneNoteTableHelpers.IsTrue(table.Attribute("bordersVisible")?.Value))
            {
                return false;
            }

            return true;
        }

        public IEnumerable<string> Handle(XElement element, IOneNoteConverterContext context)
        {
            var ns = context.OneNoteNamespace;
            var table = OneNoteTableHelpers.GetTable(element, context);
            if (table == null)
            {
                yield break;
            }

            var rows = table.Elements(ns + "Row").ToList();
            if (rows.Count == 0)
            {
                yield break;
            }

            int columnCount = rows.Select(r => r.Elements(ns + "Cell").Count()).DefaultIfEmpty(0).Max();
            if (columnCount == 0)
            {
                yield break;
            }

            var alignments = GetColumnAlignments(rows, columnCount, context);
            var sb = new System.Text.StringBuilder();

            bool oneNoteHasHeaderRow = OneNoteTableHelpers.IsTrue(table.Attribute("hasHeaderRow")?.Value);
            bool useHeaderRow = oneNoteHasHeaderRow || context.StyleConfig.TryRecognizeNonTeXShiftFormats;

            if (!useHeaderRow && FirstRowLooksLikeHeader(rows[0], columnCount, context))
            {
                useHeaderRow = true;
            }

            if (useHeaderRow)
            {
                var headerCells = ExtractRowCells(rows[0], columnCount, context, isHeader: true);
                sb.AppendLine(BuildRow(headerCells));
                sb.AppendLine(BuildSeparator(alignments));

                for (int i = 1; i < rows.Count; i++)
                {
                    var bodyCells = ExtractRowCells(rows[i], columnCount, context, isHeader: false);
                    sb.AppendLine(BuildRow(bodyCells));
                }
            }
            else
            {
                // Markdown tables require a header row. When OneNote doesn't mark a header,
                // output an empty header to avoid guessing.
                sb.AppendLine(BuildRow(Enumerable.Repeat(string.Empty, columnCount).ToList()));
                sb.AppendLine(BuildSeparator(alignments));

                for (int i = 0; i < rows.Count; i++)
                {
                    var bodyCells = ExtractRowCells(rows[i], columnCount, context, isHeader: false);
                    sb.AppendLine(BuildRow(bodyCells));
                }
            }

            yield return sb.ToString().TrimEnd();
        }

        private static bool FirstRowLooksLikeHeader(XElement row, int columnCount, IOneNoteConverterContext context)
        {
            if (row == null || context == null)
            {
                return false;
            }

            var ns = context.OneNoteNamespace;
            var cells = row.Elements(ns + "Cell").ToList();

            bool hasAnyContent = false;
            for (int i = 0; i < columnCount; i++)
            {
                var cell = i < cells.Count ? cells[i] : null;
                var html = ExtractCellHtml(cell, context);
                if (!HasSignificantTextInHtml(html))
                {
                    continue;
                }

                hasAnyContent = true;
                if (!IsHtmlFullyBold(html))
                {
                    return false;
                }
            }

            return hasAnyContent;
        }

        private static string ExtractCellHtml(XElement cell, IOneNoteConverterContext context)
        {
            if (cell == null || context == null)
            {
                return string.Empty;
            }

            var ns = context.OneNoteNamespace;
            var oeChildren = cell.Element(ns + "OEChildren");
            if (oeChildren == null)
            {
                return string.Empty;
            }

            var oe = oeChildren.Elements(ns + "OE").FirstOrDefault();
            if (oe == null)
            {
                return string.Empty;
            }

            var tElements = oe.Elements(ns + "T").ToList();
            if (tElements.Count == 0)
            {
                return string.Empty;
            }

            return string.Concat(tElements.Select(t => t.Value ?? string.Empty));
        }

        private static List<string> ExtractRowCells(XElement row, int columnCount, IOneNoteConverterContext context, bool isHeader)
        {
            var ns = context.OneNoteNamespace;
            var cells = row.Elements(ns + "Cell").ToList();
            var results = new List<string>(columnCount);

            for (int i = 0; i < columnCount; i++)
            {
                var cell = i < cells.Count ? cells[i] : null;
                var cellText = ExtractCellText(cell, context, isHeader);
                results.Add(EscapeCell(cellText));
            }

            return results;
        }

        private static string ExtractCellText(XElement cell, IOneNoteConverterContext context, bool isHeader)
        {
            if (cell == null)
            {
                return string.Empty;
            }

            var ns = context.OneNoteNamespace;
            var oeChildren = cell.Element(ns + "OEChildren");
            if (oeChildren == null)
            {
                return string.Empty;
            }

            var oe = oeChildren.Elements(ns + "OE").FirstOrDefault();
            if (oe == null)
            {
                return string.Empty;
            }

            var tElements = oe.Elements(ns + "T").ToList();
            if (tElements.Count == 0)
            {
                return string.Empty;
            }

            var html = string.Concat(tElements.Select(t => t.Value ?? string.Empty));
            if (isHeader && TryUnwrapHeaderSpan(html, out var innerHtml))
            {
                html = innerHtml;
            }

            var parsed = context.ParseInlineHtml(html);
            return NormalizeCellText(parsed);
        }

        private static string NormalizeCellText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            var normalized = text.Replace("\r\n", "\n").Replace("\r", "\n");
            if (normalized.IndexOf('\n') >= 0)
            {
                normalized = string.Join("<br>", normalized.Split(new[] { '\n' }, StringSplitOptions.None));
            }

            return normalized.Trim();
        }

        private static string EscapeCell(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            return text.Replace("|", "\\|");
        }

        private static string BuildRow(IReadOnlyList<string> cells)
        {
            return "| " + string.Join(" | ", cells) + " |";
        }

        private static string BuildSeparator(IReadOnlyList<string> alignments)
        {
            var parts = new List<string>(alignments.Count);
            foreach (var alignment in alignments)
            {
                switch (alignment)
                {
                    case "center":
                        parts.Add(":---:");
                        break;
                    case "right":
                        parts.Add("---:");
                        break;
                    default:
                        parts.Add("---");
                        break;
                }
            }

            return "| " + string.Join(" | ", parts) + " |";
        }

        private static List<string> GetColumnAlignments(List<XElement> rows, int columnCount, IOneNoteConverterContext context)
        {
            var ns = context.OneNoteNamespace;
            var alignments = new List<string>(columnCount);

            for (int i = 0; i < columnCount; i++)
            {
                string alignment = "left";
                foreach (var row in rows)
                {
                    var cell = row.Elements(ns + "Cell").ElementAtOrDefault(i);
                    var oe = cell?.Element(ns + "OEChildren")?.Element(ns + "OE");
                    var candidate = GetAlignmentFromOe(oe);
                    if (!string.IsNullOrEmpty(candidate))
                    {
                        alignment = candidate;
                        break;
                    }
                }
                alignments.Add(alignment);
            }

            return alignments;
        }

        private static string GetAlignmentFromOe(XElement oe)
        {
            if (oe == null)
            {
                return null;
            }

            var alignment = oe.Attribute("alignment")?.Value;
            if (!string.IsNullOrEmpty(alignment))
            {
                alignment = alignment.Trim().ToLowerInvariant();
                if (alignment == "left" || alignment == "center" || alignment == "right")
                {
                    return alignment;
                }
            }

            var style = oe.Attribute("style")?.Value ?? string.Empty;
            if (style.IndexOf("text-align:center", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "center";
            }
            if (style.IndexOf("text-align:right", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "right";
            }
            if (style.IndexOf("text-align:left", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "left";
            }

            return null;
        }
    }
}
