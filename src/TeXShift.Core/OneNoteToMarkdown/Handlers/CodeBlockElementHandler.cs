using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using TeXShift.Core.OneNoteToMarkdown.Abstractions;
using TeXShift.Core.OneNoteToMarkdown.Inlines;

namespace TeXShift.Core.OneNoteToMarkdown.Handlers
{
    /// <summary>
    /// Converts OneNote code block tables to fenced code blocks.
    /// </summary>
    internal sealed class CodeBlockElementHandler : IElementHandler
    {
        public bool CanHandle(XElement element, IOneNoteConverterContext context)
        {
            var table = OneNoteTableHelpers.GetTable(element, context);
            return table != null && IsCodeBlockTable(table, context);
        }

        public IEnumerable<string> Handle(XElement element, IOneNoteConverterContext context)
        {
            var ns = context.OneNoteNamespace;
            var table = OneNoteTableHelpers.GetTable(element, context);
            if (table == null)
            {
                yield break;
            }

            var cell = table.Descendants(ns + "Cell").FirstOrDefault();
            if (cell == null)
            {
                yield break;
            }

            var oeChildren = cell.Element(ns + "OEChildren");
            if (oeChildren == null)
            {
                yield break;
            }

            var lines = new List<string>();
            foreach (var lineOe in oeChildren.Elements(ns + "OE"))
            {
                var line = ExtractLine(lineOe, context);
                lines.Add(string.IsNullOrWhiteSpace(line) ? string.Empty : line);
            }

            while (lines.Count > 0 && string.IsNullOrEmpty(lines[lines.Count - 1]))
            {
                lines.RemoveAt(lines.Count - 1);
            }

            var sb = new StringBuilder();
            sb.Append("```");
            sb.AppendLine();
            foreach (var line in lines)
            {
                sb.AppendLine(line);
            }
            sb.Append("```");

            yield return sb.ToString();
        }

        private static string ExtractLine(XElement lineOe, IOneNoteConverterContext context)
        {
            if (lineOe == null)
            {
                return string.Empty;
            }

            var ns = context.OneNoteNamespace;
            var tElements = lineOe.Elements(ns + "T").ToList();
            if (tElements.Count == 0)
            {
                return string.Empty;
            }

            var sb = new StringBuilder();
            foreach (var t in tElements)
            {
                var html = t.Value ?? string.Empty;
                // In code blocks, "&nbsp;" may be part of the code (e.g., HTML snippets), so preserve it as text.
                sb.Append(HtmlStripper.StripPreservingNbspEntity(html));
            }

            return sb.ToString();
        }

        private static bool IsCodeBlockTable(XElement table, IOneNoteConverterContext context)
        {
            if (table == null)
            {
                return false;
            }

            if (!OneNoteTableHelpers.IsFalse(table.Attribute("hasHeaderRow")?.Value))
            {
                return false;
            }

            var ns = context.OneNoteNamespace;
            if (!OneNoteTableHelpers.TryGetSingleCell(table, ns, out var cell))
            {
                return false;
            }

            var oeChildren = cell.Element(ns + "OEChildren");
            if (oeChildren == null)
            {
                return false;
            }

            bool strictMode = !context.StyleConfig.TryRecognizeNonTeXShiftFormats;
            if (strictMode)
            {
                // Strict mode: require "container-like" traits to reduce false positives.
                if (!OneNoteTableHelpers.IsFalse(table.Attribute("bordersVisible")?.Value))
                {
                    return false;
                }

                if (!OneNoteTableHelpers.IsLockedSingleColumn(table, ns))
                {
                    return false;
                }

                if (cell.Attribute("shadingColor") == null)
                {
                    return false;
                }
            }

            return LooksLikeCodeLines(oeChildren, context, strictMode);
        }

        private static bool LooksLikeCodeLines(XElement oeChildren, IOneNoteConverterContext context, bool strictMode)
        {
            var ns = context.OneNoteNamespace;
            var codeConfig = context.StyleConfig.GetCodeBlockStyle();
            var expectedFontFamily = codeConfig.FontFamily ?? string.Empty;
            var expectedSpaceBetween = codeConfig.SpaceBetween;

            var lineOes = oeChildren.Elements(ns + "OE").ToList();
            if (lineOes.Count == 0)
            {
                return false;
            }

            int matched = 0;
            foreach (var lineOe in lineOes)
            {
                if (strictMode && lineOe.Element(ns + "OEChildren") != null)
                {
                    return false;
                }

                if (!lineOe.Elements(ns + "T").Any())
                {
                    if (strictMode)
                    {
                        return false;
                    }
                    continue;
                }

                var style = lineOe.Attribute("style")?.Value ?? string.Empty;
                if (string.IsNullOrWhiteSpace(expectedFontFamily) ||
                    style.IndexOf(expectedFontFamily, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    if (strictMode)
                    {
                        return false;
                    }
                    continue;
                }

                if (!SpaceBetweenMatches(lineOe.Attribute("spaceBetween")?.Value, expectedSpaceBetween, strictMode))
                {
                    if (strictMode)
                    {
                        return false;
                    }
                    continue;
                }

                matched++;
            }

            if (strictMode)
            {
                return matched == lineOes.Count;
            }

            // Fuzzy: require most lines match the configured code font/line spacing.
            int threshold = System.Math.Max(1, (int)System.Math.Ceiling(lineOes.Count * 0.7));
            return matched >= threshold;
        }

        private static bool SpaceBetweenMatches(string value, double expected, bool strictMode)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return !strictMode;
            }

            if (!double.TryParse(value, out var actual))
            {
                return !strictMode;
            }

            return System.Math.Abs(actual - expected) < 0.3;
        }
    }
}
