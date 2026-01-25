using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using TeXShift.Core.OneNoteToMarkdown.Abstractions;

namespace TeXShift.Core.OneNoteToMarkdown.Handlers
{
    /// <summary>
    /// Converts OneNote quote block tables to Markdown blockquotes.
    /// </summary>
    internal sealed class QuoteBlockElementHandler : IElementHandler
    {
        public bool CanHandle(XElement element, IOneNoteConverterContext context)
        {
            var table = OneNoteTableHelpers.GetTable(element, context);
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

            // Quotes use a shaded single-cell table. Do not compare the color value.
            if (cell.Attribute("shadingColor") == null)
            {
                return false;
            }

            bool strictMode = !context.StyleConfig.TryRecognizeNonTeXShiftFormats;
            if (strictMode && !OneNoteTableHelpers.IsLockedSingleColumn(table, ns))
            {
                return false;
            }

            // Avoid mis-detecting code blocks as quotes.
            if (new CodeBlockElementHandler().CanHandle(element, context))
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

            var cell = table.Descendants(ns + "Cell").FirstOrDefault();
            if (cell == null)
            {
                yield break;
            }

            var oeChildren = cell.Element(ns + "OEChildren");
            var innerMarkdown = context.ConvertOeChildrenToMarkdown(oeChildren);

            var lines = SplitLines(innerMarkdown);
            if (lines.Length == 0)
            {
                yield break;
            }

            var sb = new StringBuilder();
            for (int i = 0; i < lines.Length; i++)
            {
                if (i > 0)
                {
                    sb.AppendLine();
                }

                var line = lines[i] ?? string.Empty;
                if (string.IsNullOrWhiteSpace(line))
                {
                    sb.Append(">");
                }
                else
                {
                    sb.Append("> ").Append(line);
                }
            }

            yield return sb.ToString().TrimEnd();
        }

        private static string[] SplitLines(string markdown)
        {
            if (string.IsNullOrEmpty(markdown))
            {
                return new[] { string.Empty };
            }

            var normalized = markdown.Replace("\r\n", "\n").Replace("\r", "\n");
            return normalized.Split(new[] { '\n' }, StringSplitOptions.None);
        }

    }
}
