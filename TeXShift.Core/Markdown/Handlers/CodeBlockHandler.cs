using Markdig.Syntax;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TeXShift.Core.Syntax;
using TeXShift.Core.Utils;

namespace TeXShift.Core.Markdown.Handlers
{
    internal class CodeBlockHandler : IBlockHandler
    {
        public IEnumerable<XElement> Handle(Block block, IMarkdownConverterContext context)
        {
            var codeBlock = block as CodeBlock;
            var ns = context.OneNoteNamespace;
            var styleConfig = context.StyleConfig;
            var codeConfig = styleConfig.GetCodeBlockStyle();

            // Get language from fenced code block
            var language = "";
            if (codeBlock is FencedCodeBlock fenced)
            {
                language = fenced.Info ?? "";
            }

            // Create outer OE with spacing
            var outerOe = new XElement(ns + "OE");
            var spacing = styleConfig.GetCodeSpacing();
            outerOe.Add(new XAttribute("spaceBefore", spacing.SpaceBefore.ToString("F1")));
            outerOe.Add(new XAttribute("spaceAfter", spacing.SpaceAfter.ToString("F1")));

            // Create table container (single column, no borders)
            var table = new XElement(ns + "Table",
                new XAttribute("bordersVisible", "false"),
                new XAttribute("hasHeaderRow", "false"));

            // Calculate table width
            var widthReservation = styleConfig.WidthReservation;
            var tableWidth = context.CurrentAvailableWidth - widthReservation.TableSystemOverhead;

            // Create single column
            var columns = new XElement(ns + "Columns",
                new XElement(ns + "Column",
                    new XAttribute("index", "0"),
                    new XAttribute("width", tableWidth.ToString("F2")),
                    new XAttribute("isLocked", "true")));
            table.Add(columns);

            // Create row with shaded cell
            var row = new XElement(ns + "Row");
            var cell = new XElement(ns + "Cell",
                new XAttribute("shadingColor", codeConfig.BackgroundColor));

            // Create OEChildren for code lines
            var oeChildren = new XElement(ns + "OEChildren");

            // Get code content as lines
            var codeLines = GetCodeLines(codeBlock);

            // Create syntax highlighter
            var highlighter = new OneNoteCodeHighlighter(codeConfig);
            var oeStyle = codeConfig.GetOEStyle();

            // Create an OE for each line of code
            foreach (var line in codeLines)
            {
                var highlightedContent = highlighter.HighlightLine(line, language);

                // If line is empty, use a non-breaking space to maintain line height
                if (string.IsNullOrEmpty(highlightedContent))
                {
                    highlightedContent = "&nbsp;";
                }

                var lineOe = new XElement(ns + "OE",
                    new XAttribute("style", oeStyle),
                    new XElement(ns + "T", new XCData(highlightedContent)));

                oeChildren.Add(lineOe);
            }

            // If no lines, add an empty OE
            if (!codeLines.Any())
            {
                var emptyOe = new XElement(ns + "OE",
                    new XAttribute("style", oeStyle),
                    new XElement(ns + "T", new XCData("&nbsp;")));
                oeChildren.Add(emptyOe);
            }

            // Assemble table structure
            cell.Add(oeChildren);
            row.Add(cell);
            table.Add(row);
            outerOe.Add(table);

            return new[] { outerOe };
        }

        /// <summary>
        /// Extracts code lines from a CodeBlock, preserving indentation.
        /// </summary>
        private IList<string> GetCodeLines(CodeBlock codeBlock)
        {
            var lines = new List<string>();

            if (codeBlock.Lines.Lines == null)
            {
                return lines;
            }

            foreach (var line in codeBlock.Lines.Lines)
            {
                if (line.Slice.Text == null)
                {
                    continue;
                }
                lines.Add(line.ToString());
            }

            // Remove trailing empty lines
            while (lines.Count > 0 && string.IsNullOrWhiteSpace(lines[lines.Count - 1]))
            {
                lines.RemoveAt(lines.Count - 1);
            }

            return lines;
        }
    }
}
