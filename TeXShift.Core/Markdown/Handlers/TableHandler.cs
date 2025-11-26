using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TeXShift.Core.Markdown;

namespace TeXShift.Core.Markdown.Handlers
{
    internal class TableHandler : IBlockHandler
    {
        private const double DefaultColumnWidth = 72.0; // Default width in points (1 inch)

        public IEnumerable<XElement> Handle(Block block, IMarkdownConverterContext context)
        {
            var table = (Table)block;
            var ns = context.OneNoteNamespace;

            // Determine column count from the first row
            var columnCount = table.FirstOrDefault() is TableRow firstRow ? firstRow.Count : 0;
            if (columnCount == 0)
            {
                return Enumerable.Empty<XElement>();
            }

            // Create Table element
            var tableElement = new XElement(ns + "Table",
                new XAttribute("bordersVisible", "true"),
                new XAttribute("hasHeaderRow", "true"));

            // Create Columns element with equal width columns
            var columnsElement = new XElement(ns + "Columns");
            for (int i = 0; i < columnCount; i++)
            {
                columnsElement.Add(new XElement(ns + "Column",
                    new XAttribute("index", i.ToString()),
                    new XAttribute("width", DefaultColumnWidth.ToString("F1"))));
            }
            tableElement.Add(columnsElement);

            // Process each row
            foreach (var rowBlock in table)
            {
                if (!(rowBlock is TableRow row)) continue;

                var rowElement = new XElement(ns + "Row");

                foreach (var cellBlock in row)
                {
                    if (!(cellBlock is TableCell cell)) continue;

                    var cellElement = new XElement(ns + "Cell");
                    var oeChildren = new XElement(ns + "OEChildren");

                    // Check if cell contains only a single image
                    var singleImage = GetSingleImageFromCell(cell);
                    if (singleImage != null)
                    {
                        // Use shared helper to create image OE
                        oeChildren.Add(ImageElementHelper.CreateImageOE(singleImage, ns));
                    }
                    else
                    {
                        // Process cell content - cells contain paragraphs
                        var cellContent = "";
                        foreach (var cellChild in cell)
                        {
                            if (cellChild is ParagraphBlock paragraph)
                            {
                                cellContent += context.ConvertInlinesToHtml(paragraph.Inline);
                            }
                        }

                        var oe = new XElement(ns + "OE",
                            new XElement(ns + "T", new XCData(cellContent)));
                        oeChildren.Add(oe);
                    }

                    cellElement.Add(oeChildren);
                    rowElement.Add(cellElement);
                }

                tableElement.Add(rowElement);
            }

            // Wrap table in OE element
            var outerOe = new XElement(ns + "OE", tableElement);
            return new[] { outerOe };
        }

        /// <summary>
        /// Checks if cell contains only a single image and returns it.
        /// Table cells have their own structure (may contain multiple paragraphs),
        /// so this requires cell-specific logic.
        /// </summary>
        private LinkInline GetSingleImageFromCell(TableCell cell)
        {
            // Cell should have exactly one paragraph
            var paragraphs = cell.OfType<ParagraphBlock>().ToList();
            if (paragraphs.Count != 1) return null;

            // Use shared helper for the paragraph-level check
            return ImageElementHelper.GetSingleImage(paragraphs[0]);
        }
    }
}
