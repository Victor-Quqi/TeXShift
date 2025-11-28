using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TeXShift.Core.Markdown.Abstractions;

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
                int columnIndex = 0;

                foreach (var cellBlock in row)
                {
                    if (!(cellBlock is TableCell cell)) continue;

                    var cellElement = new XElement(ns + "Cell");
                    var oeChildren = new XElement(ns + "OEChildren");

                    // Get column alignment
                    var alignment = GetColumnAlignment(table, columnIndex);

                    // Check if cell contains only a single image
                    var singleImage = GetSingleImageFromCell(cell);
                    if (singleImage != null)
                    {
                        // Use shared helper to create image OE
                        var imageOe = ImageElementHelper.CreateImageOE(singleImage, ns);
                        ApplyAlignment(imageOe, alignment);
                        oeChildren.Add(imageOe);
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

                        // Apply bold for header row
                        if (row.IsHeader && !string.IsNullOrEmpty(cellContent))
                        {
                            cellContent = $"<span style='font-weight:bold'>{cellContent}</span>";
                        }

                        var oe = new XElement(ns + "OE",
                            new XElement(ns + "T", new XCData(cellContent)));
                        ApplyAlignment(oe, alignment);
                        oeChildren.Add(oe);
                    }

                    cellElement.Add(oeChildren);
                    rowElement.Add(cellElement);
                    columnIndex++;
                }

                tableElement.Add(rowElement);
            }

            // Wrap table in OE element
            var outerOe = new XElement(ns + "OE", tableElement);
            return new[] { outerOe };
        }

        /// <summary>
        /// Gets the alignment for a specific column from the table's column definitions.
        /// </summary>
        private string GetColumnAlignment(Table table, int columnIndex)
        {
            if (table.ColumnDefinitions == null || columnIndex >= table.ColumnDefinitions.Count)
            {
                return "left";
            }

            var colDef = table.ColumnDefinitions[columnIndex];
            switch (colDef.Alignment)
            {
                case TableColumnAlign.Center:
                    return "center";
                case TableColumnAlign.Right:
                    return "right";
                default:
                    return "left";
            }
        }

        /// <summary>
        /// Applies alignment attributes to an OE element.
        /// </summary>
        private void ApplyAlignment(XElement oe, string alignment)
        {
            oe.SetAttributeValue("alignment", alignment);

            if (alignment == "center" || alignment == "right")
            {
                var existingStyle = oe.Attribute("style")?.Value ?? "";
                var textAlign = $"text-align:{alignment}";
                oe.SetAttributeValue("style", string.IsNullOrEmpty(existingStyle)
                    ? textAlign
                    : $"{existingStyle};{textAlign}");
            }
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
