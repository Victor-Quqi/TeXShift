using Markdig.Syntax;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using TeXShift.Core.Markdown.Abstractions;

namespace TeXShift.Core.Markdown.Handlers
{
    internal class QuoteBlockHandler : IBlockHandler
    {
        public async Task<IReadOnlyList<XElement>> HandleAsync(Block block, IMarkdownConverterContext context)
        {
            var quoteBlock = (QuoteBlock)block;
            var ns = context.OneNoteNamespace;
            var quoteConfig = context.StyleConfig.GetQuoteBlockStyle();
            var widthReservation = context.StyleConfig.WidthReservation;

            context.IncrementQuoteDepth();
            try
            {
                var depth = context.QuoteNestingDepth;

                var oe = new XElement(ns + "OE");

                var table = new XElement(ns + "Table");
                table.Add(new XAttribute("bordersVisible", depth == 1 ? "false" : "true"));
                table.Add(new XAttribute("hasHeaderRow", "false"));

                var columns = new XElement(ns + "Columns");

                // Calculate table width using conservative reservation strategy
                var totalReservation = widthReservation.QuoteBlockTotalReservation;
                var tableWidth = context.CurrentAvailableWidth - totalReservation;

                var column = new XElement(ns + "Column");
                column.Add(new XAttribute("index", "0"));
                column.Add(new XAttribute("width", tableWidth.ToString("F2")));
                column.Add(new XAttribute("isLocked", "true"));
                columns.Add(column);
                table.Add(columns);

                var row = new XElement(ns + "Row");
                var cell = new XElement(ns + "Cell");
                cell.Add(new XAttribute("shadingColor", quoteConfig.BackgroundColor));

                // Push width reservation before processing child blocks
                context.PushWidthReservation(totalReservation);
                IReadOnlyList<XElement> childElements;
                try
                {
                    childElements = await context.ProcessBlocksAsync(quoteBlock).ConfigureAwait(false);
                }
                finally
                {
                    context.PopWidthReservation();
                }

                var oeChildren = new XElement(ns + "OEChildren");
                if (childElements.Count > 0)
                {
                    oeChildren.Add(childElements);
                }
                else
                {
                    oeChildren.Add(new XElement(ns + "OE",
                        new XElement(ns + "T", new XCData(""))));
                }
                cell.Add(oeChildren);

                row.Add(cell);
                table.Add(row);
                oe.Add(table);

                return new[] { oe };
            }
            finally
            {
                context.DecrementQuoteDepth();
            }
        }
    }
}
