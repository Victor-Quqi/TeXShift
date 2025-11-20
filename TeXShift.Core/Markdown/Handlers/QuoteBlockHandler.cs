using Markdig.Syntax;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace TeXShift.Core.Markdown.Handlers
{
    internal class QuoteBlockHandler : IBlockHandler
    {
        public IEnumerable<XElement> Handle(Block block, IMarkdownConverterContext context)
        {
            var quoteBlock = (QuoteBlock)block;
            var ns = context.OneNoteNamespace;
            var quoteConfig = context.StyleConfig.GetQuoteBlockStyle();

            context.IncrementQuoteDepth();
            var depth = context.QuoteNestingDepth;

            var oe = new XElement(ns + "OE");

            var table = new XElement(ns + "Table");
            table.Add(new XAttribute("bordersVisible", depth == 1 ? "false" : "true"));
            table.Add(new XAttribute("hasHeaderRow", "false"));

            var columns = new XElement(ns + "Columns");
            var width = context.SourceOutlineWidth.HasValue
                ? context.SourceOutlineWidth.Value - (depth * quoteConfig.WidthReduction)
                : quoteConfig.BaseWidth - (depth - 1) * quoteConfig.WidthReduction;
            var column = new XElement(ns + "Column");
            column.Add(new XAttribute("index", "0"));
            column.Add(new XAttribute("width", width.ToString("F2")));
            column.Add(new XAttribute("isLocked", "true"));
            columns.Add(column);
            table.Add(columns);

            var row = new XElement(ns + "Row");
            var cell = new XElement(ns + "Cell");
            cell.Add(new XAttribute("shadingColor", quoteConfig.BackgroundColor));

            var childElements = context.ProcessBlocks(quoteBlock).ToList();
            var oeChildren = new XElement(ns + "OEChildren");
            if (childElements.Any())
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

            context.DecrementQuoteDepth();

            return new[] { oe };
        }
    }
}
