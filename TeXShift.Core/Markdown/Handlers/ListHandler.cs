using Markdig.Syntax;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TeXShift.Core.Markdown;

namespace TeXShift.Core.Markdown.Handlers
{
    internal class ListHandler : IBlockHandler
    {
        public IEnumerable<XElement> Handle(Block block, IMarkdownConverterContext context)
        {
            var listBlock = (ListBlock)block;
            var elements = new List<XElement>();

            // The main Handle method now only iterates through the top-level items.
            // The recursive logic is delegated to the ProcessListItemBlock helper method.
            foreach (var listItem in listBlock.OfType<ListItemBlock>())
            {
                elements.Add(ProcessListItemBlock(listItem, listBlock.IsOrdered, context));
            }

            return elements;
        }

        /// <summary>
        /// Recursively processes a single list item and all its potential nested lists.
        /// This is the core of the nesting logic.
        /// </summary>
        private XElement ProcessListItemBlock(ListItemBlock listItem, bool isOrdered, IMarkdownConverterContext context)
        {
            var ns = context.OneNoteNamespace;
            var styleConfig = context.StyleConfig;

            var oe = new XElement(ns + "OE");

            // Apply list item spacing from style configuration.
            var spacing = styleConfig.GetListSpacing();
            oe.Add(new XAttribute("spaceBefore", spacing.SpaceBefore.ToString("F1")));
            oe.Add(new XAttribute("spaceAfter", spacing.SpaceAfter.ToString("F1")));
            oe.Add(new XAttribute("spaceBetween", spacing.SpaceBetween.ToString("F1")));

            // Add the List element for the bullet or number.
            var listElement = new XElement(ns + "List");
            if (isOrdered)
            {
                var number = new XElement(ns + "Number",
                    new XAttribute("numberSequence", "0"),
                    new XAttribute("numberFormat", "##."),
                    new XAttribute("fontSize", "11.0"));
                listElement.Add(number);
            }
            else
            {
                var bullet = new XElement(ns + "Bullet",
                    new XAttribute("bullet", "2"),
                    new XAttribute("fontSize", "11.0"));
                listElement.Add(bullet);
            }
            oe.Add(listElement);

            // Add the main text content when the first block is a paragraph.
            // Other blocks (quote, code, nested lists, extra paragraphs) are processed below.
            var firstBlock = listItem.FirstOrDefault();
            if (firstBlock is ParagraphBlock paragraph)
            {
                var htmlContent = context.ConvertInlinesToHtml(paragraph.Inline);
                oe.Add(new XElement(ns + "T", new XCData(htmlContent)));
            }
            else
            {
                oe.Add(new XElement(ns + "T", new XCData(string.Empty)));
            }

            // Process any remaining child blocks inside this list item to preserve nested structures.
            var remainingBlocks = listItem.Skip(firstBlock is ParagraphBlock ? 1 : 0).ToList();
            if (remainingBlocks.Any())
            {
                var childrenContainer = new XElement(ns + "OEChildren");

                // Push width reservation for list item before processing nested blocks
                var widthReservation = styleConfig.WidthReservation;
                var reservation = widthReservation.GetListItemReservation(isOrdered);
                context.PushWidthReservation(reservation);
                var convertedChildren = context.ProcessBlocks(remainingBlocks).ToList();
                context.PopWidthReservation();

                if (convertedChildren.Any())
                {
                    childrenContainer.Add(convertedChildren);
                    oe.Add(childrenContainer);
                }
            }

            return oe;
        }
    }
}
