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

            // A ListItemBlock in Markdig contains a ParagraphBlock for its text content.
            var paragraph = listItem.OfType<ParagraphBlock>().FirstOrDefault();
            if (paragraph == null)
            {
                // This case is unlikely for standard lists but provides robustness.
                return new XElement(ns + "OE", new XElement(ns + "T", new XCData("")));
            }

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

            // Add the actual text content of the list item.
            var htmlContent = context.ConvertInlinesToHtml(paragraph.Inline);
            oe.Add(new XElement(ns + "T", new XCData(htmlContent)));

            // --- RECURSION LOGIC ---
            // Check if this list item contains a nested list.
            var nestedList = listItem.OfType<ListBlock>().FirstOrDefault();
            if (nestedList != null)
            {
                // If a nested list exists, create an OEChildren container for it.
                var childrenContainer = new XElement(ns + "OEChildren");

                // Recursively call this same method for each item in the nested list.
                foreach (var nestedItem in nestedList.OfType<ListItemBlock>())
                {
                    childrenContainer.Add(ProcessListItemBlock(nestedItem, nestedList.IsOrdered, context));
                }

                // Attach the container with the nested items to the current OE element.
                oe.Add(childrenContainer);
            }

            return oe;
        }
    }
}