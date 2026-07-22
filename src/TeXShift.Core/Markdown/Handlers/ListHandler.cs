using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Markdig.Extensions.TaskLists;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using TeXShift.Core.Markdown.Abstractions;

namespace TeXShift.Core.Markdown.Handlers
{
    internal class ListHandler : IBlockHandler
    {
        public async Task<IReadOnlyList<XElement>> HandleAsync(Block block, IMarkdownConverterContext context)
        {
            var listBlock = (ListBlock)block;
            var elements = new List<XElement>();
            var orderedStart = 1;
            if (listBlock.IsOrdered && !int.TryParse(listBlock.OrderedStart, out orderedStart))
            {
                orderedStart = 1;
            }

            // ProcessListItemBlock returns multiple elements: the list item OE followed by
            // any sibling content (code blocks, blockquotes) that should appear after it.
            // Nested lists remain as OEChildren for proper indentation.
            var isFirstItem = true;
            foreach (var listItem in listBlock.OfType<ListItemBlock>())
            {
                var itemStart = isFirstItem ? orderedStart : 1;
                var itemElements = await ProcessListItemBlockAsync(listItem, listBlock.IsOrdered, itemStart, context).ConfigureAwait(false);
                elements.AddRange(itemElements);
                isFirstItem = false;
            }

            return elements;
        }

        /// <summary>
         /// Processes a single list item and returns it along with any sibling content blocks.
         /// - Nested ListBlocks remain as OEChildren (for proper indentation)
         /// - Other blocks (CodeBlock, QuoteBlock, etc.) become siblings for full-width display
         /// </summary>
        private async Task<IReadOnlyList<XElement>> ProcessListItemBlockAsync(ListItemBlock listItem, bool isOrdered, int orderedStart, IMarkdownConverterContext context)
        {
            var ns = context.OneNoteNamespace;
            var styleConfig = context.StyleConfig;
            var firstBlock = listItem.FirstOrDefault();

            // Check if this is a task list item
            var taskList = GetTaskListFromBlock(firstBlock);

            // Create the main OE element with spacing and list marker
            var oe = CreateListItemOE(ns, styleConfig, isOrdered, orderedStart, taskList);

            // Add the main text/image content
            await AddMainContentAsync(oe, firstBlock, taskList, ns, context).ConfigureAwait(false);

            // Process remaining blocks (nested lists and siblings)
            var results = new List<XElement> { oe };
            await ProcessRemainingBlocksAsync(listItem, firstBlock, oe, results, isOrdered, ns, styleConfig, context).ConfigureAwait(false);

            return results;
        }

        /// <summary>
        /// Gets TaskList inline from the first block if it's a paragraph with a task list checkbox.
        /// </summary>
        private TaskList GetTaskListFromBlock(Block firstBlock)
        {
            if (firstBlock is ParagraphBlock paragraph)
            {
                return paragraph.Inline?.Descendants<TaskList>().FirstOrDefault();
            }
            return null;
        }

        /// <summary>
        /// Creates the base OE element with spacing and list marker (Tag or List element).
        /// </summary>
        private XElement CreateListItemOE(XNamespace ns, Configuration.OneNoteStyleConfig styleConfig, bool isOrdered, int orderedStart, TaskList taskList)
        {
            var oe = new XElement(ns + "OE");

            // Apply list item spacing
            var spacing = styleConfig.GetListSpacing();
            oe.Add(new XAttribute("spaceBefore", spacing.SpaceBefore.ToString("F1")));
            oe.Add(new XAttribute("spaceAfter", spacing.SpaceAfter.ToString("F1")));
            oe.Add(new XAttribute("spaceBetween", spacing.SpaceBetween.ToString("F1")));

            // Add list marker
            if (taskList != null)
            {
                oe.Add(CreateTaskListTag(ns, taskList));
            }
            else
            {
                oe.Add(CreateListElement(ns, isOrdered, orderedStart));
            }

            return oe;
        }

        /// <summary>
        /// Creates a Tag element for task list checkbox.
        /// </summary>
        private XElement CreateTaskListTag(XNamespace ns, TaskList taskList)
        {
            var tag = new XElement(ns + "Tag",
                new XAttribute("index", "0"),
                new XAttribute("completed", taskList.Checked.ToString().ToLower()),
                new XAttribute("disabled", "false"),
                new XAttribute("creationDate", DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")));

            if (taskList.Checked)
            {
                tag.Add(new XAttribute("completionDate", DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")));
            }

            return tag;
        }

        /// <summary>
        /// Creates a List element for bullet or numbered list.
        /// </summary>
        private XElement CreateListElement(XNamespace ns, bool isOrdered, int orderedStart)
        {
            var listElement = new XElement(ns + "List");

            if (isOrdered)
            {
                var numberElement = new XElement(ns + "Number",
                    new XAttribute("numberSequence", "0"),
                    new XAttribute("numberFormat", "##."),
                    new XAttribute("fontSize", "11.0"));

                if (orderedStart != 1)
                {
                    numberElement.Add(new XAttribute("restartNumberingAt", orderedStart));
                }

                listElement.Add(numberElement);
            }
            else
            {
                listElement.Add(new XElement(ns + "Bullet",
                    new XAttribute("bullet", "2"),
                    new XAttribute("fontSize", "11.0")));
            }

            return listElement;
        }

         /// <summary>
         /// Adds the main text or image content to the list item OE.
         /// </summary>
        private async Task AddMainContentAsync(XElement oe, Block firstBlock, TaskList taskList, XNamespace ns, IMarkdownConverterContext context)
        {
            if (!(firstBlock is ParagraphBlock paragraphBlock))
            {
                oe.Add(new XElement(ns + "T", new XCData(string.Empty)));
                return;
            }

            // Check if the paragraph contains only a single image
            var singleImage = ImageElementHelper.GetSingleImage(paragraphBlock, filterTaskList: true);
            if (singleImage != null)
            {
                var imageElement = await ImageElementHelper.CreateImageElementAsync(singleImage, ns, context).ConfigureAwait(false);
                oe.Add(imageElement ?? ImageElementHelper.CreateImageFallback(singleImage, ns));
                return;
            }

            // Convert inline content to HTML
            var htmlContent = await context.ConvertInlinesToHtmlAsync(paragraphBlock.Inline).ConfigureAwait(false);

            // Trim leading whitespace for task list items
            if (taskList != null)
            {
                htmlContent = htmlContent.TrimStart();
            }

            oe.Add(new XElement(ns + "T", new XCData(htmlContent)));
        }

         /// <summary>
         /// Processes remaining child blocks: nested lists become OEChildren, others become siblings.
         /// </summary>
        private async Task ProcessRemainingBlocksAsync(ListItemBlock listItem, Block firstBlock, XElement oe, List<XElement> results,
            bool isOrdered, XNamespace ns, Configuration.OneNoteStyleConfig styleConfig, IMarkdownConverterContext context)
        {
            var remainingBlocks = listItem.Skip(firstBlock is ParagraphBlock ? 1 : 0).ToList();
            if (!remainingBlocks.Any()) return;

            var nestedListBlocks = remainingBlocks.Where(b => b is ListBlock).ToList();
            var siblingBlocks = remainingBlocks.Where(b => !(b is ListBlock)).ToList();

            // Process nested lists as OEChildren (for proper indentation)
            if (nestedListBlocks.Any())
            {
                var childrenContainer = new XElement(ns + "OEChildren");

                var reservation = styleConfig.WidthReservation.GetListItemReservation(isOrdered);
                context.PushWidthReservation(reservation);
                IReadOnlyList<XElement> convertedChildren;
                try
                {
                    convertedChildren = await context.ProcessBlocksAsync(nestedListBlocks).ConfigureAwait(false);
                }
                finally
                {
                    context.PopWidthReservation();
                }

                if (convertedChildren.Count > 0)
                {
                    childrenContainer.Add(convertedChildren);
                    oe.Add(childrenContainer);
                }
            }

            // Process other blocks as siblings
            if (siblingBlocks.Any())
            {
                var convertedSiblings = await context.ProcessBlocksAsync(siblingBlocks).ConfigureAwait(false);
                results.AddRange(convertedSiblings);
            }
        }
    }
}
