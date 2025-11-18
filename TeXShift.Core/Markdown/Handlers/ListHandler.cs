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
            var list = (ListBlock)block;
            var ns = context.OneNoteNamespace;
            var styleConfig = context.StyleConfig;
            var result = new List<XElement>();

            foreach (var item in list.OfType<ListItemBlock>())
            {
                // Each ListItemBlock can contain nested blocks (e.g., a paragraph).
                // We process these nested blocks recursively.
                foreach (var childBlock in item)
                {
                    if (childBlock is ParagraphBlock para)
                    {
                        var oe = new XElement(ns + "OE");

                        // Apply list item spacing
                        var spacing = styleConfig.GetListSpacing();
                        oe.Add(new XAttribute("spaceBefore", spacing.SpaceBefore.ToString("F1")));
                        oe.Add(new XAttribute("spaceAfter", spacing.SpaceAfter.ToString("F1")));
                        oe.Add(new XAttribute("spaceBetween", spacing.SpaceBetween.ToString("F1")));

                        // Add List element for bullet/number
                        var listElement = new XElement(ns + "List");
                        if (list.IsOrdered)
                        {
                            var number = new XElement(ns + "Number",
                                new XAttribute("numberSequence", "0"),
                                new XAttribute("numberFormat", "##."),
                                new XAttribute("fontSize", "11.0")
                            );
                            listElement.Add(number);
                        }
                        else
                        {
                            var bullet = new XElement(ns + "Bullet",
                                new XAttribute("bullet", "2"),
                                new XAttribute("fontSize", "11.0")
                            );
                            listElement.Add(bullet);
                        }
                        oe.Add(listElement);

                        // Add text content
                        var htmlContent = context.ConvertInlinesToHtml(para.Inline);
                        oe.Add(new XElement(ns + "T", new XCData(htmlContent)));

                        result.Add(oe);
                    }
                    else
                    {
                        // Handle other potential nested blocks within a list item if necessary
                        result.AddRange(context.ProcessBlocks(new[] { childBlock }));
                    }
                }
            }

            return result;
        }
    }
}