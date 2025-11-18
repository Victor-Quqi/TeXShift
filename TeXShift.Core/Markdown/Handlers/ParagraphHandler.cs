using Markdig.Syntax;
using System.Collections.Generic;
using System.Xml.Linq;
using TeXShift.Core.Markdown;

namespace TeXShift.Core.Markdown.Handlers
{
    internal class ParagraphHandler : IBlockHandler
    {
        public IEnumerable<XElement> Handle(Block block, IMarkdownConverterContext context)
        {
            var paragraph = (ParagraphBlock)block;
            var ns = context.OneNoteNamespace;
            var styleConfig = context.StyleConfig;

            var oe = new XElement(ns + "OE");

            // Apply paragraph spacing
            var spacing = styleConfig.GetParagraphSpacing();
            oe.Add(new XAttribute("spaceBefore", spacing.SpaceBefore.ToString("F1")));
            oe.Add(new XAttribute("spaceAfter", spacing.SpaceAfter.ToString("F1")));
            oe.Add(new XAttribute("spaceBetween", spacing.SpaceBetween.ToString("F1")));

            // Convert inline content to HTML
            var htmlContent = context.ConvertInlinesToHtml(paragraph.Inline);
            oe.Add(new XElement(ns + "T", new XCData(htmlContent)));

            return new[] { oe };
        }
    }
}