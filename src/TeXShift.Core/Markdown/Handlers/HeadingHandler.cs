using Markdig.Syntax;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Xml.Linq;
using TeXShift.Core.Markdown.Abstractions;

namespace TeXShift.Core.Markdown.Handlers
{
    internal class HeadingHandler : IBlockHandler
    {
        public async Task<IReadOnlyList<XElement>> HandleAsync(Block block, IMarkdownConverterContext context)
        {
            var heading = (HeadingBlock)block;
            var ns = context.OneNoteNamespace;
            var styleConfig = context.StyleConfig;

            var oe = new XElement(ns + "OE");

            // Map heading levels to OneNote quick styles
            int styleIndex = System.Math.Min(heading.Level - 1, 5);
            oe.Add(new XAttribute("quickStyleIndex", styleIndex.ToString()));

            // Apply spacing based on heading level
            var spacing = styleConfig.GetHeadingSpacing(heading.Level);
            oe.Add(new XAttribute("spaceBefore", spacing.SpaceBefore.ToString("F1")));
            oe.Add(new XAttribute("spaceAfter", spacing.SpaceAfter.ToString("F1")));
            oe.Add(new XAttribute("spaceBetween", spacing.SpaceBetween.ToString("F1")));

            // Get font configuration for this heading level
            var fontConfig = styleConfig.GetHeadingFont(heading.Level);

            // Convert inline content to HTML and apply font styles
            var htmlContent = await context.ConvertInlinesToHtmlAsync(heading.Inline).ConfigureAwait(false);
            var styleAttributes = $"font-size:{fontConfig.FontSize}pt";
            if (fontConfig.IsBold)
            {
                styleAttributes += ";font-weight:bold";
            }
            var styledHeading = $"<span style='{styleAttributes}'>{htmlContent}</span>";
            oe.Add(new XElement(ns + "T", new XCData(styledHeading)));

            return new[] { oe };
        }
    }
}
