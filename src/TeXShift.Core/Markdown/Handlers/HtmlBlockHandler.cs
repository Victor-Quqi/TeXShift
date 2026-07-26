using System.Collections.Generic;
using System.Threading.Tasks;
using System.Xml.Linq;
using Markdig.Syntax;
using TeXShift.Core.Markdown.Abstractions;
using TeXShift.Core.Markdown.Processing;
using TeXShift.Core.Utils;

namespace TeXShift.Core.Markdown.Handlers
{
    internal class HtmlBlockHandler : IBlockHandler
    {
        public Task<IReadOnlyList<XElement>> HandleAsync(Block block, IMarkdownConverterContext context)
        {
            var htmlBlock = (HtmlBlock)block;
            var rawHtml = ExtractRawHtml(htmlBlock);
            var content = HtmlLineBreakParser.TryConvertLineBreakOnlyHtml(rawHtml, out var lineBreaks)
                ? lineBreaks
                : OneNoteHtmlTextEscaper.Escape(rawHtml);

            var oe = new XElement(
                context.OneNoteNamespace + "OE",
                new XElement(context.OneNoteNamespace + "T", new XCData(content)));

            return Task.FromResult<IReadOnlyList<XElement>>(new[] { oe });
        }

        private static string ExtractRawHtml(HtmlBlock htmlBlock)
        {
            if (htmlBlock.Lines.Lines == null)
            {
                return string.Empty;
            }

            var lines = new List<string>();
            foreach (var line in htmlBlock.Lines.Lines)
            {
                if (line.Slice.Text != null)
                {
                    lines.Add(line.ToString());
                }
            }

            return string.Join("\n", lines);
        }
    }
}
