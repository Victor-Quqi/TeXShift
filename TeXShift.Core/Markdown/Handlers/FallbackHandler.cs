using Markdig.Syntax;
using System.Collections.Generic;
using System.Xml.Linq;
using TeXShift.Core.Markdown;
using TeXShift.Core.Utils;

namespace TeXShift.Core.Markdown.Handlers
{
    /// <summary>
    /// Handles any Markdig Block that doesn't have a specific handler.
    /// This acts as a safety net, ensuring that unsupported Markdown content
    /// is still rendered as plain text rather than being lost.
    /// </summary>
    internal class FallbackHandler : IBlockHandler
    {
        public IEnumerable<XElement> Handle(Block block, IMarkdownConverterContext context)
        {
            var ns = context.OneNoteNamespace;
            var oe = new XElement(ns + "OE");

            // Simply render the block's content as escaped plain text.
            var rawText = block.ToString(); // This might need adjustment based on block type
            var escapedText = HtmlEscaper.Escape(rawText);

            oe.Add(new XElement(ns + "T", new XCData(escapedText)));

            return new[] { oe };
        }
    }
}