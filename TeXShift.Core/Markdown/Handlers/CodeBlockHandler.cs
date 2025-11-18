using Markdig.Syntax;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TeXShift.Core.Markdown;
using TeXShift.Core.Utils; // Assuming an EscapeHtml utility class

namespace TeXShift.Core.Markdown.Handlers
{
    internal class CodeBlockHandler : IBlockHandler
    {
        public IEnumerable<XElement> Handle(Block block, IMarkdownConverterContext context)
        {
            var code = (CodeBlock)block;
            var ns = context.OneNoteNamespace;
            var styleConfig = context.StyleConfig;

            var oe = new XElement(ns + "OE");

            // Apply code block spacing
            var spacing = styleConfig.GetCodeSpacing();
            oe.Add(new XAttribute("spaceBefore", spacing.SpaceBefore.ToString("F1")));
            oe.Add(new XAttribute("spaceAfter", spacing.SpaceAfter.ToString("F1")));
            oe.Add(new XAttribute("spaceBetween", spacing.SpaceBetween.ToString("F1")));

            // TODO: Implement syntax highlighting here in the future.
            // For now, just wrap in a monospace font span.
            // Reconstruct the code content by joining lines with a newline, preserving indentation.
            // By accessing the .Lines property of the StringLineGroup, we provide a concrete StringLine[] array,
            // which resolves the compiler's type inference ambiguity for the .Select() extension method.
            var codeContent = string.Join("\n", code.Lines.Lines.Select(l => l.ToString()));
            var escapedCode = HtmlEscaper.Escape(codeContent);
            var htmlContent = $"<span style='font-family:Consolas'>{escapedCode}</span>";

            var textElement = new XElement(ns + "T", new XCData(htmlContent));
            oe.Add(textElement);

            return new[] { oe };
        }
    }
}