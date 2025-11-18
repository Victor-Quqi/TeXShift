using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using TeXShift.Core.Markdown;
using TeXShift.Core.Markdown.Handlers;
using TeXShift.Core.Utils;

namespace TeXShift.Core
{
    /// <summary>
    /// Converts Markdown text to OneNote XML format by dispatching to specialized block handlers.
    /// This class acts as a coordinator, parsing the Markdown and delegating the conversion
    /// of each block type to a registered handler.
    /// </summary>
    public class MarkdownConverter : IMarkdownConverter, IMarkdownConverterContext
    {
        private readonly Dictionary<Type, IBlockHandler> _blockHandlers;
        private readonly FallbackHandler _fallbackHandler = new FallbackHandler();
        private readonly MarkdownPipeline _pipeline;

        private static readonly Regex SpanLangRegex = new Regex(@"<span\s+lang=[^>]+>(.*?)</span>", RegexOptions.Compiled | RegexOptions.Singleline);

        // Explicit implementation of IMarkdownConverterContext properties
        public XNamespace OneNoteNamespace { get; } = "http://schemas.microsoft.com/office/onenote/2013/onenote";
        public OneNoteStyleConfig StyleConfig { get; }

        public MarkdownConverter(OneNoteStyleConfig styleConfig, MarkdownPipeline pipeline)
        {
            StyleConfig = styleConfig ?? throw new ArgumentNullException(nameof(styleConfig));
            _pipeline = pipeline ?? throw new ArgumentNullException(nameof(pipeline));

            // Register all the specialized handlers for each block type.
            _blockHandlers = new Dictionary<Type, IBlockHandler>
            {
                { typeof(HeadingBlock), new HeadingHandler() },
                { typeof(ParagraphBlock), new ParagraphHandler() },
                { typeof(ListBlock), new ListHandler() },
                { typeof(CodeBlock), new CodeBlockHandler() },
                // Add new handlers here, e.g., { typeof(TableBlock), new TableHandler() }
            };
        }

        public async Task<XElement> ConvertToOneNoteXmlAsync(string markdown)
        {
            if (string.IsNullOrWhiteSpace(markdown))
            {
                return CreateEmptyOutline();
            }
            return await Task.Run(() => ConvertToOneNoteXml(markdown)).ConfigureAwait(false);
        }

        private XElement ConvertToOneNoteXml(string markdown)
        {
            var sanitizedMarkdown = SanitizeText(markdown);
            var document = Markdig.Markdown.Parse(sanitizedMarkdown, _pipeline);
            var outline = new XElement(OneNoteNamespace + "Outline");

            // Add the Indents element to control layout and prevent default margins.
            var indentsElement = new XElement(OneNoteNamespace + "Indents");
            foreach (var indent in StyleConfig.Indents)
            {
                indentsElement.Add(new XElement(OneNoteNamespace + "Indent",
                    new XAttribute("level", indent.Key.ToString()),
                    new XAttribute("indent", indent.Value.ToString("F1"))));
            }
            outline.Add(indentsElement);

            var oeChildren = new XElement(OneNoteNamespace + "OEChildren");

            // Final, correct logic:
            // 1. Process all blocks into a flat list of XElements.
            // 2. Post-process to handle a specific OneNote rendering quirk:
            //    If a ListBlock follows a HeadingBlock, it MUST be nested to avoid default indentation on the Heading.
            //    Other blocks (like Paragraphs) should remain at the top level.

            var blocks = document.ToList();
            var elements = new List<XElement>();
            XElement lastHeadingElement = null;

            for (int i = 0; i < blocks.Count; i++)
            {
                var block = blocks[i];
                if (block is LinkReferenceDefinitionGroup) continue;

                var processed = HandleBlock(block).ToList();

                // Check if the current block is a list.
                if (block is ListBlock && lastHeadingElement != null)
                {
                    // This is a list that follows a heading. Nest all its generated elements.
                    var childrenContainer = lastHeadingElement.Element(OneNoteNamespace + "OEChildren");
                    if (childrenContainer == null)
                    {
                        childrenContainer = new XElement(OneNoteNamespace + "OEChildren");
                        lastHeadingElement.Add(childrenContainer);
                    }
                    childrenContainer.Add(processed);
                }
                else
                {
                    elements.AddRange(processed);
                    // If this block is a heading, track it. Otherwise, reset.
                    lastHeadingElement = (block is HeadingBlock) ? processed.LastOrDefault() : null;
                }
            }
            
            oeChildren.Add(elements);
            outline.Add(oeChildren);
            return outline;
        }

        public IEnumerable<XElement> ProcessBlocks(IEnumerable<Block> blocks)
        {
            var elements = new List<XElement>();
            foreach (var block in blocks)
            {
                elements.AddRange(HandleBlock(block));
            }
            return elements;
        }

        private IEnumerable<XElement> HandleBlock(Block block)
        {
            if (block is LinkReferenceDefinitionGroup) return Enumerable.Empty<XElement>();

            IBlockHandler handler;
            if (!_blockHandlers.TryGetValue(block.GetType(), out handler))
            {
                handler = _fallbackHandler;
            }
            return handler.Handle(block, this);
        }


        private string SanitizeText(string text)
        {
            // Recursively remove lang spans to handle nested cases.
            while (SpanLangRegex.IsMatch(text))
            {
                text = SpanLangRegex.Replace(text, "$1");
            }
            return text;
        }

        public string ConvertInlinesToHtml(ContainerInline container)
        {
            if (container == null) return string.Empty;
            var html = new StringBuilder();
            foreach (var inline in container)
            {
                if (inline is LiteralInline literal)
                {
                    html.Append(HtmlEscaper.Escape(literal.Content.ToString()));
                }
                else if (inline is EmphasisInline emphasis)
                {
                    var content = ConvertInlinesToHtml(emphasis);
                    if (emphasis.DelimiterChar == '*' || emphasis.DelimiterChar == '_')
                    {
                        if (emphasis.DelimiterCount == 2) html.Append($"<span style='font-weight:bold'>{content}</span>");
                        else if (emphasis.DelimiterCount == 1) html.Append($"<span style='font-style:italic'>{content}</span>");
                        else html.Append(content);
                    }
                    else if (emphasis.DelimiterChar == '~' && emphasis.DelimiterCount == 2)
                    {
                        html.Append($"<span style='text-decoration:line-through'>{content}</span>");
                    }
                    else html.Append(content);
                }
                else if (inline is CodeInline code)
                {
                    var style = StyleConfig.GetInlineCodeStyle();
                    // OneNote does not support 'padding' on <span> elements.
                    // We simulate padding by repeating a configured character (e.g., &nbsp;) inside the span.
                    var styleString = $"font-family:{style.FontFamily};background-color:{style.BackgroundColor}";
                    var padding = new StringBuilder();
                    for (int i = 0; i < style.PaddingCount; i++)
                    {
                        padding.Append(style.PaddingChar);
                    }
                    html.Append($"<span style='{styleString}'>{padding}{HtmlEscaper.Escape(code.Content)}{padding}</span>");
                }
                else if (inline is LineBreakInline)
                {
                    html.Append("\n");
                }
                else if (inline is ContainerInline nested)
                {
                    html.Append(ConvertInlinesToHtml(nested));
                }
            }
            return html.ToString();
        }

        private XElement CreateEmptyOutline()
        {
            return new XElement(OneNoteNamespace + "Outline",
                new XElement(OneNoteNamespace + "OEChildren",
                    new XElement(OneNoteNamespace + "OE",
                        new XElement(OneNoteNamespace + "T", new XCData(""))
                    )
                )
            );
        }
    }
}
