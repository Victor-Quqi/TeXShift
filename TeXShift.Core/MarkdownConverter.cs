using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Markdig.Extensions.TaskLists;
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
        private int _quoteNestingDepth = 0;
        private readonly Stack<double> _widthReservationStack = new Stack<double>();
        private readonly double _initialWidth;

        private static readonly Regex SpanLangRegex = new Regex(@"<span\s+lang=[^>]+>(.*?)</span>", RegexOptions.Compiled | RegexOptions.Singleline);

        // Regex to match HTML entities (e.g., &lt;, &gt;, &amp;, &#60;, &#x3C;)
        private static readonly Regex HtmlEntityRegex = new Regex(@"&(?:lt|gt|amp|quot|apos|#\d+|#x[0-9a-fA-F]+);", RegexOptions.Compiled);

        // Explicit implementation of IMarkdownConverterContext properties
        public XNamespace OneNoteNamespace { get; } = "http://schemas.microsoft.com/office/onenote/2013/onenote";
        public OneNoteStyleConfig StyleConfig { get; }
        public int QuoteNestingDepth => _quoteNestingDepth;
        public double? SourceOutlineWidth { get; }

        /// <summary>
        /// Gets the current available width after subtracting all parent reservations.
        /// Minimum value is 50.0 points to prevent degenerate cases.
        /// </summary>
        public double CurrentAvailableWidth
        {
            get
            {
                var totalReserved = _widthReservationStack.Sum();
                var available = _initialWidth - totalReserved;
                return Math.Max(available, 50.0);
            }
        }

        public MarkdownConverter(OneNoteStyleConfig styleConfig, MarkdownPipeline pipeline, double? sourceOutlineWidth = null)
        {
            StyleConfig = styleConfig ?? throw new ArgumentNullException(nameof(styleConfig));
            _pipeline = pipeline ?? throw new ArgumentNullException(nameof(pipeline));
            SourceOutlineWidth = sourceOutlineWidth;
            _initialWidth = sourceOutlineWidth ?? StyleConfig.GetQuoteBlockStyle().BaseWidth;

            // Register all the specialized handlers for each block type.
            _blockHandlers = new Dictionary<Type, IBlockHandler>
            {
                { typeof(HeadingBlock), new HeadingHandler() },
                { typeof(ParagraphBlock), new ParagraphHandler() },
                { typeof(ListBlock), new ListHandler() },
                { typeof(CodeBlock), new CodeBlockHandler() },
                { typeof(ThematicBreakBlock), new HorizontalRuleHandler() },
                { typeof(QuoteBlock), new QuoteBlockHandler() }
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

            // Step 1: HtmlDecode to restore Markdown syntax characters
            // (e.g., &gt; → >, &lt; → <, &amp;lt; → &lt;)
            // This allows Markdig to recognize syntax like "> quote" while preserving
            // user's intended HTML entities like "&lt;" for display.
            sanitizedMarkdown = WebUtility.HtmlDecode(sanitizedMarkdown);

            // Step 2: Protect remaining HTML entities from being decoded again by Markdig
            var (protectedMarkdown, entityMap) = ProtectHtmlEntities(sanitizedMarkdown);

            // Step 3: Parse Markdown with protected entities
            var document = Markdig.Markdown.Parse(protectedMarkdown, _pipeline);
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
            var blocks = document.ToList();
            var elements = PostProcessBlocks(blocks);
            oeChildren.Add(elements);
            outline.Add(oeChildren);

            // Step 4: Restore protected HTML entities to their original form
            RestoreHtmlEntities(outline, entityMap);

            return outline;
        }

        public IEnumerable<XElement> ProcessBlocks(IEnumerable<Block> blocks)
        {
            return PostProcessBlocks(blocks.ToList());
        }

        private List<XElement> PostProcessBlocks(List<Block> blocks)
        {
            var elements = new List<XElement>();
            XElement lastContainerElement = null;

            for (int i = 0; i < blocks.Count; i++)
            {
                var block = blocks[i];
                if (block is LinkReferenceDefinitionGroup) continue;

                var processed = HandleBlock(block).ToList();

                // Special handling for lists following headings or paragraphs:
                // In Markdown, a list immediately after a paragraph/heading is semantically at the same level.
                // However, in OneNote's visual hierarchy, attaching the list as a child (OEChildren) of the
                // previous element creates the expected indented appearance without requiring manual spacing.
                // This mimics OneNote's native behavior where lists are often indented under their context.
                if (block is ListBlock && lastContainerElement != null)
                {
                    var childrenContainer = lastContainerElement.Element(OneNoteNamespace + "OEChildren");
                    if (childrenContainer == null)
                    {
                        childrenContainer = new XElement(OneNoteNamespace + "OEChildren");
                        lastContainerElement.Add(childrenContainer);
                    }
                    childrenContainer.Add(processed);
                }
                else
                {
                    elements.AddRange(processed);
                    lastContainerElement = (block is HeadingBlock || block is ParagraphBlock) ? processed.LastOrDefault() : null;
                }
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

        /// <summary>
        /// Protects HTML entities in the markdown text by replacing them with placeholders.
        /// This prevents Markdig from auto-decoding entities like &lt; to &lt;, which would
        /// cause double-escaping issues when HtmlEscaper re-encodes them.
        /// </summary>
        /// <param name="markdown">The markdown text containing HTML entities</param>
        /// <returns>A tuple of (protected markdown, entity map for restoration)</returns>
        private (string, Dictionary<string, string>) ProtectHtmlEntities(string markdown)
        {
            var entityMap = new Dictionary<string, string>();
            var counter = 0;

            var result = HtmlEntityRegex.Replace(markdown, match =>
            {
                // Use Unicode Replacement Character (U+FFFD) as placeholder delimiter
                // This character is extremely rare in normal text
                var placeholder = $"\uFFFD{counter++}\uFFFD";
                entityMap[placeholder] = match.Value;
                return placeholder;
            });

            return (result, entityMap);
        }

        /// <summary>
        /// Restores HTML entities in the generated OneNote XML by replacing placeholders
        /// with their original entity strings.
        /// </summary>
        /// <param name="outline">The OneNote Outline element to process</param>
        /// <param name="entityMap">The map of placeholders to original entities</param>
        private void RestoreHtmlEntities(XElement outline, Dictionary<string, string> entityMap)
        {
            if (entityMap.Count == 0) return; // Optimization: skip if no entities to restore

            var ns = OneNoteNamespace;
            foreach (var tElement in outline.Descendants(ns + "T"))
            {
                var cdata = tElement.Nodes().OfType<XCData>().FirstOrDefault();
                if (cdata == null) continue;

                var text = cdata.Value;
                var modified = false;

                foreach (var kvp in entityMap)
                {
                    if (text.Contains(kvp.Key))
                    {
                        text = text.Replace(kvp.Key, kvp.Value);
                        modified = true;
                    }
                }

                if (modified)
                {
                    cdata.ReplaceWith(new XCData(text));
                }
            }
        }

        public void IncrementQuoteDepth()
        {
            _quoteNestingDepth++;
        }

        public void DecrementQuoteDepth()
        {
            _quoteNestingDepth--;
        }

        public void PushWidthReservation(double reservedWidth)
        {
            _widthReservationStack.Push(reservedWidth);
        }

        public void PopWidthReservation()
        {
            if (_widthReservationStack.Count > 0)
            {
                _widthReservationStack.Pop();
            }
        }

        public string ConvertInlinesToHtml(ContainerInline container)
        {
            if (container == null) return string.Empty;
            var html = new StringBuilder();
            foreach (var inline in container)
            {
                // Skip TaskList inline elements (checkboxes are handled separately in ListHandler)
                if (inline is TaskList)
                {
                    continue;
                }

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
