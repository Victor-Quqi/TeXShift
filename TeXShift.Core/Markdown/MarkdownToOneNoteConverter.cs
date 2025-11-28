using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Xml.Linq;
using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Markdig.Extensions.Mathematics;
using Markdig.Extensions.Tables;
using TeXShift.Core.Abstractions;
using TeXShift.Core.Configuration;
using TeXShift.Core.Markdown.Abstractions;
using TeXShift.Core.Markdown.Handlers;
using TeXShift.Core.Markdown.Handlers.Inlines;
using TeXShift.Core.Markdown.Processing;
using TeXShift.Core.Math;

namespace TeXShift.Core.Markdown
{
    /// <summary>
    /// Converts Markdown text to OneNote XML format by dispatching to specialized block handlers.
    /// This class acts as a coordinator, parsing the Markdown and delegating the conversion
    /// of each block type to a registered handler.
    /// </summary>
    public class MarkdownToOneNoteConverter : IMarkdownConverter, IMarkdownConverterContext
    {
        private readonly Dictionary<Type, IBlockHandler> _blockHandlers;
        private readonly FallbackHandler _fallbackHandler = new FallbackHandler();
        private readonly MarkdownPipeline _pipeline;
        private readonly HtmlEntityProtector _entityProtector = new HtmlEntityProtector();
        private readonly IInlineRenderer _inlineRenderer;
        private int _quoteNestingDepth = 0;
        private readonly Stack<double> _widthReservationStack = new Stack<double>();
        private readonly double _initialWidth;

        // Explicit implementation of IMarkdownConverterContext properties
        public XNamespace OneNoteNamespace { get; } = "http://schemas.microsoft.com/office/onenote/2013/onenote";
        public OneNoteStyleConfig StyleConfig { get; }
        public IMathService MathService { get; }
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
                return System.Math.Max(available, 50.0);
            }
        }

        public MarkdownToOneNoteConverter(OneNoteStyleConfig styleConfig, MarkdownPipeline pipeline, IMathService mathService, double? sourceOutlineWidth = null)
        {
            StyleConfig = styleConfig ?? throw new ArgumentNullException(nameof(styleConfig));
            _pipeline = pipeline ?? throw new ArgumentNullException(nameof(pipeline));
            MathService = mathService;
            SourceOutlineWidth = sourceOutlineWidth;
            _initialWidth = sourceOutlineWidth ?? StyleConfig.GetQuoteBlockStyle().BaseWidth;

            // Create the inline renderer with dependencies
            _inlineRenderer = new InlineRenderer(styleConfig, mathService);

            // Register all the specialized handlers for each block type.
            // Note: CodeBlock = indented code (4-space), FencedCodeBlock = ```code```
            _blockHandlers = new Dictionary<Type, IBlockHandler>
            {
                { typeof(HeadingBlock), new HeadingHandler() },
                { typeof(ParagraphBlock), new ParagraphHandler() },
                { typeof(ListBlock), new ListHandler() },
                { typeof(CodeBlock), new CodeBlockHandler() },
                { typeof(FencedCodeBlock), new CodeBlockHandler() },
                { typeof(ThematicBreakBlock), new HorizontalRuleHandler() },
                { typeof(QuoteBlock), new QuoteBlockHandler() },
                { typeof(Table), new TableHandler() },
                { typeof(MathBlock), new MathBlockHandler(mathService) }
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
            // Step 1: Sanitize text (remove OneNote formatting spans)
            var sanitizedMarkdown = MarkdownSanitizer.Sanitize(markdown);

            // Step 2: HtmlDecode to restore Markdown syntax characters
            // (e.g., &gt; → >, &lt; → <, &amp;lt; → &lt;)
            sanitizedMarkdown = WebUtility.HtmlDecode(sanitizedMarkdown);

            // Step 3: Protect remaining HTML entities from being decoded again by Markdig
            var (protectedMarkdown, entityMap) = _entityProtector.Protect(sanitizedMarkdown);

            // Step 4: Parse Markdown with protected entities
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

            // Step 5: Restore protected HTML entities to their original form
            _entityProtector.Restore(outline, entityMap, OneNoteNamespace);

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

                // Lists get nested under the preceding container element (heading, paragraph, or code block)
                // This preserves document order while providing consistent indentation
                if (block is ListBlock && lastContainerElement != null)
                {
                    var childrenContainer = lastContainerElement.Element(OneNoteNamespace + "OEChildren");
                    if (childrenContainer == null)
                    {
                        childrenContainer = new XElement(OneNoteNamespace + "OEChildren");
                        lastContainerElement.Add(childrenContainer);
                    }
                    childrenContainer.Add(processed);
                    // Lists don't become containers - keep the current container for subsequent lists
                }
                else
                {
                    elements.AddRange(processed);
                    // All block types (except lists) can serve as containers for subsequent lists
                    // This allows lists to maintain consistent indentation regardless of what precedes them
                    lastContainerElement = processed.LastOrDefault();
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
            return _inlineRenderer.Render(container);
        }

        public string ConvertInlinesToHtml(IEnumerable<Inline> inlines)
        {
            return _inlineRenderer.Render(inlines);
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
