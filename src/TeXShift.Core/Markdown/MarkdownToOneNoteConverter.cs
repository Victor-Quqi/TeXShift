using System;
using System.Collections.Generic;
using System.Linq;
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
using TeXShift.Core.Mermaid;
using TeXShift.Core.OneNoteMeta;

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
        private readonly MermaidBlockHandler _mermaidBlockHandler;
        private readonly MarkdownPipeline _pipeline;
        private readonly HtmlEntityProcessor _entityProcessor = new HtmlEntityProcessor();
        private readonly IInlineRenderer _inlineRenderer;
        private int _quoteNestingDepth = 0;
        private readonly Stack<double> _widthReservationStack = new Stack<double>();
        private readonly double _initialWidth;
        private Dictionary<string, string> _currentEntityMap;  // Stored during conversion for Math handlers

        // Explicit implementation of IMarkdownConverterContext properties
        public XNamespace OneNoteNamespace { get; } = "http://schemas.microsoft.com/office/onenote/2013/onenote";
        public OneNoteStyleConfig StyleConfig { get; }
        public IMathService MathService { get; }
        public IMermaidService MermaidService { get; }
        public MermaidRenderOptions MermaidOptions { get; }
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

        public MarkdownToOneNoteConverter(
            OneNoteStyleConfig styleConfig,
            MarkdownPipeline pipeline,
            IMathService mathService,
            IMermaidService mermaidService,
            MermaidRenderOptions mermaidOptions = null,
            double? sourceOutlineWidth = null)
        {
            StyleConfig = styleConfig ?? throw new ArgumentNullException(nameof(styleConfig));
            _pipeline = pipeline ?? throw new ArgumentNullException(nameof(pipeline));
            MathService = mathService;
            MermaidService = mermaidService;
            MermaidOptions = mermaidOptions;
            SourceOutlineWidth = sourceOutlineWidth;
            _initialWidth = sourceOutlineWidth ?? StyleConfig.GetQuoteBlockStyle().BaseWidth;

            // Create the inline renderer with dependencies
            _inlineRenderer = new InlineRenderer(styleConfig, mathService);

            // Create the Mermaid handler (used for fenced code blocks with mermaid info string)
            _mermaidBlockHandler = new MermaidBlockHandler(mermaidService);

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

            // Keep conversion off the caller thread (often UI/COM) while still allowing true async handlers.
            return await Task.Run(() => ConvertToOneNoteXmlInternalAsync(markdown)).ConfigureAwait(false);
        }

        private async Task<XElement> ConvertToOneNoteXmlInternalAsync(string markdown)
        {
            // Step 1: Sanitize text (remove OneNote formatting spans)
            var sanitizedMarkdown = MarkdownSanitizer.Sanitize(markdown);

            // Step 2: Normalize block-level syntax markers with minimal decoding
            // (e.g., &gt; → >, &lt; → <, &amp;lt; → &lt;)
            sanitizedMarkdown = MarkdownPreprocessor.Normalize(sanitizedMarkdown);

            // Step 3: Convert LaTeX delimiters to Markdown math syntax
            // (e.g., \(...\) → $...$, \[...\] → $$...$$)
            sanitizedMarkdown = LatexDelimiterConverter.Convert(sanitizedMarkdown);
            var sourceMarkdown = sanitizedMarkdown;

            // Step 4: Protect HTML entities from being decoded by Markdig
            var (protectedMarkdown, entityMap) = _entityProcessor.Protect(sanitizedMarkdown);
            _currentEntityMap = entityMap;  // Store for Math handlers to use
            _inlineRenderer.EntityDecoder = DecodeEntityPlaceholders;  // Set decoder for inline math

            // Step 5: Parse Markdown with protected entities
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
            var elements = await PostProcessBlocksAsync(blocks).ConfigureAwait(false);
            oeChildren.Add(elements);
            outline.Add(oeChildren);

            // Step 6: Restore and DECODE HTML entities (except in code blocks)
            _entityProcessor.RestoreAndDecode(outline, entityMap, OneNoteNamespace, StyleConfig);

            // Step 7: Persist source metadata for reverse conversion.
            TeXShiftMetaWriter.WriteSourceMeta(outline, sourceMarkdown, TeXShiftMetaKeys.ModeRender);

            return outline;
        }

        public async Task<IReadOnlyList<XElement>> ProcessBlocksAsync(IEnumerable<Block> blocks)
        {
            if (blocks == null) return Array.Empty<XElement>();
            return await PostProcessBlocksAsync(blocks.ToList()).ConfigureAwait(false);
        }

        [System.Obsolete("Use ProcessBlocksAsync instead. This method blocks on async work and may impact responsiveness.")]
        public IEnumerable<XElement> ProcessBlocks(IEnumerable<Block> blocks)
        {
            return ProcessBlocksAsync(blocks).ConfigureAwait(false).GetAwaiter().GetResult();
        }

        private async Task<IReadOnlyList<XElement>> PostProcessBlocksAsync(List<Block> blocks)
        {
            var elements = new List<XElement>();
            XElement lastContainerElement = null;

            for (int i = 0; i < blocks.Count; i++)
            {
                var block = blocks[i];
                if (block is LinkReferenceDefinitionGroup) continue;

                var processed = await HandleBlockAsync(block).ConfigureAwait(false);

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

        private async Task<IReadOnlyList<XElement>> HandleBlockAsync(Block block)
        {
            if (block is LinkReferenceDefinitionGroup) return Array.Empty<XElement>();

            if (block is FencedCodeBlock fenced)
            {
                var info = (fenced.Info ?? "").Trim();
                var language = info.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? "";
                if (string.Equals(language, "mermaid", StringComparison.OrdinalIgnoreCase))
                {
                    return await _mermaidBlockHandler.HandleAsync(block, this).ConfigureAwait(false);
                }
            }

            IBlockHandler handler;
            if (!_blockHandlers.TryGetValue(block.GetType(), out handler))
            {
                handler = _fallbackHandler;
            }
            return await handler.HandleAsync(block, this).ConfigureAwait(false);
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

        public string DecodeEntityPlaceholders(string text)
        {
            return HtmlEntityProcessor.DecodeForLatex(text, _currentEntityMap);
        }

        public Task<string> ConvertInlinesToHtmlAsync(ContainerInline container)
        {
            return _inlineRenderer.RenderAsync(container);
        }

        public Task<string> ConvertInlinesToHtmlAsync(IEnumerable<Inline> inlines)
        {
            return _inlineRenderer.RenderAsync(inlines);
        }

        [System.Obsolete("Use ConvertInlinesToHtmlAsync instead. This method blocks on async work and may impact responsiveness.")]
        public string ConvertInlinesToHtml(ContainerInline container)
        {
            return ConvertInlinesToHtmlAsync(container).ConfigureAwait(false).GetAwaiter().GetResult();
        }

        [System.Obsolete("Use ConvertInlinesToHtmlAsync instead. This method blocks on async work and may impact responsiveness.")]
        public string ConvertInlinesToHtml(IEnumerable<Inline> inlines)
        {
            return ConvertInlinesToHtmlAsync(inlines).ConfigureAwait(false).GetAwaiter().GetResult();
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
