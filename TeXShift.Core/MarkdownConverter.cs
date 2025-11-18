using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace TeXShift.Core
{
    /// <summary>
    /// Converts Markdown text to OneNote XML format.
    /// </summary>
    public class MarkdownConverter : IMarkdownConverter
    {
        private readonly XNamespace _ns = "http://schemas.microsoft.com/office/onenote/2013/onenote";
        private readonly MarkdownPipeline _pipeline;
        private readonly OneNoteStyleConfig _styleConfig;

        /// <summary>
        /// Creates a new MarkdownConverter with default configuration.
        /// Note: For better performance, use the ServiceContainer which caches the pipeline.
        /// </summary>
        public MarkdownConverter() : this(new OneNoteStyleConfig(), null)
        {
        }

        /// <summary>
        /// Creates a new MarkdownConverter with specified configuration and pipeline.
        /// </summary>
        /// <param name="styleConfig">Style configuration for OneNote elements.</param>
        /// <param name="pipeline">Optional shared MarkdownPipeline instance. If null, creates a new one.</param>
        public MarkdownConverter(OneNoteStyleConfig styleConfig, MarkdownPipeline pipeline = null)
        {
            _styleConfig = styleConfig ?? throw new ArgumentNullException(nameof(styleConfig));

            // Use provided pipeline or create new one
            _pipeline = pipeline ?? new MarkdownPipelineBuilder()
                .UseAdvancedExtensions()
                .Build();
        }

        /// <summary>
        /// Applies spacing attributes to an OE element.
        /// </summary>
        private void ApplySpacing(XElement oe, OneNoteStyleConfig.SpacingConfig spacing)
        {
            oe.Add(new XAttribute("spaceBefore", spacing.SpaceBefore.ToString("F1")));
            oe.Add(new XAttribute("spaceAfter", spacing.SpaceAfter.ToString("F1")));
            oe.Add(new XAttribute("spaceBetween", spacing.SpaceBetween.ToString("F1")));
        }

        /// <summary>
        /// Asynchronously converts Markdown text to a OneNote Outline XML element.
        /// </summary>
        public async Task<XElement> ConvertToOneNoteXmlAsync(string markdown)
        {
            if (string.IsNullOrWhiteSpace(markdown))
            {
                return CreateEmptyOutline();
            }

            // Wrap CPU-intensive Markdown parsing in Task.Run to avoid blocking UI
            return await Task.Run(() => ConvertToOneNoteXml(markdown)).ConfigureAwait(false);
        }

        /// <summary>
        /// Converts Markdown text to a OneNote Outline XML element.
        /// (Synchronous version - kept for internal use)
        /// </summary>
        private XElement ConvertToOneNoteXml(string markdown)
        {
            if (string.IsNullOrWhiteSpace(markdown))
            {
                return CreateEmptyOutline();
            }

            var document = Markdown.Parse(markdown, _pipeline);
            var outline = new XElement(_ns + "Outline");
            var oeChildren = new XElement(_ns + "OEChildren");

            XElement currentHeading = null;
            XElement currentOEChildren = null;

            foreach (var block in document)
            {
                // Skip internal blocks
                if (block is Markdig.Syntax.LinkReferenceDefinitionGroup)
                {
                    continue;
                }

                if (block is HeadingBlock heading)
                {
                    // Close previous heading's children if any
                    if (currentHeading != null && currentOEChildren != null && currentOEChildren.HasElements)
                    {
                        currentHeading.Add(currentOEChildren);
                    }

                    // Create new heading
                    currentHeading = ConvertHeading(heading);
                    currentOEChildren = new XElement(_ns + "OEChildren");
                    oeChildren.Add(currentHeading);
                }
                else if (block is ListBlock list)
                {
                    // Add list items to current heading's children
                    var listItems = ConvertList(list);
                    if (currentOEChildren != null)
                    {
                        foreach (var item in listItems)
                        {
                            currentOEChildren.Add(item);
                        }
                    }
                    else
                    {
                        // No heading, add list items directly
                        foreach (var item in listItems)
                        {
                            oeChildren.Add(item);
                        }
                    }
                }
                else
                {
                    // Close previous heading's children if any
                    if (currentHeading != null && currentOEChildren != null && currentOEChildren.HasElements)
                    {
                        currentHeading.Add(currentOEChildren);
                        currentHeading = null;
                        currentOEChildren = null;
                    }

                    // Add other blocks (paragraphs, code, etc.)
                    var oeElements = ConvertBlock(block);
                    if (oeElements != null)
                    {
                        foreach (var oe in oeElements)
                        {
                            oeChildren.Add(oe);
                        }
                    }
                }
            }

            // Close final heading's children if any
            if (currentHeading != null && currentOEChildren != null && currentOEChildren.HasElements)
            {
                currentHeading.Add(currentOEChildren);
            }

            outline.Add(oeChildren);
            return outline;
        }

        private XElement CreateEmptyOutline()
        {
            return new XElement(_ns + "Outline",
                new XElement(_ns + "OEChildren",
                    new XElement(_ns + "OE",
                        new XElement(_ns + "T", new XCData(""))
                    )
                )
            );
        }

        private System.Collections.Generic.List<XElement> ConvertBlock(Block block)
        {
            var result = new System.Collections.Generic.List<XElement>();

            // Skip Markdig internal blocks
            if (block is Markdig.Syntax.LinkReferenceDefinitionGroup)
            {
                return result;
            }

            if (block is HeadingBlock heading)
            {
                result.Add(ConvertHeading(heading));
            }
            else if (block is ParagraphBlock paragraph)
            {
                result.Add(ConvertParagraph(paragraph));
            }
            else if (block is ListBlock list)
            {
                result.AddRange(ConvertList(list));
            }
            else if (block is CodeBlock code)
            {
                result.Add(ConvertCodeBlock(code));
            }
            else
            {
                result.Add(ConvertFallback(block));
            }

            return result;
        }

        private XElement ConvertHeading(HeadingBlock heading)
        {
            var oe = new XElement(_ns + "OE");

            // Map heading levels to OneNote styles
            int styleIndex = Math.Min(heading.Level - 1, 5);
            oe.Add(new XAttribute("quickStyleIndex", styleIndex.ToString()));

            // Apply spacing based on heading level
            ApplySpacing(oe, _styleConfig.GetHeadingSpacing(heading.Level));

            // Get font configuration for this heading level
            var fontConfig = _styleConfig.GetHeadingFont(heading.Level);

            // Convert inline content to HTML and apply font styles
            var htmlContent = ConvertInlinesToHtml(heading.Inline);
            var styleAttributes = $"font-size:{fontConfig.FontSize}pt";
            if (fontConfig.IsBold)
            {
                styleAttributes += ";font-weight:bold";
            }
            var styledHeading = $"<span style='{styleAttributes}'>{htmlContent}</span>";
            oe.Add(new XElement(_ns + "T", new XCData(styledHeading)));

            return oe;
        }

        private XElement ConvertParagraph(ParagraphBlock paragraph)
        {
            var oe = new XElement(_ns + "OE");

            // Apply paragraph spacing
            ApplySpacing(oe, _styleConfig.GetParagraphSpacing());

            // Convert inline content to HTML
            var htmlContent = ConvertInlinesToHtml(paragraph.Inline);
            oe.Add(new XElement(_ns + "T", new XCData(htmlContent)));

            return oe;
        }

        private System.Collections.Generic.List<XElement> ConvertList(ListBlock list)
        {
            var result = new System.Collections.Generic.List<XElement>();

            foreach (var item in list.OfType<ListItemBlock>())
            {
                foreach (var childBlock in item)
                {
                    if (childBlock is ParagraphBlock para)
                    {
                        var oe = new XElement(_ns + "OE");

                        // Apply list item spacing
                        ApplySpacing(oe, _styleConfig.GetListSpacing());

                        // Add List element for bullet/number
                        var listElement = new XElement(_ns + "List");
                        if (list.IsOrdered)
                        {
                            // Ordered list: <Number>
                            var number = new XElement(_ns + "Number",
                                new XAttribute("numberSequence", "0"),
                                new XAttribute("numberFormat", "##."),
                                new XAttribute("fontSize", "11.0")
                            );
                            listElement.Add(number);
                        }
                        else
                        {
                            // Unordered list: <Bullet>
                            var bullet = new XElement(_ns + "Bullet",
                                new XAttribute("bullet", "2"),
                                new XAttribute("fontSize", "11.0")
                            );
                            listElement.Add(bullet);
                        }
                        oe.Add(listElement);

                        // Add text content
                        var htmlContent = ConvertInlinesToHtml(para.Inline);
                        oe.Add(new XElement(_ns + "T", new XCData(htmlContent)));

                        result.Add(oe);
                    }
                }
            }

            return result;
        }

        private XElement ConvertCodeBlock(CodeBlock code)
        {
            var oe = new XElement(_ns + "OE");

            // Apply code block spacing
            ApplySpacing(oe, _styleConfig.GetCodeSpacing());

            var textElement = new XElement(_ns + "T",
                new XCData($"<span style='font-family:Consolas'>{EscapeHtml(code.Lines.ToString())}</span>")
            );
            oe.Add(textElement);
            return oe;
        }

        private XElement ConvertFallback(Block block)
        {
            var oe = new XElement(_ns + "OE");
            oe.Add(new XElement(_ns + "T", new XCData(block.ToString())));
            return oe;
        }

        /// <summary>
        /// Converts inline elements to HTML format for OneNote CDATA.
        /// </summary>
        private string ConvertInlinesToHtml(ContainerInline container)
        {
            if (container == null) return string.Empty;

            var html = new StringBuilder();

            foreach (var inline in container)
            {
                if (inline is LiteralInline literal)
                {
                    html.Append(EscapeHtml(literal.Content.ToString()));
                }
                else if (inline is EmphasisInline emphasis)
                {
                    var content = ConvertInlinesToHtml(emphasis);

                    if (emphasis.DelimiterChar == '*' || emphasis.DelimiterChar == '_')
                    {
                        if (emphasis.DelimiterCount == 2)
                        {
                            // Bold: **text** or __text__
                            html.Append($"<span style='font-weight:bold'>{content}</span>");
                        }
                        else if (emphasis.DelimiterCount == 1)
                        {
                            // Italic: *text* or _text_
                            html.Append($"<span style='font-style:italic'>{content}</span>");
                        }
                        else
                        {
                            html.Append(content);
                        }
                    }
                    else if (emphasis.DelimiterChar == '~' && emphasis.DelimiterCount == 2)
                    {
                        // Strikethrough: ~~text~~
                        html.Append($"<span style='text-decoration:line-through'>{content}</span>");
                    }
                    else
                    {
                        html.Append(content);
                    }
                }
                else if (inline is CodeInline code)
                {
                    html.Append($"<span style='font-family:Consolas'>{EscapeHtml(code.Content)}</span>");
                }
                else if (inline is LineBreakInline)
                {
                    html.Append("\n");
                }
                else if (inline is ContainerInline nestedContainer)
                {
                    html.Append(ConvertInlinesToHtml(nestedContainer));
                }
            }

            return html.ToString();
        }

        /// <summary>
        /// Escapes HTML special characters.
        /// </summary>
        private string EscapeHtml(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            return text
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&#39;");
        }
    }
}
