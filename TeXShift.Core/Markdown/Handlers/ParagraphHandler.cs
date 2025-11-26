using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using System.Collections.Generic;
using System.Linq;
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

            // Check if paragraph contains only a single image
            var singleImage = ImageElementHelper.GetSingleImage(paragraph);
            if (singleImage != null)
            {
                return new[] { ImageElementHelper.CreateImageOE(singleImage, ns) };
            }

            // Check if paragraph contains standalone image lines mixed with text
            var segments = SplitParagraphByStandaloneImages(paragraph);
            if (segments.Count > 1)
            {
                return HandleMixedParagraph(segments, context, ns, styleConfig);
            }

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

        /// <summary>
        /// Represents a segment of a paragraph - either text content or a standalone image.
        /// </summary>
        private class ParagraphSegment
        {
            public bool IsImage { get; set; }
            public LinkInline ImageLink { get; set; }
            public List<Inline> TextInlines { get; set; }
        }

        /// <summary>
        /// Splits a paragraph into segments, separating standalone image lines from text.
        /// </summary>
        private List<ParagraphSegment> SplitParagraphByStandaloneImages(ParagraphBlock paragraph)
        {
            var segments = new List<ParagraphSegment>();
            if (paragraph.Inline == null) return segments;

            var currentTextInlines = new List<Inline>();
            var inlines = paragraph.Inline.ToList();

            for (int i = 0; i < inlines.Count; i++)
            {
                var inline = inlines[i];

                if (inline is LineBreakInline lineBreak && !lineBreak.IsHard)
                {
                    // Check if this soft break is followed by a standalone image
                    if (i + 1 < inlines.Count && IsStandaloneImageLine(inlines, i + 1, out var imageLink, out var endIndex))
                    {
                        // Save current text segment if not empty
                        if (currentTextInlines.Any(IsNonEmptyInline))
                        {
                            segments.Add(new ParagraphSegment { IsImage = false, TextInlines = new List<Inline>(currentTextInlines) });
                        }
                        currentTextInlines.Clear();

                        // Add image segment
                        segments.Add(new ParagraphSegment { IsImage = true, ImageLink = imageLink });

                        // Skip to after the image line
                        i = endIndex;
                        continue;
                    }
                    else
                    {
                        currentTextInlines.Add(inline);
                    }
                }
                else if (i == 0 && IsStandaloneImageLine(inlines, 0, out var firstImageLink, out var firstEndIndex))
                {
                    // Paragraph starts with a standalone image
                    segments.Add(new ParagraphSegment { IsImage = true, ImageLink = firstImageLink });
                    i = firstEndIndex;
                }
                else
                {
                    currentTextInlines.Add(inline);
                }
            }

            // Add remaining text segment if not empty
            if (currentTextInlines.Any(IsNonEmptyInline))
            {
                segments.Add(new ParagraphSegment { IsImage = false, TextInlines = currentTextInlines });
            }

            return segments;
        }

        /// <summary>
        /// Checks if position marks the start of a standalone image line.
        /// Returns the image link and the end index of this image line.
        /// </summary>
        private bool IsStandaloneImageLine(List<Inline> inlines, int startIndex, out LinkInline imageLink, out int endIndex)
        {
            imageLink = null;
            endIndex = startIndex;

            if (startIndex >= inlines.Count) return false;

            // Skip leading whitespace
            int pos = startIndex;
            while (pos < inlines.Count && inlines[pos] is LiteralInline lit && string.IsNullOrWhiteSpace(lit.Content.ToString()))
            {
                pos++;
            }

            // Must have a LinkInline with IsImage
            if (pos >= inlines.Count || !(inlines[pos] is LinkInline link) || !link.IsImage)
                return false;

            imageLink = link;
            pos++;

            // Skip trailing whitespace
            while (pos < inlines.Count && inlines[pos] is LiteralInline trailingLit && string.IsNullOrWhiteSpace(trailingLit.Content.ToString()))
            {
                pos++;
            }

            // Must be followed by soft line break or end of inlines
            if (pos >= inlines.Count)
            {
                endIndex = pos - 1;
                return true;
            }

            if (inlines[pos] is LineBreakInline lb && !lb.IsHard)
            {
                endIndex = pos;
                return true;
            }

            return false;
        }

        /// <summary>
        /// Checks if an inline is non-empty (not whitespace-only literal or soft break).
        /// </summary>
        private bool IsNonEmptyInline(Inline inline)
        {
            if (inline is LineBreakInline lb && !lb.IsHard) return false;
            if (inline is LiteralInline lit && string.IsNullOrWhiteSpace(lit.Content.ToString())) return false;
            return true;
        }

        /// <summary>
        /// Handles a paragraph with mixed text and standalone image segments.
        /// </summary>
        private IEnumerable<XElement> HandleMixedParagraph(List<ParagraphSegment> segments, IMarkdownConverterContext context, XNamespace ns, OneNoteStyleConfig styleConfig)
        {
            var results = new List<XElement>();
            var spacing = styleConfig.GetParagraphSpacing();

            foreach (var segment in segments)
            {
                if (segment.IsImage)
                {
                    // Handle as standalone image using shared helper
                    results.Add(ImageElementHelper.CreateImageOE(segment.ImageLink, ns));
                }
                else
                {
                    // Handle as text paragraph
                    var oe = new XElement(ns + "OE");
                    oe.Add(new XAttribute("spaceBefore", spacing.SpaceBefore.ToString("F1")));
                    oe.Add(new XAttribute("spaceAfter", spacing.SpaceAfter.ToString("F1")));
                    oe.Add(new XAttribute("spaceBetween", spacing.SpaceBetween.ToString("F1")));

                    var htmlContent = context.ConvertInlinesToHtml(segment.TextInlines);
                    oe.Add(new XElement(ns + "T", new XCData(htmlContent)));

                    results.Add(oe);
                }
            }

            return results;
        }
    }
}
