using System;
using Markdig.Extensions.Mathematics;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TeXShift.Core.Markdown;
using TeXShift.Core.Math;

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

            // Check if paragraph contains display math ($$...$$) that should be split into separate blocks
            var mathSegments = SplitParagraphByDisplayMath(paragraph);
            if (mathSegments.Count > 1 || (mathSegments.Count == 1 && mathSegments[0].IsDisplayMath))
            {
                return HandleMathParagraph(mathSegments, context, ns, styleConfig);
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
        /// Represents a segment of a paragraph - either text content, a standalone image, or display math.
        /// </summary>
        private class ParagraphSegment
        {
            public bool IsImage { get; set; }
            public bool IsDisplayMath { get; set; }
            public LinkInline ImageLink { get; set; }
            public MathInline MathInline { get; set; }
            public List<Inline> TextInlines { get; set; }
        }

        /// <summary>
        /// Splits a paragraph into segments, separating display math ($$...$$) from other content.
        /// Each display math becomes its own centered block.
        /// </summary>
        private List<ParagraphSegment> SplitParagraphByDisplayMath(ParagraphBlock paragraph)
        {
            var segments = new List<ParagraphSegment>();
            if (paragraph.Inline == null) return segments;

            var currentTextInlines = new List<Inline>();
            var inlines = paragraph.Inline.ToList();

            foreach (var inline in inlines)
            {
                if (inline is MathInline mathInline && mathInline.DelimiterCount == 2)
                {
                    // Save current text segment if not empty
                    if (currentTextInlines.Any(IsNonEmptyInline))
                    {
                        segments.Add(new ParagraphSegment { TextInlines = new List<Inline>(currentTextInlines) });
                    }
                    currentTextInlines.Clear();

                    // Add display math segment
                    segments.Add(new ParagraphSegment { IsDisplayMath = true, MathInline = mathInline });
                }
                else if (inline is LineBreakInline)
                {
                    // Skip line breaks between display math elements
                    // Only add if we have non-math content
                    if (currentTextInlines.Any(IsNonEmptyInline))
                    {
                        currentTextInlines.Add(inline);
                    }
                }
                else
                {
                    currentTextInlines.Add(inline);
                }
            }

            // Add remaining text segment if not empty
            if (currentTextInlines.Any(IsNonEmptyInline))
            {
                segments.Add(new ParagraphSegment { TextInlines = currentTextInlines });
            }

            return segments;
        }

        /// <summary>
        /// Handles a paragraph with display math, creating separate centered OE elements for each formula.
        /// </summary>
        private IEnumerable<XElement> HandleMathParagraph(List<ParagraphSegment> segments, IMarkdownConverterContext context, XNamespace ns, OneNoteStyleConfig styleConfig)
        {
            var results = new List<XElement>();
            var spacing = styleConfig.GetParagraphSpacing();

            foreach (var segment in segments)
            {
                if (segment.IsDisplayMath)
                {
                    // Create centered OE for display math
                    var oe = new XElement(ns + "OE",
                        new XAttribute("alignment", "center"),
                        new XAttribute("spaceBefore", "8.8"),
                        new XAttribute("spaceAfter", "8.8"));

                    // Convert the math inline directly
                    var mathHtml = ConvertDisplayMathToHtml(segment.MathInline, context);
                    oe.Add(new XElement(ns + "T", new XCData(mathHtml)));

                    results.Add(oe);
                }
                else if (segment.TextInlines != null && segment.TextInlines.Any())
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

        /// <summary>
        /// Converts a display math MathInline to HTML/MathML for OneNote.
        /// </summary>
        private string ConvertDisplayMathToHtml(MathInline mathInline, IMarkdownConverterContext context)
        {
            // Get MathService from context if available
            var mathService = GetMathService(context);
            if (mathService == null)
            {
                return $"$${mathInline.Content}$$";
            }

            // Auto-initialize MathService if needed
            if (!mathService.IsInitialized)
            {
                try
                {
                    mathService.InitializeAsync().GetAwaiter().GetResult();
                }
                catch (Exception ex)
                {
                    return $"[MathInit Error: {ex.Message}]";
                }
            }

            try
            {
                var latex = mathInline.Content.ToString();
                var mathml = mathService.LatexToMathMLAsync(latex, displayMode: true).GetAwaiter().GetResult();
                return mathService.WrapMathMLForOneNote(mathml);
            }
            catch
            {
                return $"[LaTeX Error: $${mathInline.Content}$$]";
            }
        }

        /// <summary>
        /// Gets the MathService from the context. Returns null if not available.
        /// </summary>
        private IMathService GetMathService(IMarkdownConverterContext context)
        {
            // Access MathService through reflection since it's not in the interface
            var converterType = context.GetType();
            var field = converterType.GetField("_mathService", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            return field?.GetValue(context) as IMathService;
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
