using System;
using Markdig.Extensions.Mathematics;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using TeXShift.Core.Configuration;
using TeXShift.Core.Errors;
using TeXShift.Core.Localization;
using TeXShift.Core.Markdown.Abstractions;

namespace TeXShift.Core.Markdown.Handlers
{
    internal class ParagraphHandler : IBlockHandler
    {
        public async Task<IReadOnlyList<XElement>> HandleAsync(Block block, IMarkdownConverterContext context)
        {
            var paragraph = (ParagraphBlock)block;
            var ns = context.OneNoteNamespace;
            var styleConfig = context.StyleConfig;

            // Check if paragraph contains only a single image
            var singleImage = ImageElementHelper.GetSingleImage(paragraph);
            if (singleImage != null)
            {
                var imageOe = await ImageElementHelper.CreateImageOEAsync(singleImage, ns, context).ConfigureAwait(false);
                return new[] { imageOe };
            }

            // Check if paragraph contains display math ($$...$$) that should be split into separate blocks
            var mathSegments = SplitParagraphByDisplayMath(paragraph);
            if (mathSegments.Count > 1 || (mathSegments.Count == 1 && mathSegments[0].IsDisplayMath))
            {
                return await HandleMathParagraphAsync(mathSegments, context, ns, styleConfig).ConfigureAwait(false);
            }

            // Check if paragraph contains standalone image lines mixed with text
            var segments = SplitParagraphByStandaloneImages(paragraph);
            if (segments.Count > 1)
            {
                return await HandleMixedParagraphAsync(segments, context, ns, styleConfig).ConfigureAwait(false);
            }

            var oe = new XElement(ns + "OE");

            // Apply paragraph spacing
            var spacing = styleConfig.GetParagraphSpacing();
            oe.Add(new XAttribute("spaceBefore", spacing.SpaceBefore.ToString("F1")));
            oe.Add(new XAttribute("spaceAfter", spacing.SpaceAfter.ToString("F1")));
            oe.Add(new XAttribute("spaceBetween", spacing.SpaceBetween.ToString("F1")));

            // Convert inline content to HTML
            var htmlContent = await context.ConvertInlinesToHtmlAsync(paragraph.Inline).ConfigureAwait(false);
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
        private async Task<IReadOnlyList<XElement>> HandleMathParagraphAsync(List<ParagraphSegment> segments, IMarkdownConverterContext context, XNamespace ns, OneNoteStyleConfig styleConfig)
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
                    var mathHtml = await ConvertDisplayMathToHtmlAsync(segment.MathInline, context).ConfigureAwait(false);
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

                    var htmlContent = await context.ConvertInlinesToHtmlAsync(segment.TextInlines).ConfigureAwait(false);
                    oe.Add(new XElement(ns + "T", new XCData(htmlContent)));

                    results.Add(oe);
                }
            }

            return results;
        }

         /// <summary>
         /// Converts a display math MathInline to HTML/MathML for OneNote.
         /// </summary>
        private async Task<string> ConvertDisplayMathToHtmlAsync(MathInline mathInline, IMarkdownConverterContext context)
        {
            // Get MathService from context
            var mathService = context.MathService;
            if (mathService == null)
            {
                return $"$${mathInline.Content}$$";
            }

             // Auto-initialize MathService if needed
             if (!mathService.IsInitialized)
             {
                 try
                 {
                    await mathService.InitializeAsync().ConfigureAwait(false);
                 }
                 catch (Exception ex)
                 {
                     throw new MathConversionException(
                         Resources.GetString("Error_MathInitFailed"),
                        $"MathService initialization failed in ParagraphHandler. {ex.GetType().Name}: {ex.Message}",
                        ex);
                }
            }

            try
             {
                  var latex = mathInline.Content.ToString();
                  // Decode HTML entity placeholders before passing to MathJax
                  latex = context.DecodeEntityPlaceholders(latex);
                var mathml = await mathService.LatexToMathMLAsync(latex, displayMode: true).ConfigureAwait(false);
                  return mathService.WrapMathMLForOneNote(mathml);
              }
              catch (Exception ex) when (!(ex is MathConversionException))
              {
                throw new MathConversionException(
                    Resources.GetString("Error_MathConversionFailed"),
                    $"LaTeX conversion failed for: {mathInline.Content}. {ex.GetType().Name}: {ex.Message}",
                    ex);
            }
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
                i = ProcessInlineForImageSplit(inlines, i, segments, currentTextInlines);
            }

            // Add remaining text segment if not empty
            AddTextSegmentIfNotEmpty(segments, currentTextInlines);

            return segments;
        }

        /// <summary>
        /// Processes a single inline element during image splitting. Returns the new index position.
        /// </summary>
        private int ProcessInlineForImageSplit(List<Inline> inlines, int i, List<ParagraphSegment> segments, List<Inline> currentTextInlines)
        {
            var inline = inlines[i];

            // Check for soft break followed by standalone image
            if (inline is LineBreakInline lineBreak && !lineBreak.IsHard)
            {
                if (TryAddStandaloneImageAfterBreak(inlines, i, segments, currentTextInlines, out var endIndex))
                {
                    return endIndex;
                }
                currentTextInlines.Add(inline);
                return i;
            }

            // Check for standalone image at paragraph start
            if (i == 0 && IsStandaloneImageLine(inlines, 0, out var firstImageLink, out var firstEndIndex))
            {
                segments.Add(new ParagraphSegment { IsImage = true, ImageLink = firstImageLink });
                return firstEndIndex;
            }

            currentTextInlines.Add(inline);
            return i;
        }

        /// <summary>
        /// Tries to add a standalone image segment after a soft break. Returns true if successful.
        /// </summary>
        private bool TryAddStandaloneImageAfterBreak(List<Inline> inlines, int breakIndex, List<ParagraphSegment> segments, List<Inline> currentTextInlines, out int endIndex)
        {
            endIndex = breakIndex;

            if (breakIndex + 1 >= inlines.Count) return false;
            if (!IsStandaloneImageLine(inlines, breakIndex + 1, out var imageLink, out endIndex)) return false;

            AddTextSegmentIfNotEmpty(segments, currentTextInlines);
            currentTextInlines.Clear();
            segments.Add(new ParagraphSegment { IsImage = true, ImageLink = imageLink });

            return true;
        }

        /// <summary>
        /// Adds a text segment to the list if it contains non-empty inlines.
        /// </summary>
        private void AddTextSegmentIfNotEmpty(List<ParagraphSegment> segments, List<Inline> textInlines)
        {
            if (textInlines.Any(IsNonEmptyInline))
            {
                segments.Add(new ParagraphSegment { IsImage = false, TextInlines = new List<Inline>(textInlines) });
            }
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
        private async Task<IReadOnlyList<XElement>> HandleMixedParagraphAsync(List<ParagraphSegment> segments, IMarkdownConverterContext context, XNamespace ns, OneNoteStyleConfig styleConfig)
        {
            var results = new List<XElement>();
            var spacing = styleConfig.GetParagraphSpacing();

            foreach (var segment in segments)
            {
                if (segment.IsImage)
                {
                    // Handle as standalone image using shared helper
                    results.Add(await ImageElementHelper.CreateImageOEAsync(segment.ImageLink, ns, context).ConfigureAwait(false));
                }
                else
                {
                    // Handle as text paragraph
                    var oe = new XElement(ns + "OE");
                    oe.Add(new XAttribute("spaceBefore", spacing.SpaceBefore.ToString("F1")));
                    oe.Add(new XAttribute("spaceAfter", spacing.SpaceAfter.ToString("F1")));
                    oe.Add(new XAttribute("spaceBetween", spacing.SpaceBetween.ToString("F1")));

                    var htmlContent = await context.ConvertInlinesToHtmlAsync(segment.TextInlines).ConfigureAwait(false);
                    oe.Add(new XElement(ns + "T", new XCData(htmlContent)));

                    results.Add(oe);
                }
            }

            return results;
        }
    }
}
