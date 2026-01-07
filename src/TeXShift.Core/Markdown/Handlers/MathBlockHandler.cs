using System;
using System.Collections.Generic;
using System.Xml.Linq;
using Markdig.Extensions.Mathematics;
using Markdig.Syntax;
using TeXShift.Core.Errors;
using TeXShift.Core.Localization;
using TeXShift.Core.Markdown.Abstractions;
using TeXShift.Core.Math;

namespace TeXShift.Core.Markdown.Handlers
{
    /// <summary>
    /// Handler for converting block-level math expressions ($$...$$) to OneNote MathML format.
    /// </summary>
    internal class MathBlockHandler : IBlockHandler
    {
        private readonly IMathService _mathService;

        public MathBlockHandler(IMathService mathService)
        {
            _mathService = mathService;
        }

        public IEnumerable<XElement> Handle(Block block, IMarkdownConverterContext context)
        {
            var mathBlock = block as MathBlock;
            if (mathBlock == null)
            {
                return System.Array.Empty<XElement>();
            }

            // Extract LaTeX content from the math block
            var latex = ExtractLatexContent(mathBlock);
            if (string.IsNullOrWhiteSpace(latex))
            {
                return System.Array.Empty<XElement>();
            }

            // Decode HTML entity placeholders before passing to MathJax
            latex = context.DecodeEntityPlaceholders(latex);

            // Check if MathService is available
            if (_mathService == null)
            {
                var fallbackOe = new XElement(context.OneNoteNamespace + "OE",
                    new XAttribute("alignment", "center"),
                    new XElement(context.OneNoteNamespace + "T",
                        new XCData($"$${latex}$$")));
                return new[] { fallbackOe };
            }

            // Auto-initialize MathService if needed
            if (!_mathService.IsInitialized)
            {
                try
                {
                    _mathService.InitializeAsync().ConfigureAwait(false).GetAwaiter().GetResult();
                }
                catch (Exception ex) when (!(ex is MathConversionException))
                {
                    throw new MathConversionException(
                        Resources.GetString("Error_MathInitFailed"),
                        $"MathService initialization failed in MathBlockHandler. {ex.GetType().Name}: {ex.Message}",
                        ex);
                }
            }

            // Convert LaTeX to MathML using MathJax
            string mathml;
            try
            {
                // Use displayMode: true for block-level math
                mathml = _mathService.LatexToMathMLAsync(latex, displayMode: true).ConfigureAwait(false).GetAwaiter().GetResult();
            }
            catch (Exception ex) when (!(ex is MathConversionException))
            {
                throw new MathConversionException(
                    Resources.GetString("Error_MathConversionFailed"),
                    $"LaTeX conversion failed for block math: {latex}. {ex.GetType().Name}: {ex.Message}",
                    ex);
            }

            // Wrap MathML for OneNote
            var wrappedMathml = _mathService.WrapMathMLForOneNote(mathml);

            // Create OneNote OE element with proper styling for block math
            // Block math should be centered with math font
            var oe = new XElement(context.OneNoteNamespace + "OE",
                new XAttribute("alignment", "center"),
                new XAttribute("spaceBefore", "8.8"),
                new XAttribute("spaceAfter", "8.8"),
                new XElement(context.OneNoteNamespace + "T",
                    new XCData(wrappedMathml)));

            return new[] { oe };
        }

        private string ExtractLatexContent(MathBlock mathBlock)
        {
            // MathBlock stores content in Lines
            if (mathBlock.Lines.Lines == null || mathBlock.Lines.Count == 0)
            {
                return string.Empty;
            }

            var lines = new List<string>();
            foreach (var line in mathBlock.Lines.Lines)
            {
                if (line.Slice.Text != null)
                {
                    lines.Add(line.Slice.ToString());
                }
            }

            return string.Join("\n", lines).Trim();
        }
    }
}
