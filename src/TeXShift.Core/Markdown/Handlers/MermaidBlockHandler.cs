using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Markdig.Syntax;
using TeXShift.Core.Markdown.Abstractions;
using TeXShift.Core.Mermaid;
using TeXShift.Core.Localization;
using TeXShift.Core.Logging;
using TeXShift.Core.Utils;

namespace TeXShift.Core.Markdown.Handlers
{
    internal class MermaidBlockHandler : IBlockHandler
    {
        private readonly IMermaidService _mermaidService;
        private readonly CodeBlockHandler _codeBlockFallback = new CodeBlockHandler();

        public MermaidBlockHandler(IMermaidService mermaidService)
        {
            _mermaidService = mermaidService;
        }

        public async Task<IReadOnlyList<XElement>> HandleAsync(Block block, IMarkdownConverterContext context)
        {
            var fenced = block as FencedCodeBlock;
            if (fenced == null)
            {
                return Array.Empty<XElement>();
            }

            var code = ExtractCode(fenced);
            if (string.IsNullOrWhiteSpace(code))
            {
                return await _codeBlockFallback.HandleAsync(block, context).ConfigureAwait(false);
            }

            // Decode HTML entity placeholders before passing to Mermaid
            code = context.DecodeEntityPlaceholders(code);

            if (_mermaidService == null)
            {
                return await CreateFallbackAsync(
                    block,
                    context,
                    "Mermaid rendering unavailable: the service is not configured.").ConfigureAwait(false);
            }

            // Auto-initialize MermaidService if needed
            if (!_mermaidService.IsInitialized)
            {
                try
                {
                    await _mermaidService.InitializeAsync().ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    return await CreateFallbackAsync(
                        block,
                        context,
                        $"Mermaid service initialization failed: {ex}").ConfigureAwait(false);
                }
            }

            MermaidRenderResult result;
            try
            {
                result = await _mermaidService.RenderToImageAsync(code, context.MermaidOptions).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                return await CreateFallbackAsync(
                    block,
                    context,
                    $"Mermaid rendering failed: {ex}").ConfigureAwait(false);
            }

            if (result == null || !result.Success || string.IsNullOrWhiteSpace(result.Base64PngData))
            {
                var failureDetail = result == null
                    ? "no render result was returned"
                    : !string.IsNullOrWhiteSpace(result.ErrorMessage)
                        ? result.ErrorMessage
                        : "no image data was returned";

                return await CreateFallbackAsync(
                    block,
                    context,
                    $"Mermaid rendering failed: {failureDetail}").ConfigureAwait(false);
            }

            var ns = context.OneNoteNamespace;
            var image = new XElement(ns + "Image",
                new XAttribute("format", "png"),
                new XAttribute("alt", "mermaid"),
                new XElement(ns + "Data", result.Base64PngData));

            return new[] { new XElement(ns + "OE", new XAttribute("alignment", "center"), image) };
        }

        private async Task<IReadOnlyList<XElement>> CreateFallbackAsync(
            Block block,
            IMarkdownConverterContext context,
            string logMessage)
        {
            RuntimeLog.Write(logMessage);

            var fallbackElements = await _codeBlockFallback
                .HandleAsync(block, context)
                .ConfigureAwait(false);
            var elements = new List<XElement>(fallbackElements.Count + 1)
            {
                CreateFailureNote(context)
            };
            elements.AddRange(fallbackElements);
            return elements;
        }

        private static XElement CreateFailureNote(IMarkdownConverterContext context)
        {
            var ns = context.OneNoteNamespace;
            var note = HtmlEscaper.Escape(Resources.GetString("Markdown_MermaidRenderFailureNote"));
            var html = $"<span style='font-size:9pt;color:#767676;font-style:italic'>{note}</span>";

            return new XElement(ns + "OE",
                new XAttribute("spaceBefore", "4.0"),
                new XAttribute("spaceAfter", "4.0"),
                new XElement(ns + "T", new XCData(html)));
        }

        private static string ExtractCode(CodeBlock codeBlock)
        {
            var lines = new List<string>();

            if (codeBlock.Lines.Lines == null)
            {
                return string.Empty;
            }

            foreach (var line in codeBlock.Lines.Lines)
            {
                if (line.Slice.Text == null)
                {
                    continue;
                }
                lines.Add(line.ToString());
            }

            // Remove trailing empty lines
            while (lines.Count > 0 && string.IsNullOrWhiteSpace(lines[lines.Count - 1]))
            {
                lines.RemoveAt(lines.Count - 1);
            }

            return string.Join("\n", lines);
        }
    }
}

