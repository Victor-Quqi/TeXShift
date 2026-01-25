using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Markdig.Syntax.Inlines;
using Markdig.Extensions.Mathematics;
using Markdig.Extensions.TaskLists;
using TeXShift.Core.Configuration;
using TeXShift.Core.Markdown.Abstractions;
using TeXShift.Core.Math;
using TeXShift.Core.Utils;

namespace TeXShift.Core.Markdown.Handlers.Inlines
{
    /// <summary>
    /// Converts Markdig inline elements to HTML for embedding in OneNote T elements.
    /// Handles emphasis (bold/italic/strikethrough), code, links, images, math, and line breaks.
    /// </summary>
    internal class InlineRenderer : IInlineRenderer
    {
        private readonly OneNoteStyleConfig _styleConfig;
        private readonly IMathService _mathService;

        public System.Func<string, string> EntityDecoder { get; set; }

        public InlineRenderer(OneNoteStyleConfig styleConfig, IMathService mathService)
        {
            _styleConfig = styleConfig;
            _mathService = mathService;
        }

        /// <summary>
        /// Converts a container of inline elements to an HTML string.
        /// </summary>
        public async Task<string> RenderAsync(ContainerInline container)
        {
            if (container == null) return string.Empty;
            return await RenderAsync((IEnumerable<Inline>)container).ConfigureAwait(false);
        }

        /// <summary>
        /// Converts a collection of inline elements to an HTML string.
        /// </summary>
        public async Task<string> RenderAsync(IEnumerable<Inline> inlines)
        {
            if (inlines == null) return string.Empty;
            var html = new StringBuilder();

            foreach (var inline in inlines)
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
                    await RenderEmphasisAsync(html, emphasis).ConfigureAwait(false);
                }
                else if (inline is CodeInline code)
                {
                    RenderCodeInline(html, code);
                }
                else if (inline is LinkInline link)
                {
                    await RenderLinkAsync(html, link).ConfigureAwait(false);
                }
                else if (inline is LineBreakInline)
                {
                    html.Append("\n");
                }
                else if (inline is MathInline mathInline)
                {
                    await RenderMathAsync(html, mathInline).ConfigureAwait(false);
                }
                else if (inline is ContainerInline nested)
                {
                    html.Append(await RenderAsync(nested).ConfigureAwait(false));
                }
            }
            return html.ToString();
        }

        private async Task RenderEmphasisAsync(StringBuilder html, EmphasisInline emphasis)
        {
            var content = await RenderAsync(emphasis).ConfigureAwait(false);
            if (emphasis.DelimiterChar == '*' || emphasis.DelimiterChar == '_')
            {
                if (emphasis.DelimiterCount == 2)
                    html.Append($"<span style='font-weight:bold'>{content}</span>");
                else if (emphasis.DelimiterCount == 1)
                    html.Append($"<span style='font-style:italic'>{content}</span>");
                else
                    html.Append(content);
            }
            else if (emphasis.DelimiterChar == '~' && emphasis.DelimiterCount == 2)
            {
                html.Append($"<span style='text-decoration:line-through'>{content}</span>");
            }
            else
            {
                html.Append(content);
            }
        }

        private void RenderCodeInline(StringBuilder html, CodeInline code)
        {
            var style = _styleConfig.GetInlineCodeStyle();
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

        private async Task RenderLinkAsync(StringBuilder html, LinkInline link)
        {
            var url = link.Url ?? "";

            // Handle images: inline images are downgraded to links
            if (link.IsImage)
            {
                var altText = await RenderAsync(link).ConfigureAwait(false);
                if (string.IsNullOrEmpty(altText))
                {
                    altText = "image";
                }
                // Downgrade to link with image icon prefix
                html.Append($"<a href=\"{HtmlEscaper.Escape(url)}\">[🖼️{altText}]</a>");
            }
            else
            {
                var content = await RenderAsync(link).ConfigureAwait(false);
                // If link text is empty, display the URL as the link text
                if (string.IsNullOrEmpty(content))
                {
                    content = HtmlEscaper.Escape(url);
                }
                html.Append($"<a href=\"{HtmlEscaper.Escape(url)}\">{content}</a>");
            }
        }

        private async Task RenderMathAsync(StringBuilder html, MathInline mathInline)
        {
            // Handle inline math ($...$) and display math ($$...$$)
            // DelimiterCount: 1 = $, 2 = $$
            var isDisplayMath = mathInline.DelimiterCount == 2;

            if (_mathService != null)
            {
                // Auto-initialize MathService if needed
                if (!_mathService.IsInitialized)
                {
                    try
                    {
                        await _mathService.InitializeAsync().ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Trace.WriteLine(ex);
                        // Initialization failed, show LaTeX source
                        var delim = isDisplayMath ? "$$" : "$";
                        html.Append($"[MathInit Error: {delim}{HtmlEscaper.Escape(mathInline.Content.ToString())}{delim}]");
                        return;
                    }
                }

                try
                {
                    var latex = mathInline.Content.ToString();
                    // Decode HTML entity placeholders before passing to MathJax
                    if (EntityDecoder != null)
                    {
                        latex = EntityDecoder(latex);
                    }
                    var mathml = await _mathService.LatexToMathMLAsync(latex, displayMode: isDisplayMath).ConfigureAwait(false);
                    var wrappedMathml = _mathService.WrapMathMLForOneNote(mathml);
                    html.Append(wrappedMathml);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Trace.WriteLine(ex);
                    // On conversion error, show the LaTeX source as plain text
                    var delim = isDisplayMath ? "$$" : "$";
                    html.Append($"[LaTeX: {delim}{HtmlEscaper.Escape(mathInline.Content.ToString())}{delim}]");
                }
            }
            else
            {
                // MathService not available, show LaTeX source
                var delim = isDisplayMath ? "$$" : "$";
                html.Append($"{delim}{HtmlEscaper.Escape(mathInline.Content.ToString())}{delim}");
            }
        }
    }
}
