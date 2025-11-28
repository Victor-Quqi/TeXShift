using System.Collections.Generic;
using System.Text;
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

        public InlineRenderer(OneNoteStyleConfig styleConfig, IMathService mathService)
        {
            _styleConfig = styleConfig;
            _mathService = mathService;
        }

        /// <summary>
        /// Converts a container of inline elements to an HTML string.
        /// </summary>
        public string Render(ContainerInline container)
        {
            if (container == null) return string.Empty;
            return Render((IEnumerable<Inline>)container);
        }

        /// <summary>
        /// Converts a collection of inline elements to an HTML string.
        /// </summary>
        public string Render(IEnumerable<Inline> inlines)
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
                    RenderEmphasis(html, emphasis);
                }
                else if (inline is CodeInline code)
                {
                    RenderCodeInline(html, code);
                }
                else if (inline is LinkInline link)
                {
                    RenderLink(html, link);
                }
                else if (inline is LineBreakInline)
                {
                    html.Append("\n");
                }
                else if (inline is MathInline mathInline)
                {
                    RenderMath(html, mathInline);
                }
                else if (inline is ContainerInline nested)
                {
                    html.Append(Render(nested));
                }
            }
            return html.ToString();
        }

        private void RenderEmphasis(StringBuilder html, EmphasisInline emphasis)
        {
            var content = Render(emphasis);
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

        private void RenderLink(StringBuilder html, LinkInline link)
        {
            var url = link.Url ?? "";

            // Handle images: inline images are downgraded to links
            if (link.IsImage)
            {
                var altText = Render(link);
                if (string.IsNullOrEmpty(altText))
                {
                    altText = "image";
                }
                // Downgrade to link with image icon prefix
                html.Append($"<a href=\"{HtmlEscaper.Escape(url)}\">[🖼️{altText}]</a>");
            }
            else
            {
                var content = Render(link);
                // If link text is empty, display the URL as the link text
                if (string.IsNullOrEmpty(content))
                {
                    content = HtmlEscaper.Escape(url);
                }
                html.Append($"<a href=\"{HtmlEscaper.Escape(url)}\">{content}</a>");
            }
        }

        private void RenderMath(StringBuilder html, MathInline mathInline)
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
                        _mathService.InitializeAsync().GetAwaiter().GetResult();
                    }
                    catch
                    {
                        // Initialization failed, show LaTeX source
                        var delim = isDisplayMath ? "$$" : "$";
                        html.Append($"[MathInit Error: {delim}{HtmlEscaper.Escape(mathInline.Content.ToString())}{delim}]");
                        return;
                    }
                }

                try
                {
                    var latex = mathInline.Content.ToString();
                    var mathml = _mathService.LatexToMathMLAsync(latex, displayMode: isDisplayMath).GetAwaiter().GetResult();
                    var wrappedMathml = _mathService.WrapMathMLForOneNote(mathml);
                    html.Append(wrappedMathml);
                }
                catch
                {
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
