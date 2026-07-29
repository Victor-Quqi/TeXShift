using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Markdig.Syntax.Inlines;
using Markdig.Extensions.Mathematics;
using Markdig.Extensions.TaskLists;
using TeXShift.Core.Configuration;
using TeXShift.Core.Markdown.Abstractions;
using TeXShift.Core.Markdown.Processing;
using TeXShift.Core.Math;
using TeXShift.Core.OneNote;
using TeXShift.Core.Utils;

namespace TeXShift.Core.Markdown.Handlers.Inlines
{
    /// <summary>
    /// Converts Markdig inline elements to HTML for embedding in OneNote T elements.
    /// Handles emphasis (bold/italic/strikethrough), code, links, images, math, and line breaks.
    /// </summary>
    internal class InlineRenderer : IInlineRenderer
    {
        private sealed class InlineRenderContext
        {
            public Stack<string> TextColors { get; } = new Stack<string>();

            public string CurrentTextColor => TextColors.Count > 0
                ? TextColors.Peek()
                : null;
        }

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
            return await RenderRootAsync((IEnumerable<Inline>)container).ConfigureAwait(false);
        }

        /// <summary>
        /// Converts a collection of inline elements to an HTML string.
        /// </summary>
        public async Task<string> RenderAsync(IEnumerable<Inline> inlines)
        {
            if (inlines == null) return string.Empty;
            return await RenderRootAsync(inlines).ConfigureAwait(false);
        }

        private async Task<string> RenderRootAsync(IEnumerable<Inline> inlines)
        {
            var context = new InlineRenderContext();
            string rendered = await RenderInlinesAsync(inlines, context).ConfigureAwait(false);
            if (context.TextColors.Count == 0)
            {
                return rendered;
            }

            var html = new StringBuilder(rendered);
            while (context.TextColors.Count > 0)
            {
                html.Append("</span>");
                context.TextColors.Pop();
            }
            return html.ToString();
        }

        private async Task<string> RenderInlinesAsync(
            IEnumerable<Inline> inlines,
            InlineRenderContext context)
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
                    await RenderEmphasisAsync(html, emphasis, context).ConfigureAwait(false);
                }
                else if (inline is CodeInline code)
                {
                    RenderCodeInline(html, code);
                }
                else if (inline is LinkInline link)
                {
                    await RenderLinkAsync(html, link, context).ConfigureAwait(false);
                }
                else if (inline is LineBreakInline)
                {
                    html.Append("\n");
                }
                else if (inline is HtmlInline htmlInline && HtmlLineBreakParser.IsLineBreakTag(htmlInline.Tag))
                {
                    html.Append('\n');
                }
                else if (inline is HtmlInline safeHtmlInline &&
                    TryGetSafeStyleTag(safeHtmlInline.Tag, context.TextColors, out var safeStyleTag))
                {
                    html.Append(safeStyleTag);
                }
                else if (inline is MathInline mathInline)
                {
                    await RenderMathAsync(html, mathInline).ConfigureAwait(false);
                }
                else if (inline is ContainerInline nested)
                {
                    html.Append(await RenderInlinesAsync(nested, context).ConfigureAwait(false));
                }
            }
            return html.ToString();
        }

        private static bool TryGetSafeStyleTag(
            string tag,
            Stack<string> textColors,
            out string safeTag)
        {
            safeTag = null;
            if (string.IsNullOrWhiteSpace(tag))
            {
                return false;
            }

            string normalized = tag.Trim();
            if (normalized.Length < 3 || normalized[0] != '<' || normalized[normalized.Length - 1] != '>')
            {
                return false;
            }

            string rawTag = normalized.Substring(1, normalized.Length - 2).Trim();
            bool isClosing = rawTag.StartsWith("/", StringComparison.Ordinal);
            if (isClosing)
            {
                rawTag = rawTag.Substring(1).TrimStart();
            }

            if (rawTag.Length == 0 || rawTag.EndsWith("/", StringComparison.Ordinal))
            {
                return false;
            }

            int nameLength = 0;
            while (nameLength < rawTag.Length && char.IsLetter(rawTag[nameLength]))
            {
                nameLength++;
            }

            if (nameLength == 0 ||
                (nameLength < rawTag.Length && !char.IsWhiteSpace(rawTag[nameLength])))
            {
                return false;
            }

            string attributes = rawTag.Substring(nameLength);
            if (isClosing && !string.IsNullOrWhiteSpace(attributes))
            {
                return false;
            }

            string tagName = rawTag.Substring(0, nameLength).ToLowerInvariant();
            switch (tagName)
            {
                case "strong":
                case "b":
                    safeTag = isClosing
                        ? "</span>"
                        : "<span style='" + OneNoteInlineStyles.BoldCss + "'>";
                    return true;
                case "em":
                case "i":
                    safeTag = isClosing
                        ? "</span>"
                        : "<span style='" + OneNoteInlineStyles.ItalicCss + "'>";
                    return true;
                case "span":
                    if (isClosing)
                    {
                        if (textColors == null || textColors.Count == 0)
                        {
                            return false;
                        }
                        textColors.Pop();
                        safeTag = "</span>";
                        return true;
                    }
                    if (CssColorParser.TryGetColorFromAttributes(attributes, out string color))
                    {
                        textColors?.Push(color);
                        safeTag = "<span style='color:" + color + "'>";
                        return true;
                    }
                    return false;
                case "mark":
                    safeTag = isClosing
                        ? "</span>"
                        : "<span style='" + OneNoteInlineStyles.HighlightCss + "'>";
                    return true;
                case "u":
                case "ins":
                    safeTag = isClosing
                        ? "</span>"
                        : "<span style='" + OneNoteInlineStyles.UnderlineCss + "'>";
                    return true;
                case "s":
                case "del":
                    safeTag = isClosing
                        ? "</span>"
                        : "<span style='" + OneNoteInlineStyles.StrikeCss + "'>";
                    return true;
                case "sup":
                case "sub":
                    safeTag = isClosing
                        ? "</" + tagName + ">"
                        : "<" + tagName + ">";
                    return true;
                default:
                    return false;
            }
        }

        private async Task RenderEmphasisAsync(
            StringBuilder html,
            EmphasisInline emphasis,
            InlineRenderContext context)
        {
            var content = await RenderInlinesAsync(emphasis, context).ConfigureAwait(false);
            if (emphasis.DelimiterChar == '*' || emphasis.DelimiterChar == '_')
            {
                if (emphasis.DelimiterCount == 2)
                    AppendStyledSpan(html, OneNoteInlineStyles.BoldCss, content);
                else if (emphasis.DelimiterCount == 1)
                    AppendStyledSpan(html, OneNoteInlineStyles.ItalicCss, content);
                else
                    html.Append(content);
            }
            else if (emphasis.DelimiterChar == '~' && emphasis.DelimiterCount == 2)
            {
                AppendStyledSpan(html, OneNoteInlineStyles.StrikeCss, content);
            }
            else if (emphasis.DelimiterChar == '=' && emphasis.DelimiterCount == 2)
            {
                AppendStyledSpan(html, OneNoteInlineStyles.HighlightCss, content);
            }
            else if (emphasis.DelimiterChar == '+' && emphasis.DelimiterCount == 2)
            {
                AppendStyledSpan(html, OneNoteInlineStyles.UnderlineCss, content);
            }
            else if (emphasis.DelimiterChar == '^' && emphasis.DelimiterCount == 1)
            {
                AppendInlineTag(html, "sup", content);
            }
            else if (emphasis.DelimiterChar == '~' && emphasis.DelimiterCount == 1)
            {
                AppendInlineTag(html, "sub", content);
            }
            else
            {
                html.Append(content);
            }
        }

        private static void AppendStyledSpan(StringBuilder html, string style, string content)
        {
            html.Append("<span style='").Append(style).Append("'>").Append(content).Append("</span>");
        }

        private static void AppendInlineTag(StringBuilder html, string tagName, string content)
        {
            html.Append('<').Append(tagName).Append('>')
                .Append(content)
                .Append("</").Append(tagName).Append('>');
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

        private async Task RenderLinkAsync(
            StringBuilder html,
            LinkInline link,
            InlineRenderContext context)
        {
            var url = link.Url ?? "";

            // Handle images: inline images are downgraded to links
            if (link.IsImage)
            {
                var altText = await RenderInlinesAsync(link, context).ConfigureAwait(false);
                if (string.IsNullOrEmpty(altText))
                {
                    altText = "image";
                }
                // Downgrade to link with image icon prefix
                html.Append($"<a href=\"{HtmlEscaper.Escape(url)}\">[🖼️{altText}]</a>");
            }
            else
            {
                string activeTextColor = context?.CurrentTextColor;
                var content = await RenderInlinesAsync(link, context).ConfigureAwait(false);
                // If link text is empty, display the URL as the link text
                if (string.IsNullOrEmpty(content))
                {
                    content = HtmlEscaper.Escape(url);
                }
                string colorStyle = string.IsNullOrEmpty(activeTextColor)
                    ? string.Empty
                    : " style='color:" + activeTextColor + "'";
                html.Append($"<a href=\"{HtmlEscaper.Escape(url)}\"{colorStyle}>{content}</a>");
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
                        // Initialization failed, show error message + LaTeX source for diagnosis
                        var delim = isDisplayMath ? "$$" : "$";
                        var errType = ex.GetType().Name;
                        var errMsg = ex.InnerException != null
                            ? $"{ex.Message} -> {ex.InnerException.GetType().Name}: {ex.InnerException.Message}"
                            : ex.Message;
                        html.Append($"[MathInit Error ({errType}: {HtmlEscaper.Escape(errMsg)}): {delim}{HtmlEscaper.Escape(mathInline.Content.ToString())}{delim}]");
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
