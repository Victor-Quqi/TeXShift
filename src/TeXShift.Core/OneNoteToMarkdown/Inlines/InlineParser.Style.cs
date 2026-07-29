using System;
using System.Collections.Generic;
using System.Text;
using TeXShift.Core.OneNote;

namespace TeXShift.Core.OneNoteToMarkdown.Inlines
{
    internal sealed partial class InlineParser
    {
        // OneNote flattens nested spans. Put styles that commonly cover wider ranges first
        // so transitions remain valid Markdown after the original nesting information is lost.
        private static readonly InlineStyle[] StyleNestingOrder =
        {
            InlineStyle.Highlight,
            InlineStyle.Underline,
            InlineStyle.Strike,
            InlineStyle.Bold,
            InlineStyle.Italic,
            InlineStyle.Superscript,
            InlineStyle.Subscript,
        };

        private InlineStyle GetSpanStyleFlags(
            string style,
            bool ignoreBold,
            bool includeBackgroundHighlight = true)
        {
            if (string.IsNullOrEmpty(style))
            {
                return InlineStyle.None;
            }

            InlineStyle flags = InlineStyle.None;
            string textDecoration = GetCssPropertyValue(style, "text-decoration");
            string textDecorationLine = GetCssPropertyValue(style, "text-decoration-line");
            if (ContainsCssToken(textDecoration, "line-through") ||
                ContainsCssToken(textDecorationLine, "line-through"))
            {
                flags |= InlineStyle.Strike;
            }
            if (ContainsCssToken(textDecoration, "underline") ||
                ContainsCssToken(textDecorationLine, "underline"))
            {
                flags |= InlineStyle.Underline;
            }

            string fontWeight = GetCssPropertyValue(style, "font-weight");
            if (!ignoreBold && string.Equals(fontWeight, "bold", StringComparison.OrdinalIgnoreCase))
            {
                flags |= InlineStyle.Bold;
            }

            string fontStyle = GetCssPropertyValue(style, "font-style");
            if (string.Equals(fontStyle, "italic", StringComparison.OrdinalIgnoreCase))
            {
                flags |= InlineStyle.Italic;
            }

            string verticalAlign = GetCssPropertyValue(style, "vertical-align");
            if (string.Equals(verticalAlign, "super", StringComparison.OrdinalIgnoreCase))
            {
                flags |= InlineStyle.Superscript;
            }
            else if (string.Equals(verticalAlign, "sub", StringComparison.OrdinalIgnoreCase))
            {
                flags |= InlineStyle.Subscript;
            }

            if (includeBackgroundHighlight)
            {
                string background = GetCssPropertyValue(style, "background-color")
                    ?? GetCssPropertyValue(style, "background");
                if (OneNoteInlineStyles.IsCanonicalHighlightColor(background) ||
                    (_tryRecognizeNonTeXShiftFormats && OneNoteInlineStyles.IsVisibleBackgroundColor(background)))
                {
                    flags |= InlineStyle.Highlight;
                }
            }

            return flags;
        }

        private static string GetCssPropertyValue(string style, string propertyName)
        {
            if (string.IsNullOrWhiteSpace(style) || string.IsNullOrWhiteSpace(propertyName))
            {
                return null;
            }

            string result = null;
            foreach (var declaration in style.Split(';'))
            {
                int colon = declaration.IndexOf(':');
                if (colon <= 0)
                {
                    continue;
                }

                string name = declaration.Substring(0, colon).Trim();
                if (!string.Equals(name, propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                result = declaration.Substring(colon + 1).Trim();
                const string important = "!important";
                if (result.EndsWith(important, StringComparison.OrdinalIgnoreCase))
                {
                    result = result.Substring(0, result.Length - important.Length).TrimEnd();
                }
            }

            return result;
        }

        private static bool ContainsCssToken(string value, string token)
        {
            if (string.IsNullOrWhiteSpace(value) || string.IsNullOrWhiteSpace(token))
            {
                return false;
            }

            var tokens = value.Split(new[] { ' ', '\t', '\r', '\n', ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var candidate in tokens)
            {
                if (string.Equals(candidate, token, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static InlineStyle CurrentStyle(Stack<Frame> stack)
        {
            if (stack == null || stack.Count == 0)
            {
                return InlineStyle.None;
            }

            InlineStyle style = InlineStyle.None;
            foreach (var frame in stack)
            {
                style |= frame.Style;
            }
            return style;
        }

        private static void EnsureStyle(InlineStyle desired, ref InlineStyle emitted, StringBuilder output)
        {
            if (output == null)
            {
                return;
            }

            if (desired == emitted)
            {
                return;
            }

            var emittedStyles = GetOrderedStyles(emitted);
            var desiredStyles = GetOrderedStyles(desired);
            int common = 0;
            while (common < emittedStyles.Count &&
                   common < desiredStyles.Count &&
                   emittedStyles[common] == desiredStyles[common])
            {
                common++;
            }

            for (int i = emittedStyles.Count - 1; i >= common; i--)
            {
                output.Append(GetClosingDelimiter(emittedStyles[i], emitted));
            }

            for (int i = common; i < desiredStyles.Count; i++)
            {
                output.Append(GetOpeningDelimiter(desiredStyles[i], desired));
            }

            emitted = desired;
        }

        private static List<InlineStyle> GetOrderedStyles(InlineStyle styles)
        {
            var result = new List<InlineStyle>(StyleNestingOrder.Length);
            foreach (var style in StyleNestingOrder)
            {
                if ((styles & style) != 0)
                {
                    result.Add(style);
                }
            }

            return result;
        }

        private static string GetOpeningDelimiter(InlineStyle style, InlineStyle fullStyle)
        {
            if (style == InlineStyle.Subscript && (fullStyle & InlineStyle.Strike) != 0)
            {
                // A triple tilde run is not parsed by Markdig as nested strike + subscript.
                return "<sub>";
            }

            return GetMarkdownDelimiter(style);
        }

        private static string GetClosingDelimiter(InlineStyle style, InlineStyle fullStyle)
        {
            if (style == InlineStyle.Subscript && (fullStyle & InlineStyle.Strike) != 0)
            {
                return "</sub>";
            }

            return GetMarkdownDelimiter(style);
        }

        private static string GetMarkdownDelimiter(InlineStyle style)
        {
            switch (style)
            {
                case InlineStyle.Strike:
                    return "~~";
                case InlineStyle.Bold:
                    return "**";
                case InlineStyle.Italic:
                    return "*";
                case InlineStyle.Highlight:
                    return "==";
                case InlineStyle.Underline:
                    return "++";
                case InlineStyle.Superscript:
                    return "^";
                case InlineStyle.Subscript:
                    return "~";
                default:
                    return string.Empty;
            }
        }

        private static void FlushPendingWhitespace(
            InlineStyle nextDesiredStyle,
            ref string pendingWhitespace,
            ref InlineStyle pendingWhitespaceStyle,
            ref InlineStyle emittedStyle,
            StringBuilder output)
        {
            if (string.IsNullOrEmpty(pendingWhitespace))
            {
                return;
            }

            var wsStyle = pendingWhitespaceStyle & nextDesiredStyle;
            EnsureStyle(wsStyle, ref emittedStyle, output);
            output.Append(pendingWhitespace);
            pendingWhitespace = null;
            pendingWhitespaceStyle = InlineStyle.None;
        }

        private static string SplitTrailingWhitespaceIfStyled(string text, InlineStyle style, out string core)
        {
            core = text ?? string.Empty;
            if (string.IsNullOrEmpty(core) || style == InlineStyle.None)
            {
                return null;
            }

            int split = core.Length;
            while (split > 0)
            {
                char c = core[split - 1];
                if (c == '\r' || c == '\n')
                {
                    break;
                }

                if (!char.IsWhiteSpace(c))
                {
                    break;
                }

                split--;
            }

            if (split == core.Length)
            {
                return null;
            }

            string tail = core.Substring(split);
            core = core.Substring(0, split);
            return tail;
        }

        private static string SplitLeadingWhitespaceOnStyleChange(
            string text,
            InlineStyle desiredStyle,
            InlineStyle emittedStyle,
            out string core,
            out InlineStyle leadingStyle)
        {
            core = text ?? string.Empty;
            leadingStyle = emittedStyle;
            if (string.IsNullOrEmpty(core) || desiredStyle == emittedStyle)
            {
                return null;
            }

            var emittedStyles = GetOrderedStyles(emittedStyle);
            var desiredStyles = GetOrderedStyles(desiredStyle);
            int common = 0;
            leadingStyle = InlineStyle.None;
            while (common < emittedStyles.Count &&
                   common < desiredStyles.Count &&
                   emittedStyles[common] == desiredStyles[common])
            {
                leadingStyle |= emittedStyles[common];
                common++;
            }

            // Closing inner styles keeps the whitespace inside an already-open outer style.
            if (common == desiredStyles.Count)
            {
                return null;
            }

            int split = 0;
            while (split < core.Length)
            {
                char c = core[split];
                if (c == '\r' || c == '\n' || !char.IsWhiteSpace(c))
                {
                    break;
                }

                split++;
            }

            if (split == 0)
            {
                return null;
            }

            string leading = core.Substring(0, split);
            core = core.Substring(split);
            return leading;
        }
    }
}
