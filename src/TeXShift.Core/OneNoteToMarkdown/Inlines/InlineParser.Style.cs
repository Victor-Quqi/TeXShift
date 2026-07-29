using System;
using System.Collections.Generic;
using System.Text;
using TeXShift.Core.OneNote;
using TeXShift.Core.Utils;

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

        private static string GetSpanTextColor(string style)
        {
            return CssColorParser.TryGetColorFromStyle(style, out string color)
                ? color
                : null;
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

        private static InlineFormat CurrentFormat(Stack<Frame> stack)
        {
            if (stack == null || stack.Count == 0)
            {
                return InlineFormat.None;
            }

            InlineStyle style = InlineStyle.None;
            string textColor = null;
            foreach (var frame in stack)
            {
                style |= frame.Style;
                if (textColor == null && !string.IsNullOrWhiteSpace(frame.TextColor))
                {
                    textColor = frame.TextColor;
                }
            }
            return new InlineFormat(style, textColor);
        }

        private static InlineFormat MergeFormats(InlineFormat outer, InlineFormat inner)
        {
            return new InlineFormat(
                outer.Style | inner.Style,
                inner.TextColor ?? outer.TextColor);
        }

        private static void EnsureFormat(
            InlineFormat desired,
            ref InlineFormat emitted,
            StringBuilder output)
        {
            if (desired.Equals(emitted) || output == null)
            {
                return;
            }

            bool colorChanged = !string.Equals(
                desired.TextColor,
                emitted.TextColor,
                StringComparison.OrdinalIgnoreCase);
            InlineStyle emittedStyle = emitted.Style;

            if (colorChanged)
            {
                EnsureStyle(InlineStyle.None, ref emittedStyle, output);
                if (!string.IsNullOrEmpty(emitted.TextColor))
                {
                    output.Append("</span>");
                }
                if (!string.IsNullOrEmpty(desired.TextColor))
                {
                    output.Append("<span style=\"color:")
                        .Append(desired.TextColor)
                        .Append("\">");
                }
            }

            EnsureStyle(desired.Style, ref emittedStyle, output);
            emitted = desired;
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
            InlineFormat nextDesiredFormat,
            ref string pendingWhitespace,
            ref InlineFormat pendingWhitespaceFormat,
            ref InlineFormat emittedFormat,
            StringBuilder output)
        {
            if (string.IsNullOrEmpty(pendingWhitespace))
            {
                return;
            }

            bool sameColor = string.Equals(
                pendingWhitespaceFormat.TextColor,
                nextDesiredFormat.TextColor,
                StringComparison.OrdinalIgnoreCase);
            var whitespaceFormat = sameColor
                ? new InlineFormat(
                    pendingWhitespaceFormat.Style & nextDesiredFormat.Style,
                    pendingWhitespaceFormat.TextColor)
                : InlineFormat.None;
            EnsureFormat(whitespaceFormat, ref emittedFormat, output);
            output.Append(pendingWhitespace);
            pendingWhitespace = null;
            pendingWhitespaceFormat = InlineFormat.None;
        }

        private static string SplitTrailingWhitespaceIfStyled(
            string text,
            InlineFormat format,
            out string core)
        {
            core = text ?? string.Empty;
            if (string.IsNullOrEmpty(core) ||
                (format.Style == InlineStyle.None && string.IsNullOrEmpty(format.TextColor)))
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
            InlineFormat desiredFormat,
            InlineFormat emittedFormat,
            out string core,
            out InlineFormat leadingFormat)
        {
            core = text ?? string.Empty;
            leadingFormat = emittedFormat;
            if (string.IsNullOrEmpty(core) || desiredFormat.Equals(emittedFormat))
            {
                return null;
            }

            bool sameColor = string.Equals(
                desiredFormat.TextColor,
                emittedFormat.TextColor,
                StringComparison.OrdinalIgnoreCase);
            var emittedStyles = sameColor
                ? GetOrderedStyles(emittedFormat.Style)
                : new List<InlineStyle>();
            var desiredStyles = GetOrderedStyles(desiredFormat.Style);
            int common = 0;
            InlineStyle leadingStyle = InlineStyle.None;
            while (common < emittedStyles.Count &&
                   common < desiredStyles.Count &&
                   emittedStyles[common] == desiredStyles[common])
            {
                leadingStyle |= emittedStyles[common];
                common++;
            }

            leadingFormat = new InlineFormat(
                leadingStyle,
                sameColor ? desiredFormat.TextColor : null);

            // Closing inner styles keeps the whitespace inside an already-open outer style.
            bool opensColor = !sameColor && !string.IsNullOrEmpty(desiredFormat.TextColor);
            bool opensStyle = common < desiredStyles.Count;
            if (!opensColor && !opensStyle)
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
