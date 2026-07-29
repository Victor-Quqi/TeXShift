using System;
using System.Collections.Generic;
using System.Text;

namespace TeXShift.Core.OneNoteToMarkdown.Inlines
{
    internal sealed partial class InlineParser
    {
        private bool IsInlineCodeSpan(string style)
        {
            if (string.IsNullOrEmpty(style))
            {
                return false;
            }

            var bg = TryGetBackgroundColor(style);
            if (string.IsNullOrEmpty(bg) || string.IsNullOrWhiteSpace(_inlineCodeBackgroundColor))
            {
                return false;
            }

            return string.Equals(bg, _inlineCodeBackgroundColor, StringComparison.OrdinalIgnoreCase);
        }

        private static string TryGetBackgroundColor(string style)
        {
            if (string.IsNullOrEmpty(style))
            {
                return null;
            }

            var match = BackgroundColorRegex.Match(style);
            return match.Success ? match.Groups["v"].Value : null;
        }

        private bool HasInlineCodeMonospaceFont(string style)
        {
            if (string.IsNullOrEmpty(style) || string.IsNullOrWhiteSpace(_inlineCodeFontFamily))
            {
                return false;
            }

            return style.IndexOf(_inlineCodeFontFamily, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool HasLeadingNbspPadding(string html, int startIndex, int paddingCount)
        {
            if (string.IsNullOrEmpty(html) || paddingCount <= 0)
            {
                return true;
            }

            int i = startIndex;
            for (int n = 0; n < paddingCount; n++)
            {
                if (i >= html.Length)
                {
                    return false;
                }

                // TeXShift inline code padding uses &nbsp; by default; OneNote may rewrite it
                // into other non-breaking-space encodings. Stay strict to avoid false positives.
                if (StartsWithIgnoreCase(html, i, "&nbsp;") ||
                    StartsWithIgnoreCase(html, i, "&#160;") ||
                    StartsWithIgnoreCase(html, i, "&#xA0;") ||
                    StartsWithIgnoreCase(html, i, "&#xa0;"))
                {
                    i += 6;
                    continue;
                }

                if (html[i] == '\u00A0')
                {
                    i += 1;
                    continue;
                }

                return false;
            }

            return true;
        }

        private static bool StartsWithIgnoreCase(string text, int startIndex, string token)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(token))
            {
                return false;
            }

            if (startIndex < 0 || startIndex + token.Length > text.Length)
            {
                return false;
            }

            return string.Compare(text, startIndex, token, 0, token.Length, StringComparison.OrdinalIgnoreCase) == 0;
        }

        private void QueueInlineCode(
            ParseState state,
            InlineFormat format,
            bool inlineCodeHasMonospace)
        {
            if (state == null)
            {
                return;
            }

            // OneNote can split padding, content, and trailing padding into adjacent spans.
            // Delay emission so those spans can be reconstructed as one code run.
            state.HasPendingInlineCode = true;
            state.PendingInlineCodeFormat = format;
            state.PendingInlineCodeHasMonospace = inlineCodeHasMonospace;
        }

        private void FlushPendingInlineCode(ParseState state)
        {
            if (state == null || !state.HasPendingInlineCode)
            {
                return;
            }

            var format = state.PendingInlineCodeFormat;
            bool hasMonospace = state.PendingInlineCodeHasMonospace;
            state.HasPendingInlineCode = false;
            state.PendingInlineCodeFormat = InlineFormat.None;
            state.PendingInlineCodeHasMonospace = false;
            EmitInlineCode(state, format, hasMonospace);
        }

        private void EmitInlineCode(ParseState state, InlineFormat desired, bool inlineCodeHasMonospace)
        {
            if (state == null)
            {
                return;
            }

            FlushPendingWhitespace(desired, ref state.PendingWhitespace, ref state.PendingWhitespaceFormat, ref state.EmittedFormat, state.Output);
            EnsureFormat(desired, ref state.EmittedFormat, state.Output);

            // Strict mode (TeXShift-only): background match is already strong enough.
            // Fuzzy mode: require either monospace font or TeXShift-style padding,
            // to avoid misclassifying arbitrary highlighted text as inline code.
            string raw = NormalizeText(state.CodeBuffer.ToString());
            bool shouldEmitInlineCode = true;
            if (_tryRecognizeNonTeXShiftFormats)
            {
                shouldEmitInlineCode =
                    inlineCodeHasMonospace ||
                    _inlineCodePaddingCount == 0 ||
                    HasPadding(raw, _inlineCodePaddingCount);
            }

            if (!shouldEmitInlineCode)
            {
                state.Output.Append(raw);
                state.CodeBuffer.Clear();
                return;
            }

            string codeText = TrimPadding(raw, _inlineCodePaddingCount);
            codeText = codeText.Replace("\r", string.Empty).Replace("\n", " ");
            state.Output.Append(WrapInlineCode(codeText));
            state.CodeBuffer.Clear();
        }

        private static bool HasPadding(string text, int paddingCount)
        {
            if (string.IsNullOrEmpty(text) || paddingCount <= 0)
            {
                return true;
            }

            int leading = 0;
            while (leading < text.Length && IsPaddingChar(text[leading]))
            {
                leading++;
            }

            int trailing = 0;
            int i = text.Length - 1;
            while (i >= 0 && IsPaddingChar(text[i]))
            {
                trailing++;
                i--;
            }

            return leading >= paddingCount && trailing >= paddingCount;
        }

        private static string TrimPadding(string text, int paddingCount)
        {
            if (string.IsNullOrEmpty(text) || paddingCount <= 0)
            {
                return text ?? string.Empty;
            }

            int start = 0;
            int removed = 0;
            while (removed < paddingCount && start < text.Length && IsPaddingChar(text[start]))
            {
                start++;
                removed++;
            }

            int end = text.Length;
            removed = 0;
            while (removed < paddingCount && end > start && IsPaddingChar(text[end - 1]))
            {
                end--;
                removed++;
            }

            return text.Substring(start, end - start);
        }

        private static bool IsPaddingChar(char c)
        {
            return c == ' ' || c == '\u00A0';
        }

        private static string WrapInlineCode(string code)
        {
            if (code == null)
            {
                code = string.Empty;
            }

            int maxBackticks = 0;
            int current = 0;
            foreach (char ch in code)
            {
                if (ch == '`')
                {
                    current++;
                    if (current > maxBackticks)
                    {
                        maxBackticks = current;
                    }
                }
                else
                {
                    current = 0;
                }
            }

            string fence = new string('`', System.Math.Max(1, maxBackticks + 1));
            if (code.StartsWith(" ", StringComparison.Ordinal) ||
                code.EndsWith(" ", StringComparison.Ordinal) ||
                code.StartsWith("`", StringComparison.Ordinal) ||
                code.EndsWith("`", StringComparison.Ordinal))
            {
                return fence + " " + code + " " + fence;
            }

            return fence + code + fence;
        }
    }
}
