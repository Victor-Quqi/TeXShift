using System.Collections.Generic;
using System.Text;
using TeXShift.Core.Utils;

namespace TeXShift.Core.OneNoteToMarkdown.Inlines
{
    internal sealed partial class InlineParser
    {
        private void AppendText(string text, ParseState state)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            string decoded = OneNoteHtmlEntityDecoder.Decode(text);
            if (state.SuppressLeadingNewlineAfterBreak)
            {
                // OneNote serializes <br /> followed by a formatting newline in CDATA.
                decoded = RemoveOneLeadingNewline(decoded);
                state.SuppressLeadingNewlineAfterBreak = false;
                if (decoded.Length == 0)
                {
                    return;
                }
            }

            if (state.InlineCodeDepth > 0)
            {
                state.CodeBuffer.Append(decoded);
                return;
            }

            FlushPendingInlineCode(state);
            var desired = CurrentStyle(state.Stack);
            FlushPendingWhitespace(desired, ref state.PendingWhitespace, ref state.PendingWhitespaceStyle, ref state.EmittedStyle, state.Output);

            string tail = SplitTrailingWhitespaceIfStyled(decoded, desired, out string core);

            EnsureStyle(desired, ref state.EmittedStyle, state.Output);
            state.Output.Append(core);

            if (!string.IsNullOrEmpty(tail))
            {
                state.PendingWhitespace = tail;
                state.PendingWhitespaceStyle = desired;
            }
        }

        private void AppendLiteral(string literal, ParseState state)
        {
            if (string.IsNullOrEmpty(literal))
            {
                return;
            }

            if (state.InlineCodeDepth > 0)
            {
                state.CodeBuffer.Append(literal);
                return;
            }

            state.SuppressLeadingNewlineAfterBreak = false;
            FlushPendingInlineCode(state);
            var desired = CurrentStyle(state.Stack);
            FlushPendingWhitespace(desired, ref state.PendingWhitespace, ref state.PendingWhitespaceStyle, ref state.EmittedStyle, state.Output);
            EnsureStyle(desired, ref state.EmittedStyle, state.Output);
            state.Output.Append(literal);
        }

        private void AppendLineBreak(ParseState state)
        {
            if (state.InlineCodeDepth > 0)
            {
                state.CodeBuffer.Append("\n");
                return;
            }

            FlushPendingInlineCode(state);
            var desired = CurrentStyle(state.Stack);
            FlushPendingWhitespace(desired, ref state.PendingWhitespace, ref state.PendingWhitespaceStyle, ref state.EmittedStyle, state.Output);
            EnsureStyle(desired, ref state.EmittedStyle, state.Output);
            state.Output.Append("\n");
            state.SuppressLeadingNewlineAfterBreak = true;
        }

        private static string RemoveOneLeadingNewline(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            if (text.StartsWith("\r\n", System.StringComparison.Ordinal))
            {
                return text.Substring(2);
            }

            if (text[0] == '\r' || text[0] == '\n')
            {
                return text.Substring(1);
            }

            return text;
        }

        private static string NormalizeText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            // NBSP becomes a normal space; remove OneNote's zero-width sentinels.
            return text.Replace('\u00A0', ' ').Replace("\u200B", string.Empty);
        }
    }
}
