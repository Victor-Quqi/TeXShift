using System.Collections.Generic;
using System.Text;
using TeXShift.Core.Utils;

namespace TeXShift.Core.OneNoteToMarkdown.Inlines
{
    internal sealed partial class InlineParser
    {
        private static void AppendText(string text, ParseState state)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            string decoded = OneNoteHtmlEntityDecoder.Decode(text);

            if (state.InlineCodeDepth > 0)
            {
                state.CodeBuffer.Append(decoded);
                return;
            }

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

        private static void AppendLiteral(string literal, ParseState state)
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

            var desired = CurrentStyle(state.Stack);
            FlushPendingWhitespace(desired, ref state.PendingWhitespace, ref state.PendingWhitespaceStyle, ref state.EmittedStyle, state.Output);
            EnsureStyle(desired, ref state.EmittedStyle, state.Output);
            state.Output.Append(literal);
        }

        private static void AppendLineBreak(ParseState state)
        {
            if (state.InlineCodeDepth > 0)
            {
                state.CodeBuffer.Append("\n");
                return;
            }

            var desired = CurrentStyle(state.Stack);
            FlushPendingWhitespace(desired, ref state.PendingWhitespace, ref state.PendingWhitespaceStyle, ref state.EmittedStyle, state.Output);
            EnsureStyle(desired, ref state.EmittedStyle, state.Output);
            state.Output.Append("\n");
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
