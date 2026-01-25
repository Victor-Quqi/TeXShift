using System;
using System.Collections.Generic;
using System.Text;

namespace TeXShift.Core.OneNoteToMarkdown.Inlines
{
    internal sealed partial class InlineParser
    {
        private static InlineStyle GetSpanStyleFlags(string style, bool ignoreBold)
        {
            if (string.IsNullOrEmpty(style))
            {
                return InlineStyle.None;
            }

            string s = NormalizeStyleForContains(style);

            InlineStyle flags = InlineStyle.None;
            if (s.Contains("text-decoration:line-through"))
            {
                flags |= InlineStyle.Strike;
            }
            if (!ignoreBold && s.Contains("font-weight:bold"))
            {
                flags |= InlineStyle.Bold;
            }
            if (s.Contains("font-style:italic"))
            {
                flags |= InlineStyle.Italic;
            }

            return flags;
        }

        private static string NormalizeStyleForContains(string style)
        {
            if (string.IsNullOrEmpty(style))
            {
                return string.Empty;
            }

            // OneNote may insert newlines into the style attribute value.
            // Remove all whitespace to make substring checks stable.
            var sb = new StringBuilder(style.Length);
            foreach (char ch in style)
            {
                if (!char.IsWhiteSpace(ch))
                {
                    sb.Append(char.ToLowerInvariant(ch));
                }
            }

            return sb.ToString();
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

            // Close markers (inner-most first) in reverse order of opening: Italic, Bold, Strike.
            var toClose = emitted & ~desired;
            if ((toClose & InlineStyle.Italic) != 0)
            {
                output.Append("*");
                emitted &= ~InlineStyle.Italic;
            }
            if ((toClose & InlineStyle.Bold) != 0)
            {
                output.Append("**");
                emitted &= ~InlineStyle.Bold;
            }
            if ((toClose & InlineStyle.Strike) != 0)
            {
                output.Append("~~");
                emitted &= ~InlineStyle.Strike;
            }

            // Open markers (outer-most first): Strike, Bold, Italic.
            var toOpen = desired & ~emitted;
            if ((toOpen & InlineStyle.Strike) != 0)
            {
                output.Append("~~");
                emitted |= InlineStyle.Strike;
            }
            if ((toOpen & InlineStyle.Bold) != 0)
            {
                output.Append("**");
                emitted |= InlineStyle.Bold;
            }
            if ((toOpen & InlineStyle.Italic) != 0)
            {
                output.Append("*");
                emitted |= InlineStyle.Italic;
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
    }
}

