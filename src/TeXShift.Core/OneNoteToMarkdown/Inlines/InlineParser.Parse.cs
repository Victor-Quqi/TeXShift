using System;
using System.Collections.Generic;
using System.Text;
using TeXShift.Core.Localization;
using TeXShift.Core.Utils;

namespace TeXShift.Core.OneNoteToMarkdown.Inlines
{
    internal sealed partial class InlineParser
    {
        private sealed class ParseState
        {
            public readonly Stack<Frame> Stack;
            public readonly StringBuilder Output;
            public readonly StringBuilder CodeBuffer;
            public InlineStyle EmittedStyle;
            public int InlineCodeDepth;
            public bool HasPendingInlineCode;
            public InlineStyle PendingInlineCodeStyle;
            public bool PendingInlineCodeHasMonospace;
            public bool SuppressLeadingNewlineAfterBreak;
            public string PendingWhitespace;
            public InlineStyle PendingWhitespaceStyle;

            public ParseState(int outputCapacity)
            {
                Stack = new Stack<Frame>();
                Output = new StringBuilder(outputCapacity);
                CodeBuffer = new StringBuilder();
                EmittedStyle = InlineStyle.None;
                InlineCodeDepth = 0;
                HasPendingInlineCode = false;
                PendingInlineCodeStyle = InlineStyle.None;
                PendingInlineCodeHasMonospace = false;
                SuppressLeadingNewlineAfterBreak = false;
                PendingWhitespace = null;
                PendingWhitespaceStyle = InlineStyle.None;
            }
        }

        public string Parse(string html, InlineParseMode mode = InlineParseMode.Default)
        {
            if (string.IsNullOrEmpty(html))
            {
                return string.Empty;
            }

            bool ignoreBold = mode == InlineParseMode.Heading;

            var state = new ParseState(html.Length);

            int i = 0;
            while (i < html.Length)
            {
                int lt = html.IndexOf('<', i);
                if (lt < 0)
                {
                    AppendText(html.Substring(i), state);
                    break;
                }

                if (lt > i)
                {
                    AppendText(html.Substring(i, lt - i), state);
                }

                if (TryConsumeHtmlComment(html, lt, state, out i, out bool stop))
                {
                    if (stop)
                    {
                        break;
                    }

                    continue;
                }

                if (!TryConsumeHtmlTag(html, lt, ignoreBold, state, out i))
                {
                    break;
                }
            }

            FinalizeParse(state);
            return NormalizeText(state.Output.ToString());
        }

        private bool TryConsumeHtmlComment(
            string html,
            int lt,
            ParseState state,
            out int nextIndex,
            out bool stop)
        {
            nextIndex = lt;
            stop = false;

            // HTML comments: <!-- ... -->
            if (lt + 3 >= html.Length || html[lt + 1] != '!' || html[lt + 2] != '-' || html[lt + 3] != '-')
            {
                return false;
            }

            int commentEnd = html.IndexOf("-->", lt + 4, StringComparison.Ordinal);
            if (commentEnd < 0)
            {
                AppendText(html.Substring(lt), state);
                nextIndex = html.Length;
                stop = true;
                return true;
            }

            string commentBody = html.Substring(lt + 4, commentEnd - (lt + 4));
            nextIndex = commentEnd + 3;

            // OneNote MathML is stored inside conditional comments: <!--[if mathML]>...<![endif]-->
            // Without TeXShift meta we cannot reliably recover LaTeX, so emit a placeholder.
            if (commentBody.IndexOf("mathml", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                AppendLiteral(Resources.GetString("Reverse_MathSourceMissing"), state);
            }

            return true;
        }

        private bool TryConsumeHtmlTag(
            string html,
            int lt,
            bool ignoreBold,
            ParseState state,
            out int nextIndex)
        {
            nextIndex = lt;

            int gt = html.IndexOf('>', lt + 1);
            if (gt < 0)
            {
                AppendText(html.Substring(lt), state);
                nextIndex = html.Length;
                return false;
            }

            string rawTag = html.Substring(lt + 1, gt - lt - 1).Trim();
            nextIndex = gt + 1;

            if (rawTag.Length == 0)
            {
                return true;
            }

            bool isClosing = rawTag[0] == '/';

            string tagName = GetTagName(isClosing ? rawTag.Substring(1) : rawTag);
            if (string.IsNullOrEmpty(tagName))
            {
                return true;
            }

            tagName = tagName.ToLowerInvariant();

            // Preserve unknown tags as literal text (they are often user-authored "<...>" sequences).
            if (!IsSupportedInlineTag(tagName))
            {
                AppendLiteral("<" + rawTag + ">", state);
                return true;
            }

            if (isClosing)
            {
                if (state.InlineCodeDepth == 0)
                {
                    FlushPendingInlineCode(state);
                }
                CloseUntil(tagName, state);
                return true;
            }

            if (tagName == "br")
            {
                AppendLineBreak(state);
                return true;
            }

            if (tagName == "span")
            {
                string style = GetAttributeValue(rawTag, HtmlRegexes.StyleAttr) ?? string.Empty;
                if (IsInlineCodeSpan(style))
                {
                    bool hasMonospace = HasInlineCodeMonospaceFont(style);
                    bool treatAsInlineCode =
                        !_tryRecognizeNonTeXShiftFormats ||
                        hasMonospace ||
                        _inlineCodePaddingCount == 0 ||
                        HasLeadingNbspPadding(html, nextIndex, _inlineCodePaddingCount);

                    if (!treatAsInlineCode)
                    {
                        state.Stack.Push(new Frame("span", GetSpanStyleFlags(style, ignoreBold)));
                        return true;
                    }

                    var spanStyle = GetSpanStyleFlags(style, ignoreBold, includeBackgroundHighlight: false);
                    var desiredCodeStyle = CurrentStyle(state.Stack) | spanStyle;
                    bool continuesPendingCode = false;
                    if (state.HasPendingInlineCode)
                    {
                        if (state.PendingInlineCodeStyle == desiredCodeStyle)
                        {
                            continuesPendingCode = true;
                            hasMonospace |= state.PendingInlineCodeHasMonospace;
                            state.HasPendingInlineCode = false;
                            state.PendingInlineCodeStyle = InlineStyle.None;
                            state.PendingInlineCodeHasMonospace = false;
                        }
                        else
                        {
                            FlushPendingInlineCode(state);
                        }
                    }

                    state.InlineCodeDepth++;
                    if (state.InlineCodeDepth == 1 && !continuesPendingCode)
                    {
                        state.CodeBuffer.Clear();
                    }

                    // Inline code suppresses other styles; capture literal text until the span closes.
                    state.Stack.Push(new Frame("span", spanStyle, isInlineCode: true, inlineCodeHasMonospace: hasMonospace));
                    return true;
                }

                FlushPendingInlineCode(state);
                state.Stack.Push(new Frame("span", GetSpanStyleFlags(style, ignoreBold)));
                return true;
            }

            FlushPendingInlineCode(state);
            var tagStyle = GetTagStyle(tagName, ignoreBold);
            if (tagName != "a")
            {
                state.Stack.Push(new Frame(tagName, tagStyle));
                return true;
            }

            if (tagName == "a")
            {
                string href = GetAttributeValue(rawTag, HrefAttrRegex);
                href = OneNoteHtmlEntityDecoder.Decode(href ?? string.Empty);

                if (state.InlineCodeDepth == 0)
                {
                    var desired = CurrentStyle(state.Stack);
                    FlushPendingWhitespace(desired, ref state.PendingWhitespace, ref state.PendingWhitespaceStyle, ref state.EmittedStyle, state.Output);
                    EnsureStyle(desired, ref state.EmittedStyle, state.Output);
                    state.Output.Append('[');
                    state.Stack.Push(new Frame("a", InlineStyle.None, href: href, contentStartIndex: state.Output.Length));
                }
                return true;
            }

            return true;
        }

        private static bool IsSupportedInlineTag(string tagName)
        {
            switch (tagName)
            {
                case "span":
                case "a":
                case "br":
                case "b":
                case "strong":
                case "i":
                case "em":
                case "u":
                case "ins":
                case "s":
                case "strike":
                case "del":
                case "sup":
                case "sub":
                case "mark":
                    return true;
                default:
                    return false;
            }
        }

        private static InlineStyle GetTagStyle(string tagName, bool ignoreBold)
        {
            switch (tagName)
            {
                case "b":
                case "strong":
                    return ignoreBold ? InlineStyle.None : InlineStyle.Bold;
                case "i":
                case "em":
                    return InlineStyle.Italic;
                case "u":
                case "ins":
                    return InlineStyle.Underline;
                case "s":
                case "strike":
                case "del":
                    return InlineStyle.Strike;
                case "sup":
                    return InlineStyle.Superscript;
                case "sub":
                    return InlineStyle.Subscript;
                case "mark":
                    return InlineStyle.Highlight;
                default:
                    return InlineStyle.None;
            }
        }

        private void FinalizeParse(ParseState state)
        {
            // Flush unterminated inline code.
            if (state.InlineCodeDepth > 0)
            {
                state.InlineCodeDepth = 0;
                EmitInlineCode(state, CurrentStyle(state.Stack), inlineCodeHasMonospace: false);
            }
            else
            {
                FlushPendingInlineCode(state);
            }

            // Close any unclosed links.
            while (state.Stack.Count > 0)
            {
                var frame = state.Stack.Pop();
                if (frame.IsInlineCode)
                {
                    continue;
                }

                if (string.Equals(frame.TagName, "a", StringComparison.OrdinalIgnoreCase))
                {
                    FlushPendingWhitespace(CurrentStyle(state.Stack), ref state.PendingWhitespace, ref state.PendingWhitespaceStyle, ref state.EmittedStyle, state.Output);
                    EnsureStyle(CurrentStyle(state.Stack), ref state.EmittedStyle, state.Output);
                    if (!TryReplaceImageLink(state.Output, frame))
                    {
                        state.Output.Append("](").Append(frame.Href ?? string.Empty).Append(')');
                    }
                }
            }

            FlushPendingWhitespace(InlineStyle.None, ref state.PendingWhitespace, ref state.PendingWhitespaceStyle, ref state.EmittedStyle, state.Output);
            EnsureStyle(InlineStyle.None, ref state.EmittedStyle, state.Output);
        }

        private void CloseUntil(
            string tagName,
            ParseState state)
        {
            if (state == null || state.Stack.Count == 0)
            {
                return;
            }

            string needle = tagName.ToLowerInvariant();
            while (state.Stack.Count > 0)
            {
                var frame = state.Stack.Pop();
                if (frame.IsInlineCode)
                {
                    state.InlineCodeDepth = System.Math.Max(0, state.InlineCodeDepth - 1);
                    if (state.InlineCodeDepth == 0)
                    {
                        QueueInlineCode(
                            state,
                            CurrentStyle(state.Stack) | frame.Style,
                            frame.InlineCodeHasMonospace);
                    }
                }

                if (string.Equals(frame.TagName, needle, StringComparison.OrdinalIgnoreCase))
                {
                    if (string.Equals(needle, "a", StringComparison.OrdinalIgnoreCase) && state.InlineCodeDepth == 0)
                    {
                        FlushPendingWhitespace(CurrentStyle(state.Stack), ref state.PendingWhitespace, ref state.PendingWhitespaceStyle, ref state.EmittedStyle, state.Output);
                        EnsureStyle(CurrentStyle(state.Stack), ref state.EmittedStyle, state.Output);
                        if (!TryReplaceImageLink(state.Output, frame))
                        {
                            state.Output.Append("](").Append(frame.Href ?? string.Empty).Append(')');
                        }
                    }
                    return;
                }
            }
        }

        private static bool TryReplaceImageLink(StringBuilder output, Frame frame)
        {
            if (output == null || frame == null || string.IsNullOrEmpty(frame.Href))
            {
                return false;
            }

            int contentStartIndex = frame.ContentStartIndex;
            if (contentStartIndex < 0 || contentStartIndex > output.Length)
            {
                return false;
            }

            string inner = output.ToString(contentStartIndex, output.Length - contentStartIndex);

            // U+1F5BC (FRAME WITH PICTURE) + optional U+FE0F (variation selector-16)
            const string prefixWithVs16 = "[\uD83D\uDDBC\uFE0F";
            const string prefixNoVs16 = "[\uD83D\uDDBC";
            if (!(inner.StartsWith(prefixWithVs16, StringComparison.Ordinal) || inner.StartsWith(prefixNoVs16, StringComparison.Ordinal)))
            {
                return false;
            }

            if (!inner.EndsWith("]", StringComparison.Ordinal))
            {
                return false;
            }

            int prefixLength = inner.StartsWith(prefixWithVs16, StringComparison.Ordinal) ? prefixWithVs16.Length : prefixNoVs16.Length;
            var altText = inner.Substring(prefixLength, inner.Length - prefixLength - 1);
            if (string.IsNullOrWhiteSpace(altText))
            {
                altText = "image";
            }

            // Escape ']' in alt text to keep Markdown well-formed.
            altText = altText.Replace("]", "\\]");

            int start = System.Math.Max(0, contentStartIndex - 1); // include the '[' we inserted for the link
            output.Remove(start, output.Length - start);
            output.Append("![").Append(altText).Append("](").Append(frame.Href).Append(")");
            return true;
        }
    }
}
