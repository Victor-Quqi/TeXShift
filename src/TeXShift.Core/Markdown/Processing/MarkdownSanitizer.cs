using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using TeXShift.Core.Utils;

namespace TeXShift.Core.Markdown.Processing
{
    /// <summary>
    /// Sanitizes Markdown text by removing OneNote formatting artifacts.
    /// This ensures Markdown syntax isn't broken by span tags that OneNote adds for formatting.
    /// </summary>
    internal static class MarkdownSanitizer
    {
        private static readonly Regex SpanTagRegex = new Regex(
            "<\\s*(?<closing>/)?\\s*span\\b(?<attrs>(?:[^>\"']|\"[^\"]*\"|'[^']*')*)>",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Removes all span tags from the text while preserving their content.
        /// This prevents OneNote formatting from breaking Markdown syntax.
        /// </summary>
        /// <param name="text">The text to sanitize</param>
        /// <returns>Sanitized text with span tags removed</returns>
        public static string Sanitize(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            var protectedCode = CodeBlockProtector.Protect(text);
            text = protectedCode.protectedText;

            var output = new StringBuilder(text.Length);
            var preservedSpans = new Stack<bool>();
            int position = 0;

            foreach (Match match in SpanTagRegex.Matches(text))
            {
                output.Append(text, position, match.Index - position);
                position = match.Index + match.Length;

                bool isClosing = match.Groups["closing"].Success;
                if (isClosing)
                {
                    if (preservedSpans.Count > 0 && preservedSpans.Pop())
                    {
                        output.Append("</span>");
                    }
                    continue;
                }

                string attributes = match.Groups["attrs"].Value;
                bool isSelfClosing = attributes.TrimEnd().EndsWith("/", System.StringComparison.Ordinal);
                string color = null;
                bool preserve = !isSelfClosing &&
                    CssColorParser.TryGetColorFromAttributes(attributes, out color);
                if (!isSelfClosing)
                {
                    preservedSpans.Push(preserve);
                }

                if (preserve)
                {
                    output.Append("<span style=\"color:").Append(color).Append("\">");
                }
            }

            output.Append(text, position, text.Length - position);
            return CodeBlockProtector.Restore(output.ToString(), protectedCode.codeMap);
        }
    }
}
