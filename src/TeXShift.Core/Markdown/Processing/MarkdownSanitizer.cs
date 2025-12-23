using System.Text.RegularExpressions;

namespace TeXShift.Core.Markdown.Processing
{
    /// <summary>
    /// Sanitizes Markdown text by removing OneNote formatting artifacts.
    /// This ensures Markdown syntax isn't broken by span tags that OneNote adds for formatting.
    /// </summary>
    internal static class MarkdownSanitizer
    {
        // Regex to remove all span tags (including style spans from OneNote font formatting)
        // These can break Markdown syntax like "- [ ]" when they wrap list markers
        private static readonly Regex SpanTagRegex = new Regex(@"<span\s[^>]*>(.*?)</span>", RegexOptions.Compiled | RegexOptions.Singleline);

        /// <summary>
        /// Removes all span tags from the text while preserving their content.
        /// This prevents OneNote formatting from breaking Markdown syntax.
        /// </summary>
        /// <param name="text">The text to sanitize</param>
        /// <returns>Sanitized text with span tags removed</returns>
        public static string Sanitize(string text)
        {
            // Remove all span tags (lang, style, etc.) that OneNote adds for formatting
            // These can break Markdown syntax like "- [ ]" when they wrap list markers
            while (SpanTagRegex.IsMatch(text))
            {
                text = SpanTagRegex.Replace(text, "$1");
            }

            return text;
        }
    }
}
