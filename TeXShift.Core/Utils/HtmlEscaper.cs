namespace TeXShift.Core.Utils
{
    /// <summary>
    /// Provides a utility method for escaping HTML special characters.
    /// </summary>
    public static class HtmlEscaper
    {
        /// <summary>
        /// Escapes special characters in a string for safe inclusion in HTML.
        /// </summary>
        /// <param name="text">The input text.</param>
        /// <returns>The escaped string.</returns>
        public static string Escape(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            return text
                .Replace("&", "&amp;") // Must be first
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&#39;");
        }
    }
}