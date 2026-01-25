namespace TeXShift.Core.Utils
{
    /// <summary>
    /// Escapes plain text for safe inclusion inside OneNote <one:T> HTML content.
    ///
    /// OneNote treats the <one:T> CDATA as HTML, so unescaped '<' would start a tag and '&' would start an entity.
    /// This escaper is intentionally minimal (does NOT escape quotes) so that round-tripped Markdown remains natural.
    /// </summary>
    public static class OneNoteHtmlTextEscaper
    {
        public static string Escape(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            // Must escape '&' first to avoid double-escaping.
            return text
                .Replace("&", "&amp;")
                .Replace("<", "&lt;");
        }
    }
}
