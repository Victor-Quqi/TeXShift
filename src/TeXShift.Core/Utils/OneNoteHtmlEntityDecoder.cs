using System.Net;

namespace TeXShift.Core.Utils
{
    /// <summary>
    /// Decodes HTML entities from OneNote <one:T> content.
    ///
    /// NOTE: .NET Framework's WebUtility.HtmlDecode does not decode numeric character references
    /// for code points above U+FFFF (e.g., &#127881;), but OneNote uses them for some emoji.
    /// We decode numeric entities first (single-pass semantics), then delegate to WebUtility for
    /// named entities and the remaining cases.
    /// </summary>
    internal static class OneNoteHtmlEntityDecoder
    {
        public static string Decode(string text)
        {
            return DecodeCore(text, preserveNbspEntity: false);
        }

        /// <summary>
        /// Decodes entities but keeps literal "&amp;nbsp;" as text (i.e., does NOT convert it to U+00A0).
        /// Useful for fenced code blocks where "&amp;nbsp;" may be part of the code, not whitespace.
        /// </summary>
        public static string DecodePreservingNbspEntity(string text)
        {
            return DecodeCore(text, preserveNbspEntity: true);
        }

        private static string DecodeCore(string text, bool preserveNbspEntity)
        {
            if (text == null)
            {
                return null;
            }

            if (text.Length == 0)
            {
                return string.Empty;
            }

            // Decode numeric entities first to support astral code points.
            string withNumeric = HtmlNumericEntityDecoder.DecodeInText(
                text,
                HtmlNumericEntityDecoder.DecodingPolicy.OneNoteText);

            if (!preserveNbspEntity)
            {
                return WebUtility.HtmlDecode(withNumeric);
            }

            // Protect "&nbsp;" so WebUtility doesn't translate it into U+00A0.
            const string token = "__TEXSHIFT_NBSP_ENTITY__";
            string protectedText = withNumeric.Replace("&nbsp;", token);
            string decoded = WebUtility.HtmlDecode(protectedText);
            return decoded.Replace(token, "&nbsp;");
        }
    }
}
