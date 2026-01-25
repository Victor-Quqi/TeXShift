using System;
using System.Net;
using System.Text.RegularExpressions;

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
        private static readonly Regex NumericEntityRegex = new Regex(
            "&#(?:(?:x|X)(?<hex>[0-9A-Fa-f]+)|(?<dec>[0-9]+));",
            RegexOptions.Compiled);

        public static string Decode(string text)
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
            string withNumeric = NumericEntityRegex.Replace(text, match =>
            {
                try
                {
                    int codePoint;
                    if (match.Groups["hex"].Success)
                    {
                        codePoint = Convert.ToInt32(match.Groups["hex"].Value, 16);
                    }
                    else
                    {
                        codePoint = Convert.ToInt32(match.Groups["dec"].Value, 10);
                    }

                    // Only decode valid Unicode scalar values.
                    if (codePoint < 0
                        || codePoint > 0x10FFFF
                        || (codePoint >= 0xD800 && codePoint <= 0xDFFF))
                    {
                        return match.Value;
                    }

                    return char.ConvertFromUtf32(codePoint);
                }
                catch
                {
                    return match.Value;
                }
            });

            return WebUtility.HtmlDecode(withNumeric);
        }

        /// <summary>
        /// Decodes entities but keeps literal "&amp;nbsp;" as text (i.e., does NOT convert it to U+00A0).
        /// Useful for fenced code blocks where "&amp;nbsp;" may be part of the code, not whitespace.
        /// </summary>
        public static string DecodePreservingNbspEntity(string text)
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
            string withNumeric = NumericEntityRegex.Replace(text, match =>
            {
                try
                {
                    int codePoint;
                    if (match.Groups["hex"].Success)
                    {
                        codePoint = Convert.ToInt32(match.Groups["hex"].Value, 16);
                    }
                    else
                    {
                        codePoint = Convert.ToInt32(match.Groups["dec"].Value, 10);
                    }

                    if (codePoint < 0
                        || codePoint > 0x10FFFF
                        || (codePoint >= 0xD800 && codePoint <= 0xDFFF))
                    {
                        return match.Value;
                    }

                    return char.ConvertFromUtf32(codePoint);
                }
                catch
                {
                    return match.Value;
                }
            });

            // Protect "&nbsp;" so WebUtility doesn't translate it into U+00A0.
            const string token = "__TEXSHIFT_NBSP_ENTITY__";
            string protectedText = withNumeric.Replace("&nbsp;", token);
            string decoded = WebUtility.HtmlDecode(protectedText);
            return decoded.Replace(token, "&nbsp;");
        }
    }
}
