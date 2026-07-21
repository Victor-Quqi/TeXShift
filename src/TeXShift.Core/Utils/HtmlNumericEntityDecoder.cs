using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace TeXShift.Core.Utils
{
    internal static class HtmlNumericEntityDecoder
    {
        internal enum DecodingPolicy
        {
            OneNoteText,
            ProtectedMarkdownEntity
        }

        private static readonly Regex NumericEntityRegex = new Regex(
            "&#(?:(?:x|X)(?<hex>[0-9A-Fa-f]+)|(?<dec>[0-9]+));",
            RegexOptions.Compiled);

        public static string DecodeInText(string text, DecodingPolicy policy)
        {
            return NumericEntityRegex.Replace(text, match => DecodeMatch(match, policy));
        }

        public static string DecodeEntity(string entity, DecodingPolicy policy)
        {
            if (policy == DecodingPolicy.ProtectedMarkdownEntity)
            {
                if (entity.StartsWith("&#x") || entity.StartsWith("&#X"))
                {
                    string hex = entity.Substring(3, entity.Length - 4);
                    return DecodeValue(entity, hex, true, policy);
                }

                if (entity.StartsWith("&#"))
                {
                    string dec = entity.Substring(2, entity.Length - 3);
                    return DecodeValue(entity, dec, false, policy);
                }

                return entity;
            }

            var match = NumericEntityRegex.Match(entity);
            if (!match.Success || match.Index != 0 || match.Length != entity.Length)
            {
                return entity;
            }

            return DecodeMatch(match, policy);
        }

        private static string DecodeMatch(Match match, DecodingPolicy policy)
        {
            bool isHex = match.Groups["hex"].Success;
            string digits = isHex ? match.Groups["hex"].Value : match.Groups["dec"].Value;
            return DecodeValue(match.Value, digits, isHex, policy);
        }

        private static string DecodeValue(string entity, string digits, bool isHex, DecodingPolicy policy)
        {
            try
            {
                if (!TryParseCodePoint(digits, isHex, policy, out int codePoint))
                {
                    return entity;
                }

                if (codePoint < 0 || codePoint > 0x10FFFF)
                {
                    return entity;
                }

                // OneNote text historically decodes U+0000 and preserves surrogate references.
                // Protected Markdown entities preserve U+0000 and let ConvertFromUtf32 reject surrogates.
                if (policy == DecodingPolicy.OneNoteText
                    && codePoint >= 0xD800
                    && codePoint <= 0xDFFF)
                {
                    return entity;
                }

                if (policy == DecodingPolicy.ProtectedMarkdownEntity && codePoint == 0)
                {
                    return entity;
                }

                return char.ConvertFromUtf32(codePoint);
            }
            catch
            {
                if (policy == DecodingPolicy.OneNoteText)
                {
                    return entity;
                }

                throw;
            }
        }

        private static bool TryParseCodePoint(
            string digits,
            bool isHex,
            DecodingPolicy policy,
            out int codePoint)
        {
            if (policy == DecodingPolicy.OneNoteText)
            {
                codePoint = Convert.ToInt32(digits, isHex ? 16 : 10);
                return true;
            }

            return isHex
                ? int.TryParse(digits, NumberStyles.HexNumber, null, out codePoint)
                : int.TryParse(digits, out codePoint);
        }
    }
}
