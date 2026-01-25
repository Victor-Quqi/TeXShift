using System;
using System.Text;
using TeXShift.Core.Utils;

namespace TeXShift.Core.OneNoteToMarkdown.Inlines
{
    /// <summary>
    /// Best-effort HTML tag stripper (used to remove syntax highlighting spans in code blocks).
    /// </summary>
    internal static class HtmlStripper
    {
        public static string Strip(string html)
        {
            return StripCore(html, preserveNbspEntity: false);
        }

        public static string StripPreservingNbspEntity(string html)
        {
            return StripCore(html, preserveNbspEntity: true);
        }

        private static string StripCore(string html, bool preserveNbspEntity)
        {
            if (string.IsNullOrEmpty(html))
            {
                return string.Empty;
            }

            var sb = new StringBuilder(html.Length);

            int i = 0;
            while (i < html.Length)
            {
                int lt = html.IndexOf('<', i);
                if (lt < 0)
                {
                    AppendDecodedText(sb, html.Substring(i), preserveNbspEntity);
                    break;
                }

                if (lt > i)
                {
                    AppendDecodedText(sb, html.Substring(i, lt - i), preserveNbspEntity);
                }

                int gt = html.IndexOf('>', lt + 1);
                if (gt < 0)
                {
                    AppendDecodedText(sb, html.Substring(lt), preserveNbspEntity);
                    break;
                }

                // Tag content without < and >
                string rawTag = html.Substring(lt + 1, gt - lt - 1).Trim();
                i = gt + 1;

                if (rawTag.Length == 0)
                {
                    continue;
                }

                // Preserve line breaks where possible.
                string tagName = GetTagName(rawTag);
                if (string.Equals(tagName, "br", StringComparison.OrdinalIgnoreCase))
                {
                    sb.Append("\n");
                }
            }

            return NormalizeText(sb.ToString());
        }

        private static string GetTagName(string rawTag)
        {
            if (string.IsNullOrEmpty(rawTag))
            {
                return null;
            }

            // Trim leading '/' for closing tags.
            int start = rawTag[0] == '/' ? 1 : 0;
            int end = start;
            while (end < rawTag.Length)
            {
                char c = rawTag[end];
                if (char.IsWhiteSpace(c) || c == '/' || c == '>')
                {
                    break;
                }
                end++;
            }

            if (end <= start)
            {
                return null;
            }

            return rawTag.Substring(start, end - start);
        }

        private static void AppendDecodedText(StringBuilder sb, string text, bool preserveNbspEntity)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            sb.Append(preserveNbspEntity
                ? OneNoteHtmlEntityDecoder.DecodePreservingNbspEntity(text)
                : OneNoteHtmlEntityDecoder.Decode(text));
        }

        private static string NormalizeText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            return text.Replace('\u00A0', ' ');
        }
    }
}
