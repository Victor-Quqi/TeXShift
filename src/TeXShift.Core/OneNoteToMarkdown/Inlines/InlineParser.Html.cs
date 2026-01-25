using System;
using System.Text.RegularExpressions;

namespace TeXShift.Core.OneNoteToMarkdown.Inlines
{
    internal sealed partial class InlineParser
    {
        private static string GetTagName(string rawTag)
        {
            if (string.IsNullOrEmpty(rawTag))
            {
                return null;
            }

            int end = 0;
            while (end < rawTag.Length)
            {
                char c = rawTag[end];
                if (char.IsWhiteSpace(c) || c == '/' || c == '>')
                {
                    break;
                }
                end++;
            }

            if (end == 0)
            {
                return null;
            }

            return rawTag.Substring(0, end);
        }

        private static string GetAttributeValue(string rawTag, Regex regex)
        {
            if (string.IsNullOrEmpty(rawTag))
            {
                return null;
            }

            var m = regex.Match(rawTag);
            if (!m.Success)
            {
                return null;
            }

            return m.Groups["v"]?.Value;
        }
    }
}

