using System;
using System.Text;

namespace TeXShift.Core.Markdown.Processing
{
    internal static class HtmlLineBreakParser
    {
        public static bool IsLineBreakTag(string htmlTag)
        {
            if (string.IsNullOrWhiteSpace(htmlTag))
            {
                return false;
            }

            var tag = htmlTag.Trim();
            if (tag.Length < 4 || tag[0] != '<' || tag[tag.Length - 1] != '>')
            {
                return false;
            }

            var declaration = tag.Substring(1, tag.Length - 2).Trim();
            if (declaration.StartsWith("/", StringComparison.Ordinal))
            {
                return false;
            }

            if (declaration.EndsWith("/", StringComparison.Ordinal))
            {
                declaration = declaration.Substring(0, declaration.Length - 1).TrimEnd();
            }

            if (declaration.Length < 2 ||
                !declaration.StartsWith("br", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return declaration.Length == 2 || char.IsWhiteSpace(declaration[2]);
        }

        public static bool TryConvertLineBreakOnlyHtml(string html, out string lineBreaks)
        {
            lineBreaks = null;
            if (string.IsNullOrWhiteSpace(html))
            {
                return false;
            }

            var builder = new StringBuilder();
            var index = 0;

            while (index < html.Length)
            {
                while (index < html.Length && char.IsWhiteSpace(html[index]))
                {
                    index++;
                }

                if (index >= html.Length)
                {
                    break;
                }

                if (!TryReadTag(html, ref index, out var tag) || !IsLineBreakTag(tag))
                {
                    return false;
                }

                builder.Append('\n');
            }

            if (builder.Length == 0)
            {
                return false;
            }

            lineBreaks = builder.ToString();
            return true;
        }

        private static bool TryReadTag(string html, ref int index, out string tag)
        {
            tag = null;
            if (index >= html.Length || html[index] != '<')
            {
                return false;
            }

            var start = index;
            char quote = '\0';
            index++;

            while (index < html.Length)
            {
                var current = html[index];
                if (quote != '\0')
                {
                    if (current == quote)
                    {
                        quote = '\0';
                    }
                }
                else if (current == '\'' || current == '"')
                {
                    quote = current;
                }
                else if (current == '>')
                {
                    index++;
                    tag = html.Substring(start, index - start);
                    return true;
                }

                index++;
            }

            return false;
        }
    }
}
