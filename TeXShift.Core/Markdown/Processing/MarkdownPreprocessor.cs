using System.Text;

namespace TeXShift.Core.Markdown.Processing
{
    /// <summary>
    /// Applies minimal, line-based normalization before Markdown parsing.
    /// </summary>
    internal static class MarkdownPreprocessor
    {
        public static string Normalize(string markdown)
        {
            if (string.IsNullOrEmpty(markdown))
            {
                return markdown;
            }

            var builder = new StringBuilder(markdown.Length);
            var lineStart = 0;

            for (int i = 0; i < markdown.Length; i++)
            {
                var ch = markdown[i];
                if (ch != '\r' && ch != '\n') continue;

                var line = markdown.Substring(lineStart, i - lineStart);
                builder.Append(DecodeBlockquoteMarkers(line));

                if (ch == '\r' && i + 1 < markdown.Length && markdown[i + 1] == '\n')
                {
                    builder.Append("\r\n");
                    i++;
                }
                else
                {
                    builder.Append(ch);
                }

                lineStart = i + 1;
            }

            if (lineStart <= markdown.Length)
            {
                builder.Append(DecodeBlockquoteMarkers(markdown.Substring(lineStart)));
            }

            return builder.ToString();
        }

        private static string DecodeBlockquoteMarkers(string line)
        {
            if (string.IsNullOrEmpty(line))
            {
                return line;
            }

            var builder = new StringBuilder(line.Length);
            var index = 0;

            while (index < line.Length && (line[index] == ' ' || line[index] == '\t'))
            {
                builder.Append(line[index]);
                index++;
            }

            var sawQuote = false;
            while (index < line.Length)
            {
                if (IsEncodedQuote(line, index))
                {
                    builder.Append('>');
                    index += 4;
                    sawQuote = true;
                    continue;
                }

                if (sawQuote && (line[index] == ' ' || line[index] == '\t'))
                {
                    builder.Append(line[index]);
                    index++;
                    continue;
                }

                break;
            }

            builder.Append(line.Substring(index));
            return builder.ToString();
        }

        private static bool IsEncodedQuote(string line, int index)
        {
            return index + 3 < line.Length
                && line[index] == '&'
                && line[index + 1] == 'g'
                && line[index + 2] == 't'
                && line[index + 3] == ';';
        }
    }
}
