using System;
using System.Net;
using System.Text;

namespace TeXShift.Core.Utils
{
    internal static class MarkdownCodeBlockEntityDecoder
    {
        public static string DecodeEntitiesInFencedCodeBlocks(string markdown)
        {
            if (string.IsNullOrEmpty(markdown))
            {
                return markdown;
            }

            var builder = new StringBuilder(markdown.Length);
            bool inFence = false;
            char fenceChar = '\0';
            int fenceLength = 0;

            int index = 0;
            int lineStart = 0;

            while (index <= markdown.Length)
            {
                bool endOfText = index == markdown.Length;
                bool isLineBreak = !endOfText && (markdown[index] == '\n' || markdown[index] == '\r');

                if (endOfText || isLineBreak)
                {
                    string line = markdown.Substring(lineStart, index - lineStart);
                    string trimmed = line.TrimStart();

                    if (TryGetFence(trimmed, out char currentFenceChar, out int currentFenceLength))
                    {
                        if (!inFence)
                        {
                            inFence = true;
                            fenceChar = currentFenceChar;
                            fenceLength = currentFenceLength;
                        }
                        else if (currentFenceChar == fenceChar && currentFenceLength >= fenceLength)
                        {
                            inFence = false;
                            fenceChar = '\0';
                            fenceLength = 0;
                        }

                        builder.Append(line);
                    }
                    else
                    {
                        builder.Append(inFence ? WebUtility.HtmlDecode(line) : line);
                    }

                    if (!endOfText)
                    {
                        if (markdown[index] == '\r' && index + 1 < markdown.Length && markdown[index + 1] == '\n')
                        {
                            builder.Append("\r\n");
                            index += 2;
                        }
                        else
                        {
                            builder.Append(markdown[index]);
                            index++;
                        }

                        lineStart = index;
                        continue;
                    }

                    break;
                }

                index++;
            }

            return builder.ToString();
        }

        private static bool TryGetFence(string trimmedLine, out char fenceChar, out int fenceLength)
        {
            fenceChar = '\0';
            fenceLength = 0;

            if (string.IsNullOrEmpty(trimmedLine))
            {
                return false;
            }

            char ch = trimmedLine[0];
            if (ch != '`' && ch != '~')
            {
                return false;
            }

            int count = 0;
            while (count < trimmedLine.Length && trimmedLine[count] == ch)
            {
                count++;
            }

            if (count < 3)
            {
                return false;
            }

            fenceChar = ch;
            fenceLength = count;
            return true;
        }
    }
}
