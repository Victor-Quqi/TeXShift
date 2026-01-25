using System;
using System.Collections.Generic;
using System.Text;
using TeXShift.Core.Utils;

namespace TeXShift.Core.OneNoteToMarkdown.Handlers
{
    internal sealed partial class TableElementHandler
    {
        private static bool HasSignificantTextInHtml(string html)
        {
            if (string.IsNullOrEmpty(html))
            {
                return false;
            }

            int i = 0;
            while (i < html.Length)
            {
                int lt = html.IndexOf('<', i);
                if (lt < 0)
                {
                    return ContainsNonWhitespaceText(html.Substring(i));
                }

                if (lt > i && ContainsNonWhitespaceText(html.Substring(i, lt - i)))
                {
                    return true;
                }

                int gt = html.IndexOf('>', lt + 1);
                if (gt < 0)
                {
                    return ContainsNonWhitespaceText(html.Substring(lt));
                }

                i = gt + 1;
            }

            return false;
        }

        private static bool IsHtmlFullyBold(string html)
        {
            if (string.IsNullOrEmpty(html))
            {
                return true;
            }

            var spanStack = new Stack<bool>();
            bool boldActive = false;
            bool sawText = false;

            int i = 0;
            while (i < html.Length)
            {
                int lt = html.IndexOf('<', i);
                if (lt < 0)
                {
                    if (ContainsNonWhitespaceText(html.Substring(i)))
                    {
                        sawText = true;
                        if (!boldActive)
                        {
                            return false;
                        }
                    }
                    break;
                }

                if (lt > i && ContainsNonWhitespaceText(html.Substring(i, lt - i)))
                {
                    sawText = true;
                    if (!boldActive)
                    {
                        return false;
                    }
                }

                int gt = html.IndexOf('>', lt + 1);
                if (gt < 0)
                {
                    if (ContainsNonWhitespaceText(html.Substring(lt)))
                    {
                        sawText = true;
                        if (!boldActive)
                        {
                            return false;
                        }
                    }
                    break;
                }

                string rawTag = html.Substring(lt + 1, gt - lt - 1).Trim();
                i = gt + 1;

                if (rawTag.Length == 0)
                {
                    continue;
                }

                bool isClosing = rawTag[0] == '/';
                string tagName = GetTagName(isClosing ? rawTag.Substring(1) : rawTag);
                if (!string.Equals(tagName, "span", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (isClosing)
                {
                    boldActive = spanStack.Count > 0 ? spanStack.Pop() : false;
                    continue;
                }

                spanStack.Push(boldActive);
                var style = NormalizeStyleForContains(GetStyle(rawTag));
                if (style.Contains("font-weight:normal"))
                {
                    boldActive = false;
                }
                else if (style.Contains("font-weight:bold"))
                {
                    boldActive = true;
                }
            }

            return sawText;
        }

        private static bool ContainsNonWhitespaceText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            var decoded = OneNoteHtmlEntityDecoder.Decode(text);
            if (string.IsNullOrEmpty(decoded))
            {
                return false;
            }

            foreach (char ch in decoded)
            {
                if (!char.IsWhiteSpace(ch))
                {
                    return true;
                }
            }

            return false;
        }

        private static string GetTagName(string rawTag)
        {
            if (string.IsNullOrWhiteSpace(rawTag))
            {
                return null;
            }

            int i = 0;
            while (i < rawTag.Length && char.IsWhiteSpace(rawTag[i]))
            {
                i++;
            }

            int start = i;
            while (i < rawTag.Length && !char.IsWhiteSpace(rawTag[i]))
            {
                i++;
            }

            if (i <= start)
            {
                return null;
            }

            var name = rawTag.Substring(start, i - start);
            if (name.EndsWith("/", StringComparison.Ordinal))
            {
                name = name.Substring(0, name.Length - 1);
            }

            return name;
        }

        private static bool TryUnwrapHeaderSpan(string html, out string innerHtml)
        {
            innerHtml = html;
            if (string.IsNullOrEmpty(html))
            {
                return false;
            }

            var match = HtmlRegexes.OuterSpan.Match(html);
            if (!match.Success)
            {
                return false;
            }

            var attrs = match.Groups["attrs"].Value ?? string.Empty;
            var style = GetStyle(attrs);
            if (string.IsNullOrEmpty(style))
            {
                return false;
            }

            if (!NormalizeStyleForContains(style).Contains("font-weight:bold"))
            {
                return false;
            }

            innerHtml = match.Groups["inner"].Value ?? string.Empty;
            return true;
        }

        private static string GetStyle(string attrs)
        {
            if (string.IsNullOrEmpty(attrs))
            {
                return null;
            }

            var match = HtmlRegexes.StyleAttr.Match(attrs);
            return match.Success ? match.Groups["v"].Value : null;
        }

        private static string NormalizeStyleForContains(string style)
        {
            if (string.IsNullOrEmpty(style))
            {
                return string.Empty;
            }

            // OneNote may insert whitespace/newlines into the style value.
            // Normalize for stable substring checks.
            var sb = new StringBuilder(style.Length);
            foreach (char ch in style)
            {
                if (!char.IsWhiteSpace(ch))
                {
                    sb.Append(char.ToLowerInvariant(ch));
                }
            }

            return sb.ToString();
        }
    }
}

