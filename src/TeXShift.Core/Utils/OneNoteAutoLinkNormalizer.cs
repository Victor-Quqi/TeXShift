using System;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace TeXShift.Core.Utils
{
    /// <summary>
    /// OneNote may "linkify" plain text by wrapping it with a simple &lt;a href="..."&gt;...&lt;/a&gt; HTML tag
    /// inside &lt;one:T&gt; CDATA. This breaks Markdown round-tripping because the input is no longer pure Markdown.
    ///
    /// This normalizer removes only the auto-generated hyperlink wrapper (keeps displayed text),
    /// while preserving other HTML tags for potential future support.
    /// </summary>
    internal static class OneNoteAutoLinkNormalizer
    {
        // Matches <a ...>...</a> across line breaks.
        private static readonly Regex AnchorRegex = new Regex(
            "<a\\b(?<attrs>[^>]*)>(?<text>.*?)</a>",
            RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.Singleline);

        // Extracts href="..." / href='...' / href=... from the anchor attributes.
        private static readonly Regex HrefAttrRegex = new Regex(
            "\\bhref\\s*=\\s*(\"(?<v>[^\"]*)\"|'(?<v>[^']*)'|(?<v>[^\\s>]+))",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static string Normalize(string htmlOrText)
        {
            if (string.IsNullOrEmpty(htmlOrText) || htmlOrText.IndexOf("<a", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return htmlOrText;
            }

            return AnchorRegex.Replace(htmlOrText, match =>
            {
                var attrs = match.Groups["attrs"]?.Value ?? string.Empty;
                var innerText = match.Groups["text"]?.Value ?? string.Empty;

                if (!IsSimpleHrefOnlyAnchor(attrs))
                {
                    return match.Value;
                }

                var href = GetHref(attrs);
                var decodedText = WebUtility.HtmlDecode(innerText);
                var decodedHref = WebUtility.HtmlDecode(href);

                // Unwrap only when it looks like OneNote auto-linkification (high confidence),
                // so we don't destroy user-authored <a> tags (future HTML support).
                if (IsLikelyOneNoteAutoLink(decodedText, decodedHref))
                {
                    return decodedText;
                }

                return match.Value;
            });
        }

        private static bool IsSimpleHrefOnlyAnchor(string attrs)
        {
            if (string.IsNullOrWhiteSpace(attrs))
            {
                return false;
            }

            bool hasHref = false;
            int i = 0;
            while (i < attrs.Length)
            {
                // Skip whitespace between attributes.
                while (i < attrs.Length && char.IsWhiteSpace(attrs[i]))
                {
                    i++;
                }
                if (i >= attrs.Length)
                {
                    break;
                }

                // Read attribute name.
                int nameStart = i;
                while (i < attrs.Length)
                {
                    char c = attrs[i];
                    if (char.IsWhiteSpace(c) || c == '=' || c == '>' || c == '/')
                    {
                        break;
                    }
                    i++;
                }

                if (i <= nameStart)
                {
                    i++;
                    continue;
                }

                string name = attrs.Substring(nameStart, i - nameStart);

                // Skip whitespace after name.
                while (i < attrs.Length && char.IsWhiteSpace(attrs[i]))
                {
                    i++;
                }

                // Skip attribute value (if any). We must not treat "v=4" in a URL query string as another attribute.
                if (i < attrs.Length && attrs[i] == '=')
                {
                    i++;
                    while (i < attrs.Length && char.IsWhiteSpace(attrs[i]))
                    {
                        i++;
                    }

                    if (i < attrs.Length)
                    {
                        char quote = attrs[i];
                        if (quote == '"' || quote == '\'')
                        {
                            i++;
                            while (i < attrs.Length && attrs[i] != quote)
                            {
                                i++;
                            }
                            if (i < attrs.Length)
                            {
                                i++;
                            }
                        }
                        else
                        {
                            while (i < attrs.Length && !char.IsWhiteSpace(attrs[i]) && attrs[i] != '>')
                            {
                                i++;
                            }
                        }
                    }
                }

                if (string.Equals(name, "href", StringComparison.OrdinalIgnoreCase))
                {
                    hasHref = true;
                    continue;
                }

                // Any additional attribute makes it less likely to be OneNote's auto-generated wrapper.
                return false;
            }

            return hasHref;
        }

        private static string GetHref(string attrs)
        {
            var m = HrefAttrRegex.Match(attrs ?? string.Empty);
            if (!m.Success)
            {
                return string.Empty;
            }

            return m.Groups["v"]?.Value ?? string.Empty;
        }

        private static bool IsLikelyOneNoteAutoLink(string decodedText, string decodedHref)
        {
            var text = (decodedText ?? string.Empty).Trim();
            var href = (decodedHref ?? string.Empty).Trim();

            if (text.Length == 0 || href.Length == 0)
            {
                return false;
            }

            // Most common: display text equals href (after HTML decoding).
            if (string.Equals(text, href, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // Auto-linkification often adds the scheme (http/https/mailto) to href only.
            var strippedHttp = StripHttpScheme(href);
            if (!string.IsNullOrEmpty(strippedHttp) && string.Equals(text, strippedHttp, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            var strippedMailto = StripMailtoScheme(href);
            if (!string.IsNullOrEmpty(strippedMailto) && string.Equals(text, strippedMailto, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // File/UNC links: href may be file://server/share while the display is \\server\share (or C:\path).
            if (href.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
            {
                var expectedFileHref = TryConvertWindowsPathToFileUrl(text);
                if (!string.IsNullOrEmpty(expectedFileHref) &&
                    string.Equals(NormalizeFileUrl(href), NormalizeFileUrl(expectedFileHref), StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static string StripHttpScheme(string href)
        {
            if (string.IsNullOrWhiteSpace(href))
            {
                return null;
            }

            const string http = "http://";
            const string https = "https://";

            if (href.StartsWith(http, StringComparison.OrdinalIgnoreCase))
            {
                return href.Substring(http.Length);
            }

            if (href.StartsWith(https, StringComparison.OrdinalIgnoreCase))
            {
                return href.Substring(https.Length);
            }

            return null;
        }

        private static string StripMailtoScheme(string href)
        {
            if (string.IsNullOrWhiteSpace(href))
            {
                return null;
            }

            const string mailto = "mailto:";
            if (href.StartsWith(mailto, StringComparison.OrdinalIgnoreCase))
            {
                return href.Substring(mailto.Length);
            }

            return null;
        }

        private static string TryConvertWindowsPathToFileUrl(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            // UNC path: \\server\share\path -> file://server/share/path
            if (text.Length >= 2 && text[0] == '\\' && text[1] == '\\')
            {
                var path = text.Trim();
                while (path.Length >= 2 && path[0] == '\\' && path[1] == '\\')
                {
                    path = path.Substring(2);
                }

                path = path.Replace('\\', '/');
                return "file://" + path;
            }

            // Drive path: C:\Users\test -> file:///C:/Users/test
            if (text.Length >= 3 && char.IsLetter(text[0]) && text[1] == ':' && (text[2] == '\\' || text[2] == '/'))
            {
                var path = text.Trim().Replace('\\', '/');
                // Ensure "C:/" form (not "C:Users").
                if (path.Length >= 2 && path[1] == ':' && path.Length == 2)
                {
                    path += "/";
                }
                else if (path.Length >= 3 && path[1] == ':' && path[2] != '/')
                {
                    path = path.Substring(0, 2) + "/" + path.Substring(2);
                }

                return "file:///" + path;
            }

            return null;
        }

        private static string NormalizeFileUrl(string fileUrl)
        {
            if (string.IsNullOrWhiteSpace(fileUrl))
            {
                return string.Empty;
            }

            // Treat file://server/share and file:////server/share as equivalent for comparison.
            var url = fileUrl.Trim();
            if (url.StartsWith("file:////", StringComparison.OrdinalIgnoreCase))
            {
                url = "file://" + url.Substring("file:////".Length);
            }
            else if (url.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
            {
                // Keep drive-path canonical form.
                url = "file:///" + url.Substring("file:///".Length);
            }
            else if (url.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
            {
                url = "file://" + url.Substring("file://".Length);
            }

            // OneNote may produce extra slashes when it linkifies text that contains escaped backslashes
            // (e.g., display: "\\\\server\\share", href: "file://server/share").
            // Collapse repeated slashes in the path part for stable comparisons.
            if (url.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
            {
                var rest = url.Substring("file:///".Length);
                return "file:///" + CollapseRepeatedSlashes(rest);
            }
            if (url.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
            {
                var rest = url.Substring("file://".Length);
                return "file://" + CollapseRepeatedSlashes(rest);
            }

            return url;
        }

        private static string CollapseRepeatedSlashes(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            var sb = new StringBuilder(text.Length);
            bool lastSlash = false;
            foreach (var ch in text)
            {
                if (ch == '/')
                {
                    if (lastSlash)
                    {
                        continue;
                    }
                    lastSlash = true;
                    sb.Append('/');
                    continue;
                }

                lastSlash = false;
                sb.Append(ch);
            }

            return sb.ToString();
        }
    }
}
