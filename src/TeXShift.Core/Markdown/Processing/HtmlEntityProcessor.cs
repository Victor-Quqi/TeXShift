using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using TeXShift.Core.Configuration;

namespace TeXShift.Core.Markdown.Processing
{
    /// <summary>
    /// Protects HTML entities during Markdown processing, then decodes them in the final output.
    /// This ensures proper LaTeX parsing while avoiding double-encoding issues.
    /// 
    /// For code blocks/inline code: entities are preserved as-is (no decoding)
    /// For other content: entities are decoded (&amp;lt; → &lt;)
    /// </summary>
    internal class HtmlEntityProcessor
    {
        // Regex to match HTML entities (e.g., &lt;, &gt;, &amp;, &quot;, &apos;, &#60;, &#x3C;)
        private static readonly Regex HtmlEntityRegex = new Regex(@"&(?:lt|gt|amp|quot|apos|nbsp|#\d+|#x[0-9a-fA-F]+);", RegexOptions.Compiled);
        private const char PlaceholderPrefix = '\uE100';
        private const char PlaceholderSuffix = '\uE101';
        private const int PlaceholderBase = 0xE200;
        private const int PlaceholderRange = 0xF8FF - PlaceholderBase + 1;
        private const int PlaceholderDigits = 3;
        private static readonly Regex PlaceholderRegex = new Regex(
            $"\uE100[\uE200-\uF8FF]{{{PlaceholderDigits}}}\uE101",
            RegexOptions.Compiled);

        /// <summary>
        /// Protects HTML entities in the markdown text by replacing them with placeholders.
        /// </summary>
        public (string protectedText, Dictionary<string, string> entityMap) Protect(string markdown)
        {
            var entityMap = new Dictionary<string, string>();
            var counter = 0;

            var result = HtmlEntityRegex.Replace(markdown, match =>
            {
                var placeholder = CreatePlaceholder(counter++);
                entityMap[placeholder] = match.Value;
                return placeholder;
            });

            return (result, entityMap);
        }

        /// <summary>
        /// Restores and DECODES HTML entities in the generated OneNote XML.
        /// Code elements (T elements inside code blocks) are NOT decoded.
        /// </summary>
        public void RestoreAndDecode(XElement outline, Dictionary<string, string> entityMap, XNamespace ns, OneNoteStyleConfig styleConfig)
        {
            if (entityMap.Count == 0) return;

            foreach (var tElement in outline.Descendants(ns + "T"))
            {
                var cdata = tElement.Nodes().OfType<XCData>().FirstOrDefault();
                if (cdata == null) continue;

                var isCodeBlockLine = IsCodeBlockLine(tElement, ns, styleConfig);
                var inlineCodeRanges = (!isCodeBlockLine && styleConfig != null)
                    ? FindInlineCodeSpanRanges(cdata.Value, styleConfig.GetInlineCodeStyle()?.FontFamily)
                    : null;

                var modified = false;
                var updated = PlaceholderRegex.Replace(cdata.Value, match =>
                {
                    if (entityMap.TryGetValue(match.Value, out var entity))
                    {
                        modified = true;
                        // Decode entities for non-code content, preserve for code.
                        var preserve = isCodeBlockLine || IsWithinRanges(match.Index, inlineCodeRanges);
                        return preserve ? entity : DecodeEntity(entity);
                    }
                    return match.Value;
                });

                if (modified)
                {
                    cdata.ReplaceWith(new XCData(updated));
                }
            }
        }

        /// <summary>
        /// Checks if the current T belongs to a code block line OE generated for fenced/indented code blocks.
        /// </summary>
        private static bool IsCodeBlockLine(XElement tElement, XNamespace ns, OneNoteStyleConfig styleConfig)
        {
            if (tElement == null || styleConfig == null)
            {
                return false;
            }

            // Must be inside a table cell; quote blocks also use tables, so additionally match code font style.
            if (!tElement.Ancestors(ns + "Cell").Any())
            {
                return false;
            }

            var oeParent = tElement.Parent;
            var style = oeParent?.Attribute("style")?.Value ?? string.Empty;
            var codeFontFamily = styleConfig.GetCodeBlockStyle()?.FontFamily ?? string.Empty;
            if (string.IsNullOrWhiteSpace(codeFontFamily))
            {
                return false;
            }

            return style.IndexOf(codeFontFamily, System.StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// Decodes a single HTML entity to its character equivalent.
        /// </summary>
        private string DecodeEntity(string entity)
        {
            return DecodeEntityStatic(entity);
        }

        /// <summary>
        /// Decodes HTML entity placeholders in LaTeX content before passing to MathJax.
        /// This is needed because Math content goes to MathJax before the final XML restore.
        /// </summary>
        /// <param name="latex">LaTeX string potentially containing placeholders</param>
        /// <param name="entityMap">The entity map from Protect()</param>
        /// <returns>LaTeX string with entities decoded</returns>
        public static string DecodeForLatex(string latex, Dictionary<string, string> entityMap)
        {
            if (string.IsNullOrEmpty(latex) || entityMap == null || entityMap.Count == 0)
            {
                return latex;
            }

            return PlaceholderRegex.Replace(latex, match =>
            {
                if (entityMap.TryGetValue(match.Value, out var entity))
                {
                    return DecodeEntityStatic(entity);
                }
                return match.Value;
            });
        }

        /// <summary>
        /// Static version of DecodeEntity for use in static methods.
        /// Decodes a single HTML entity to its character equivalent.
        /// </summary>
        private static string DecodeEntityStatic(string entity)
        {
            switch (entity)
            {
                case "&lt;": return "<";
                case "&gt;": return ">";
                case "&amp;": return "&";
                case "&quot;": return "\"";
                case "&apos;": return "'";
                case "&#39;": return "'";
                case "&nbsp;": return "\u00A0";
                default:
                    // Handle numeric entities
                    if (entity.StartsWith("&#x") || entity.StartsWith("&#X"))
                    {
                        var hex = entity.Substring(3, entity.Length - 4);
                        if (int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out int code) && code > 0 && code <= 0x10FFFF)
                        {
                            return char.ConvertFromUtf32(code);
                        }
                    }
                    else if (entity.StartsWith("&#"))
                    {
                        var num = entity.Substring(2, entity.Length - 3);
                        if (int.TryParse(num, out int code) && code > 0 && code <= 0x10FFFF)
                        {
                            return char.ConvertFromUtf32(code);
                        }
                    }
                    return entity; // Unknown entity, preserve as-is
            }
        }

        private static string CreatePlaceholder(int index)
        {
            var chars = new char[PlaceholderDigits + 2];
            chars[0] = PlaceholderPrefix;
            for (int i = PlaceholderDigits; i >= 1; i--)
            {
                var digit = index % PlaceholderRange;
                chars[i] = (char)(PlaceholderBase + digit);
                index /= PlaceholderRange;
            }
            chars[PlaceholderDigits + 1] = PlaceholderSuffix;

            return new string(chars);
        }

        private static List<(int start, int end)> FindInlineCodeSpanRanges(string html, string inlineCodeFontFamily)
        {
            var ranges = new List<(int start, int end)>();
            if (string.IsNullOrEmpty(html) || string.IsNullOrWhiteSpace(inlineCodeFontFamily))
            {
                return ranges;
            }

            const string openToken = "<span";
            const string closeToken = "</span>";

            int i = 0;
            while (i < html.Length)
            {
                int open = html.IndexOf(openToken, i, System.StringComparison.OrdinalIgnoreCase);
                if (open < 0)
                {
                    break;
                }

                int gt = html.IndexOf('>', open);
                if (gt < 0)
                {
                    break;
                }

                var tag = html.Substring(open, gt - open + 1);
                if (IsInlineCodeSpanTag(tag, inlineCodeFontFamily))
                {
                    int close = html.IndexOf(closeToken, gt + 1, System.StringComparison.OrdinalIgnoreCase);
                    if (close < 0)
                    {
                        i = gt + 1;
                        continue;
                    }

                    int end = close + closeToken.Length;
                    ranges.Add((open, end));
                    i = end;
                    continue;
                }

                i = gt + 1;
            }

            return ranges;
        }

        private static bool IsInlineCodeSpanTag(string spanOpenTag, string inlineCodeFontFamily)
        {
            if (string.IsNullOrEmpty(spanOpenTag) || string.IsNullOrWhiteSpace(inlineCodeFontFamily))
            {
                return false;
            }

            if (spanOpenTag.IndexOf("background-color", System.StringComparison.OrdinalIgnoreCase) < 0)
            {
                return false;
            }

            return spanOpenTag.IndexOf(inlineCodeFontFamily, System.StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsWithinRanges(int index, List<(int start, int end)> ranges)
        {
            if (ranges == null || ranges.Count == 0)
            {
                return false;
            }

            foreach (var r in ranges)
            {
                if (index >= r.start && index < r.end)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
