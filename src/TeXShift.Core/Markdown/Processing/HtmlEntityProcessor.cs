using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

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
        public void RestoreAndDecode(XElement outline, Dictionary<string, string> entityMap, XNamespace ns)
        {
            if (entityMap.Count == 0) return;

            foreach (var tElement in outline.Descendants(ns + "T"))
            {
                var cdata = tElement.Nodes().OfType<XCData>().FirstOrDefault();
                if (cdata == null) continue;

                // Check if this T element is inside a code block (Cell with shadingColor)
                var isInCodeBlock = IsInCodeBlock(tElement, ns);

                var modified = false;
                var updated = PlaceholderRegex.Replace(cdata.Value, match =>
                {
                    if (entityMap.TryGetValue(match.Value, out var entity))
                    {
                        modified = true;
                        // Decode entities for non-code content, preserve for code
                        return isInCodeBlock ? entity : DecodeEntity(entity);
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
        /// Check if a T element is inside a code block (Table Cell with shadingColor background).
        /// </summary>
        private bool IsInCodeBlock(XElement tElement, XNamespace ns)
        {
            var ancestor = tElement.Ancestors(ns + "Cell").FirstOrDefault();
            if (ancestor != null && ancestor.Attribute("shadingColor") != null)
            {
                return true;
            }
            
            // Also check for inline code by looking for font-family:Consolas in style
            var oeParent = tElement.Parent;
            if (oeParent != null)
            {
                var style = oeParent.Attribute("style")?.Value ?? "";
                if (style.Contains("Consolas") || style.Contains("monospace"))
                {
                    return true;
                }
            }

            return false;
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
    }
}
