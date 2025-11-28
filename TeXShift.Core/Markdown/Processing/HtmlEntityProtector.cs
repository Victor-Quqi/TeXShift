using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace TeXShift.Core.Markdown.Processing
{
    /// <summary>
    /// Protects and restores HTML entities during Markdown processing.
    /// Prevents double-encoding issues when Markdig processes content that contains
    /// HTML entities like &amp;lt;, &amp;gt;, etc.
    /// </summary>
    internal class HtmlEntityProtector
    {
        // Regex to match HTML entities (e.g., &lt;, &gt;, &amp;, &quot;, &apos;, &#60;, &#x3C;)
        private static readonly Regex HtmlEntityRegex = new Regex(@"&(?:lt|gt|amp|quot|apos|#\d+|#x[0-9a-fA-F]+);", RegexOptions.Compiled);

        /// <summary>
        /// Protects HTML entities in the markdown text by replacing them with placeholders.
        /// This prevents Markdig from auto-decoding entities like &amp;lt; to &lt;, which would
        /// cause double-escaping issues when HtmlEscaper re-encodes them.
        /// </summary>
        /// <param name="markdown">The markdown text containing HTML entities</param>
        /// <returns>A tuple of (protected markdown, entity map for restoration)</returns>
        public (string protectedText, Dictionary<string, string> entityMap) Protect(string markdown)
        {
            var entityMap = new Dictionary<string, string>();
            var counter = 0;

            var result = HtmlEntityRegex.Replace(markdown, match =>
            {
                // Use Unicode Replacement Character (U+FFFD) as placeholder delimiter
                // This character is extremely rare in normal text
                var placeholder = $"\uFFFD{counter++}\uFFFD";
                entityMap[placeholder] = match.Value;
                return placeholder;
            });

            return (result, entityMap);
        }

        /// <summary>
        /// Restores HTML entities in the generated OneNote XML by replacing placeholders
        /// with their original entity strings.
        /// </summary>
        /// <param name="outline">The OneNote Outline element to process</param>
        /// <param name="entityMap">The map of placeholders to original entities</param>
        /// <param name="ns">The OneNote XML namespace</param>
        public void Restore(XElement outline, Dictionary<string, string> entityMap, XNamespace ns)
        {
            if (entityMap.Count == 0) return; // Optimization: skip if no entities to restore

            foreach (var tElement in outline.Descendants(ns + "T"))
            {
                var cdata = tElement.Nodes().OfType<XCData>().FirstOrDefault();
                if (cdata == null) continue;

                var text = cdata.Value;
                var modified = false;

                foreach (var kvp in entityMap)
                {
                    if (text.Contains(kvp.Key))
                    {
                        text = text.Replace(kvp.Key, kvp.Value);
                        modified = true;
                    }
                }

                if (modified)
                {
                    cdata.ReplaceWith(new XCData(text));
                }
            }
        }
    }
}
