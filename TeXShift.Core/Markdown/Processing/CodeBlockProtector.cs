using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace TeXShift.Core.Markdown.Processing
{
    /// <summary>
    /// Protects fenced and inline code blocks by replacing them with placeholders.
    /// </summary>
    internal static class CodeBlockProtector
    {
        // Placeholder prefix using Unicode Private Use Area to avoid conflicts
        private const string PlaceholderPrefix = "\uE000";
        private const string PlaceholderSuffix = "\uE001";

        // Regex patterns for code protection
        private static readonly Regex FencedCodeBlockRegex = new Regex(
            @"```[\s\S]*?```|~~~[\s\S]*?~~~",
            RegexOptions.Compiled);

        private static readonly Regex InlineCodeRegex = new Regex(
            @"`[^`\r\n]+`",
            RegexOptions.Compiled);

        public static (string protectedText, Dictionary<string, string> codeMap) Protect(string text)
        {
            var codeMap = new Dictionary<string, string>();
            if (string.IsNullOrEmpty(text))
            {
                return (text, codeMap);
            }

            var counter = 0;

            // Protect fenced code blocks first (``` or ~~~)
            text = FencedCodeBlockRegex.Replace(text, match =>
            {
                var placeholder = $"{PlaceholderPrefix}FENCE{counter++}{PlaceholderSuffix}";
                codeMap[placeholder] = match.Value;
                return placeholder;
            });

            // Protect inline code
            text = InlineCodeRegex.Replace(text, match =>
            {
                var placeholder = $"{PlaceholderPrefix}CODE{counter++}{PlaceholderSuffix}";
                codeMap[placeholder] = match.Value;
                return placeholder;
            });

            return (text, codeMap);
        }

        public static string Restore(string text, Dictionary<string, string> codeMap)
        {
            if (string.IsNullOrEmpty(text) || codeMap == null || codeMap.Count == 0)
            {
                return text;
            }

            foreach (var kvp in codeMap)
            {
                text = text.Replace(kvp.Key, kvp.Value);
            }

            return text;
        }
    }
}
