using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace TeXShift.Core.Markdown.Processing
{
    /// <summary>
    /// Converts LaTeX-style delimiters to Markdown-style math delimiters.
    /// Transforms \(...\) to $...$ and \[...\] to $$...$$ while protecting code blocks.
    /// </summary>
    internal static class LatexDelimiterConverter
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

        // Regex patterns for LaTeX delimiters (non-greedy, supports multiline for block math)
        private static readonly Regex InlineMathRegex = new Regex(
            @"\\\((.+?)\\\)",
            RegexOptions.Compiled | RegexOptions.Singleline);

        private static readonly Regex BlockMathRegex = new Regex(
            @"\\\[([\s\S]+?)\\\]",
            RegexOptions.Compiled);

        // Regex to match multiline $$ blocks where $$ is not on its own line
        // Matches $$<non-whitespace>...$$ where content contains newlines
        private static readonly Regex MultilineMathBlockRegex = new Regex(
            @"\$\$([^\$\n\r][^\$]*)\$\$",
            RegexOptions.Compiled | RegexOptions.Singleline);

        /// <summary>
        /// Converts LaTeX delimiters to Markdown math syntax and normalizes multiline math blocks.
        /// Protects code blocks to prevent false conversions.
        /// </summary>
        /// <param name="markdown">The markdown text with potential LaTeX delimiters</param>
        /// <returns>Markdown with converted math delimiters and normalized math blocks</returns>
        public static string Convert(string markdown)
        {
            if (string.IsNullOrEmpty(markdown))
            {
                return markdown;
            }

            // Quick check: skip processing if no math content present
            if (!ContainsMathContent(markdown))
            {
                return markdown;
            }

            // Step 1: Protect code blocks
            var (protectedText, codeMap) = ProtectCodeBlocks(markdown);

            // Step 2: Convert LaTeX delimiters (\[...\] → $$...$$, \(...\) → $...$)
            var converted = ConvertDelimiters(protectedText);

            // Step 3: Normalize multiline math blocks (ensure $$ is on its own line)
            converted = NormalizeMultilineMathBlocks(converted);

            // Step 4: Restore code blocks
            var result = RestoreCodeBlocks(converted, codeMap);

            return result;
        }

        private static bool ContainsMathContent(string text)
        {
            // Check for LaTeX delimiters or $$ math blocks
            return text.Contains(@"\(") || text.Contains(@"\[") || text.Contains("$$");
        }

        private static (string protectedText, Dictionary<string, string> codeMap) ProtectCodeBlocks(string text)
        {
            var codeMap = new Dictionary<string, string>();
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

        private static string ConvertDelimiters(string text)
        {
            // Convert block math first: \[...\] → $$...$$
            text = BlockMathRegex.Replace(text, match =>
            {
                var content = match.Groups[1].Value;
                return $"$${content}$$";
            });

            // Convert inline math: \(...\) → $...$
            text = InlineMathRegex.Replace(text, match =>
            {
                var content = match.Groups[1].Value;
                return $"${content}$";
            });

            return text;
        }

        /// <summary>
        /// Normalizes multiline math blocks by ensuring $$ delimiters are on their own lines.
        /// This is required for Markdig to correctly recognize block-level math.
        /// Without this, Markdig's GenericAttributes extension may corrupt curly braces like {pmatrix}.
        /// </summary>
        private static string NormalizeMultilineMathBlocks(string text)
        {
            return MultilineMathBlockRegex.Replace(text, match =>
            {
                var content = match.Groups[1].Value;

                // Only normalize if content contains newlines (true multiline block)
                if (content.Contains("\n"))
                {
                    // Ensure $$ is on its own line: $$\ncontent\n$$
                    return $"$$\n{content}\n$$";
                }

                // Single-line math block - leave as is
                return match.Value;
            });
        }

        private static string RestoreCodeBlocks(string text, Dictionary<string, string> codeMap)
        {
            foreach (var kvp in codeMap)
            {
                text = text.Replace(kvp.Key, kvp.Value);
            }
            return text;
        }
    }
}
