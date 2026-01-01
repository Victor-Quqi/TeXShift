using System.Collections.Generic;

namespace TeXShift.Core.Syntax
{
    /// <summary>
    /// Interface for syntax highlighting services.
    /// </summary>
    public interface ISyntaxHighlighter
    {
        /// <summary>
        /// Highlights an entire code block with proper cross-line state preservation.
        /// This is the preferred method for code blocks with multi-line comments or strings.
        /// </summary>
        /// <param name="lines">The lines of code to highlight.</param>
        /// <param name="language">The programming language identifier (e.g., "csharp", "javascript").</param>
        /// <returns>List of HTML strings with syntax highlighting spans, one per input line.</returns>
        IList<string> HighlightBlock(IList<string> lines, string language);

        /// <summary>
        /// Checks if the specified language is supported for syntax highlighting.
        /// </summary>
        /// <param name="language">The programming language identifier.</param>
        /// <returns>True if the language is supported; otherwise, false.</returns>
        bool IsLanguageSupported(string language);
    }
}
