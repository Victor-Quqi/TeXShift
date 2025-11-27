namespace TeXShift.Core.Syntax
{
    /// <summary>
    /// Interface for syntax highlighting services.
    /// </summary>
    public interface ISyntaxHighlighter
    {
        /// <summary>
        /// Highlights a line of code and returns OneNote-compatible HTML.
        /// The output uses only inline styles (e.g., &lt;span style='color:#xxx'&gt;).
        /// </summary>
        /// <param name="line">The line of code to highlight.</param>
        /// <param name="language">The programming language identifier (e.g., "csharp", "javascript").</param>
        /// <returns>HTML string with syntax highlighting spans.</returns>
        string HighlightLine(string line, string language);

        /// <summary>
        /// Checks if the specified language is supported for syntax highlighting.
        /// </summary>
        /// <param name="language">The programming language identifier.</param>
        /// <returns>True if the language is supported; otherwise, false.</returns>
        bool IsLanguageSupported(string language);
    }
}
