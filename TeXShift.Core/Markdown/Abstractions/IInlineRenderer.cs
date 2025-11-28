using Markdig.Syntax.Inlines;

namespace TeXShift.Core.Markdown.Abstractions
{
    /// <summary>
    /// Converts Markdig inline elements to HTML for embedding in OneNote T elements.
    /// </summary>
    internal interface IInlineRenderer
    {
        /// <summary>
        /// Converts a container of inline elements to an HTML string.
        /// </summary>
        /// <param name="container">The container of inline elements.</param>
        /// <returns>An HTML-formatted string.</returns>
        string Render(ContainerInline container);

        /// <summary>
        /// Converts a collection of inline elements to an HTML string.
        /// </summary>
        /// <param name="inlines">The collection of inline elements.</param>
        /// <returns>An HTML-formatted string.</returns>
        string Render(System.Collections.Generic.IEnumerable<Inline> inlines);
    }
}
