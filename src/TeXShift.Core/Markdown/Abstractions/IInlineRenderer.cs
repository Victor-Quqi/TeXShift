using Markdig.Syntax.Inlines;
using System.Collections.Generic;
using System.Threading.Tasks;

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
        Task<string> RenderAsync(ContainerInline container);

        /// <summary>
        /// Converts a collection of inline elements to an HTML string.
        /// </summary>
        /// <param name="inlines">The collection of inline elements.</param>
        /// <returns>An HTML-formatted string.</returns>
        Task<string> RenderAsync(IEnumerable<Inline> inlines);

        /// <summary>
        /// Sets or gets the entity decoder function for decoding HTML entities in math content.
        /// This should be set by the converter before rendering.
        /// </summary>
        System.Func<string, string> EntityDecoder { get; set; }
    }
}
