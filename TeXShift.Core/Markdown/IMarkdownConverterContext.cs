using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using System.Collections.Generic;
using System.Xml.Linq;

namespace TeXShift.Core.Markdown
{
    /// <summary>
    /// Provides a context for block handlers during Markdown to OneNote XML conversion.
    /// This allows handlers to access shared resources like style configurations, the XML namespace,
    /// and methods to recursively process nested blocks or inlines.
    /// </summary>
    internal interface IMarkdownConverterContext
    {
        /// <summary>
        /// Gets the XML namespace for OneNote.
        /// </summary>
        XNamespace OneNoteNamespace { get; }

        /// <summary>
        /// Gets the style configuration for OneNote elements.
        /// </summary>
        OneNoteStyleConfig StyleConfig { get; }

        /// <summary>
        /// Converts a container of inline elements (like bold, italic, code) into an HTML string
        /// suitable for embedding within a OneNote <T> element's CDATA section.
        /// </summary>
        /// <param name="container">The container of inline elements.</param>
        /// <returns>An HTML-formatted string.</returns>
        string ConvertInlinesToHtml(ContainerInline container);

        /// <summary>
        /// Recursively processes a collection of blocks using the main converter's logic.
        /// This is useful for handlers that contain nested blocks, like ListBlock.
        /// </summary>
        /// <param name="blocks">The collection of blocks to process.</param>
        /// <returns>A collection of converted OneNote XML elements.</returns>
        IEnumerable<XElement> ProcessBlocks(IEnumerable<Block> blocks);
    }
}