using Markdig.Syntax;
using System.Collections.Generic;
using System.Xml.Linq;

namespace TeXShift.Core.Markdown.Handlers
{
    /// <summary>
    /// Defines the contract for a class that handles a specific type of Markdig Block.
    /// Each handler is responsible for converting one type of block (e.g., HeadingBlock, ListBlock)
    /// into the corresponding OneNote XML structure.
    /// </summary>
    internal interface IBlockHandler
    {
        /// <summary>
        /// Converts a specific Markdig Block into one or more OneNote XML elements (<OE>).
        /// </summary>
        /// <param name="block">The Markdig block to process. This is guaranteed to be of the type the handler supports.</param>
        /// <param name="context">The converter context, providing access to styles, the XML namespace, and methods for recursive conversion.</param>
        /// <returns>A collection of XElement objects representing the converted content.</returns>
        IEnumerable<XElement> Handle(Block block, IMarkdownConverterContext context);
    }
}