using System.Xml.Linq;
using TeXShift.Core.Configuration;
using TeXShift.Core.OneNoteToMarkdown.Inlines;

namespace TeXShift.Core.OneNoteToMarkdown.Abstractions
{
    /// <summary>
    /// Provides shared state and helpers for element handlers during reverse conversion.
    /// </summary>
    internal interface IOneNoteConverterContext
    {
        /// <summary>
        /// OneNote XML namespace.
        /// </summary>
        XNamespace OneNoteNamespace { get; }

        /// <summary>
        /// Style configuration (used for strict matching against configured values).
        /// </summary>
        OneNoteStyleConfig StyleConfig { get; }

        /// <summary>
        /// Current list indentation level (0 = top level).
        /// </summary>
        int CurrentIndentLevel { get; set; }

        /// <summary>
        /// Current index for ordered list items (1-based).
        /// </summary>
        int CurrentListIndex { get; set; }

        /// <summary>
        /// Whether the current list sequence is ordered.
        /// </summary>
        bool CurrentListIsOrdered { get; set; }

        /// <summary>
        /// Parses OneNote rich-text HTML into Markdown inlines.
        /// </summary>
        string ParseInlineHtml(string oneNoteHtml, InlineParseMode mode = InlineParseMode.Default);

        /// <summary>
        /// Converts an OEChildren element to Markdown blocks using the current converter.
        /// </summary>
        string ConvertOeChildrenToMarkdown(XElement oeChildren);

        /// <summary>
        /// Adds a non-fatal warning to the conversion result.
        /// </summary>
        void AddWarning(string warning);
    }
}
