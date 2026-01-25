using System.Threading.Tasks;
using System.Xml.Linq;

namespace TeXShift.Core.OneNoteToMarkdown.Abstractions
{
    /// <summary>
    /// Interface for converting OneNote XML fragments to Markdown.
    /// </summary>
    public interface IOneNoteToMarkdownConverter
    {
        /// <summary>
        /// Converts a OneNote XML element (Outline/OEChildren/OE/Page) to Markdown.
        /// </summary>
        Task<ReverseConversionResult> ConvertToMarkdownAsync(XElement oneNoteElement);
    }
}

