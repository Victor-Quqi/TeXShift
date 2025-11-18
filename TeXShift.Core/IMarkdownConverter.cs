using System.Threading.Tasks;
using System.Xml.Linq;

namespace TeXShift.Core
{
    /// <summary>
    /// Interface for converting Markdown text to OneNote XML format.
    /// </summary>
    public interface IMarkdownConverter
    {
        /// <summary>
        /// Asynchronously converts Markdown text to a OneNote Outline XML element.
        /// </summary>
        /// <param name="markdown">The Markdown text to convert.</param>
        /// <returns>A task containing the converted OneNote XML element.</returns>
        Task<XElement> ConvertToOneNoteXmlAsync(string markdown);
    }
}
