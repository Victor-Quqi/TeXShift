using System.Threading.Tasks;
using System.Xml.Linq;

namespace TeXShift.Core
{
    /// <summary>
    /// Interface for writing converted content back to OneNote pages.
    /// </summary>
    public interface IContentWriter
    {
        /// <summary>
        /// Asynchronously replaces content in OneNote based on the read result and converted XML.
        /// </summary>
        /// <param name="readResult">The original read result containing metadata.</param>
        /// <param name="newOutlineXml">The new Outline XML element to insert.</param>
        /// <returns>A task representing the asynchronous write operation.</returns>
        Task ReplaceContentAsync(ReadResult readResult, XElement newOutlineXml);
    }
}
