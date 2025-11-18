using System.Threading.Tasks;

namespace TeXShift.Core
{
    /// <summary>
    /// Interface for reading and extracting content from OneNote pages.
    /// </summary>
    public interface IContentReader
    {
        /// <summary>
        /// Asynchronously extracts text content based on the user's current selection or cursor position.
        /// </summary>
        /// <returns>A task containing the read result with extracted content and metadata.</returns>
        Task<ReadResult> ExtractContentAsync();
    }
}
