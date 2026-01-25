using System.Collections.Generic;

namespace TeXShift.Core.OneNoteToMarkdown
{
    /// <summary>
    /// Result of a OneNote XML -> Markdown reverse conversion.
    /// </summary>
    public sealed class ReverseConversionResult
    {
        /// <summary>
        /// Converted Markdown output.
        /// </summary>
        public string Markdown { get; set; }

        /// <summary>
        /// Best-effort warnings collected during conversion (non-fatal).
        /// </summary>
        public List<string> Warnings { get; } = new List<string>();
    }
}

