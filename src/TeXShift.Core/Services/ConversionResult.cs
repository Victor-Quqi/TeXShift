using System;
using TeXShift.Core.OneNote;

namespace TeXShift.Core.Services
{
    /// <summary>
    /// Represents the result of a conversion pipeline run.
    /// </summary>
    public class ConversionResult
    {
        /// <summary>
        /// Whether the pipeline completed successfully.
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// The read result returned by the content reader.
        /// </summary>
        public ReadResult ReadResult { get; set; }

        /// <summary>
        /// The folder where debug artifacts were written.
        /// </summary>
        public string DebugOutputFolder { get; set; }

        /// <summary>
        /// The full path of the exported PDF, if requested.
        /// </summary>
        public string PdfPath { get; set; }

        /// <summary>
        /// The exception captured during execution, if any.
        /// </summary>
        public Exception Error { get; set; }
    }
}
