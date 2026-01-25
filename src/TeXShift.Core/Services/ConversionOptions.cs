namespace TeXShift.Core.Services
{
    /// <summary>
    /// Options for executing the conversion pipeline.
    /// </summary>
    public class ConversionOptions
    {
        /// <summary>
        /// Whether to write debug artifacts to disk.
        /// </summary>
        public bool WriteDebugFiles { get; set; }

        /// <summary>
        /// Whether to export the current page as a PDF after conversion.
        /// </summary>
        public bool ExportPdf { get; set; }

        /// <summary>
        /// Whether to dump the full OneNote page XML (including binary payloads) for debugging.
        /// Default: false (avoid stalls on pages with large base64 content).
        /// </summary>
        public bool DumpFullPageXml { get; set; }

        /// <summary>
        /// Output directory for debug files and optional PDF export.
        /// </summary>
        public string OutputDirectory { get; set; }
    }
}
