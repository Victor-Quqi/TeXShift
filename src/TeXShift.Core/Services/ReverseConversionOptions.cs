namespace TeXShift.Core.Services
{
    /// <summary>
    /// Options for executing the reverse conversion pipeline (OneNote -> Markdown).
    /// </summary>
    public sealed class ReverseConversionOptions
    {
        /// <summary>
        /// Whether to write debug artifacts to disk.
        /// </summary>
        public bool WriteDebugFiles { get; set; }

        /// <summary>
        /// Whether to dump the full OneNote page XML (including binary payloads) for debugging.
        /// Default: false (avoid stalls on pages with large base64 content).
        /// </summary>
        public bool DumpFullPageXml { get; set; }

        /// <summary>
        /// Output directory for debug files (optional).
        /// </summary>
        public string OutputDirectory { get; set; }
    }
}

