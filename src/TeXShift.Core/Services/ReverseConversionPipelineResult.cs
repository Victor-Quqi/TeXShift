using System;
using TeXShift.Core.OneNote;
using TeXShift.Core.OneNoteToMarkdown;

namespace TeXShift.Core.Services
{
    /// <summary>
    /// Represents the result of a reverse conversion pipeline run.
    /// </summary>
    public sealed class ReverseConversionPipelineResult
    {
        public bool Success { get; set; }
        public ReadResult ReadResult { get; set; }
        public ReverseConversionResult ReverseResult { get; set; }
        public string DebugOutputFolder { get; set; }
        public Exception Error { get; set; }
    }
}

