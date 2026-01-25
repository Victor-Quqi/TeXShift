namespace TeXShift.Tests.E2E.Models
{
    internal sealed class CliResult
    {
        public string Status { get; set; }
        public string TestName { get; set; }
        public string OutputDirectory { get; set; }
        public OutputFiles Files { get; set; }
        public long DurationMs { get; set; }
        public string Error { get; set; }
    }
}
