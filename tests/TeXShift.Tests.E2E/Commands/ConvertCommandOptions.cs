using System.IO;

namespace TeXShift.Tests.E2E.Commands
{
    internal sealed class ConvertCommandOptions
    {
        public FileInfo Input { get; set; }

        public string Markdown { get; set; }

        public FileInfo Page { get; set; }

        public string Caret { get; set; }

        public string[] Select { get; set; }

        public bool Foreground { get; set; }

        public DirectoryInfo Output { get; set; }

        public bool Cleanup { get; set; } = true;
    }
}
