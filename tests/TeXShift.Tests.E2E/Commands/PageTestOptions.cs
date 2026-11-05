using System.IO;

namespace TeXShift.Tests.E2E.Commands
{
    internal enum ConversionDriver { Ribbon, Core }

    internal sealed class PageTestOptions
    {
        public FileInfo Settings { get; set; }

        public FileInfo Input { get; set; }
        public FileInfo PasteInput { get; set; }

        public string Markdown { get; set; }

        public FileInfo Page { get; set; }

        public FileInfo Changes { get; set; }

        public string Caret { get; set; }
        public string Outline { get; set; }
        public bool EmptyOutline { get; set; }

        public string[] Select { get; set; }

        public string Operation { get; set; } = "convert";
        public ConversionDriver Driver { get; set; } = ConversionDriver.Ribbon;

        public DirectoryInfo Output { get; set; }

        public bool Cleanup { get; set; } = true;
    }
}
