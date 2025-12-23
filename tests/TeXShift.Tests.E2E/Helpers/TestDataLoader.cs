using System;
using System.IO;
using System.Text;

namespace TeXShift.Tests.E2E.Helpers
{
    internal static class TestDataLoader
    {
        public static string LoadMarkdown(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentException("File name is required.", nameof(fileName));
            }

            var root = FindRepositoryRoot();
            var path = Path.Combine(root, "misc", "test_example", fileName);
            if (!File.Exists(path))
            {
                throw new FileNotFoundException($"Markdown test file not found: {path}", path);
            }

            return File.ReadAllText(path, Encoding.UTF8);
        }

        private static string FindRepositoryRoot()
        {
            var current = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (current != null)
            {
                var candidate = Path.Combine(current.FullName, "misc", "test_example");
                if (Directory.Exists(candidate))
                {
                    return current.FullName;
                }
                current = current.Parent;
            }

            throw new DirectoryNotFoundException("Unable to locate repository root from test base directory.");
        }
    }
}
