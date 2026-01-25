using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using TeXShift.Core.Configuration;
using TeXShift.Core.OneNoteToMarkdown;
using TeXShift.Tests.E2E.Models;

namespace TeXShift.Tests.E2E.Commands
{
    internal static class ReverseXmlCommand
    {
        public static async Task<int> RunAsync(string inputXml, DirectoryInfo output, bool strict)
        {
            if (output == null)
            {
                return CommandHelpers.EmitArgumentError("Output directory is required.", null, null, 0);
            }

            var resolvedInputXml = ResolveInputXmlPath(inputXml);
            if (resolvedInputXml == null || !resolvedInputXml.Exists)
            {
                return CommandHelpers.EmitArgumentError($"Input XML file not found: {resolvedInputXml?.FullName ?? inputXml}", null, output, 0);
            }

            if (!output.Exists)
            {
                output.Create();
            }

            var stopwatch = Stopwatch.StartNew();
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string testName = Path.GetFileNameWithoutExtension(resolvedInputXml.Name);

            try
            {
                var xml = File.ReadAllText(resolvedInputXml.FullName);
                var doc = XDocument.Parse(xml);

                var styleConfig = new OneNoteStyleConfig();
                styleConfig.SetReverseConversionOptions(!strict);
                var converter = new OneNoteToMarkdownConverter(styleConfig);
                var result = await converter.ConvertToMarkdownAsync(doc.Root).ConfigureAwait(false);

                string outFile = Path.Combine(output.FullName, $"R02_Reversed_Markdown_{timestamp}.md");
                File.WriteAllText(outFile, result?.Markdown ?? string.Empty, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

                // Output warnings for debugging
                if (result?.Warnings != null && result.Warnings.Count > 0)
                {
                    Console.Error.WriteLine($"[DEBUG] Reverse conversion warnings ({result.Warnings.Count}):");
                    foreach (var warning in result.Warnings)
                    {
                        Console.Error.WriteLine($"  - {warning}");
                    }
                }

                stopwatch.Stop();
                CommandHelpers.EmitResult(new CliResult
                {
                    Status = "success",
                    TestName = testName,
                    OutputDirectory = output.FullName,
                    Files = new OutputFiles { ReversedMarkdown = Path.GetFileName(outFile) },
                    DurationMs = stopwatch.ElapsedMilliseconds
                });
                return CommandHelpers.ExitSuccess;
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                string message = $"Reverse-xml failed: {ex.Message}";
                CommandHelpers.EmitError(message, ex);
                CommandHelpers.EmitResult(new CliResult
                {
                    Status = "failure",
                    TestName = testName,
                    OutputDirectory = output.FullName,
                    DurationMs = stopwatch.ElapsedMilliseconds,
                    Error = message
                });
                return CommandHelpers.ExitReverseFailed;
            }
        }

        private static FileInfo ResolveInputXmlPath(string inputXml)
        {
            if (string.IsNullOrWhiteSpace(inputXml))
            {
                return null;
            }

            // Support wildcards so callers can pass "F04_Final_Page_XML_Basic_*.xml" without relying on shell expansion.
            if (inputXml.IndexOfAny(new[] { '*', '?' }) >= 0)
            {
                try
                {
                    var directory = Path.GetDirectoryName(inputXml);
                    if (string.IsNullOrEmpty(directory))
                    {
                        directory = Directory.GetCurrentDirectory();
                    }

                    var pattern = Path.GetFileName(inputXml);
                    if (string.IsNullOrEmpty(pattern))
                    {
                        return null;
                    }

                    var matches = Directory.GetFiles(directory, pattern, SearchOption.TopDirectoryOnly);
                    if (matches == null || matches.Length == 0)
                    {
                        return null;
                    }

                    var newest = System.Linq.Enumerable.FirstOrDefault(
                        System.Linq.Enumerable.OrderByDescending(
                            System.Linq.Enumerable.Select(matches, p => new FileInfo(p)),
                            f => f.LastWriteTimeUtc));

                    return newest;
                }
                catch
                {
                    // Invalid path/pattern - fall through to FileInfo creation for consistent error reporting.
                }
            }

            try
            {
                return new FileInfo(inputXml);
            }
            catch
            {
                return null;
            }
        }
    }
}
