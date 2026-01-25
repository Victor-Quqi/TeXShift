using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using TeXShift.Tests.E2E.Models;

namespace TeXShift.Tests.E2E.Commands
{
    internal static class CommandHelpers
    {
        public const int ExitSuccess = 0;
        public const int ExitConversionFailed = 1;
        public const int ExitOneNoteUnavailable = 2;
        public const int ExitInvalidArguments = 3;
        public const int ExitPdfFailed = 4;
        public const int ExitReverseFailed = 5;

        public static int EmitArgumentError(string message, string testName, DirectoryInfo output, long durationMs)
        {
            EmitError(message, null);
            EmitResult(new CliResult
            {
                Status = "failure",
                TestName = testName,
                OutputDirectory = output?.FullName,
                DurationMs = durationMs,
                Error = message
            });
            return ExitInvalidArguments;
        }

        public static void EmitError(string message, Exception ex)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            Console.Error.WriteLine(message);
            if (ex != null)
            {
                Console.Error.WriteLine(ex);
            }
        }

        public static void EmitResult(CliResult result)
        {
            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true
            };

            string payload = JsonSerializer.Serialize(result, options);
            Console.Out.WriteLine(payload);
        }

        public static OutputFiles CollectOutputFiles(DirectoryInfo output, DateTime startUtc, string pdfPath)
        {
            if (output == null || !output.Exists)
            {
                return null;
            }

            return new OutputFiles
            {
                Markdown = FindNewestFile(output, "F01_Input_Markdown_*.md", startUtc),
                OriginalXml = FindNewestFile(output, "F02_Original_XML_*.xml", startUtc),
                ConvertedXml = FindNewestFile(output, "F03_Converted_XML_*.xml", startUtc),
                FinalXml = FindNewestFile(output, "F04_Final_Page_XML_Basic_*.xml", startUtc),
                FinalXmlFull = FindNewestFile(output, "F04_Final_Page_XML_Full_*.xml", startUtc),
                Perf = FindNewestFile(output, "F06_Perf_*.txt", startUtc),
                Pdf = !string.IsNullOrWhiteSpace(pdfPath)
                    ? Path.GetFileName(pdfPath)
                    : FindNewestFile(output, "F05_Export_*.pdf", startUtc)
            };
        }

        public static string FindNewestFile(DirectoryInfo directory, string pattern, DateTime startUtc)
        {
            var candidates = directory.GetFiles(pattern)
                .Where(file => file.LastWriteTimeUtc >= startUtc.AddSeconds(-2))
                .OrderByDescending(file => file.LastWriteTimeUtc)
                .FirstOrDefault();

            return candidates?.Name;
        }
    }
}
