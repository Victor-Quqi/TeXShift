using System;
using System.CommandLine;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading.Tasks;
using TeXShift.Core.Services;

namespace TeXShift.Tests.E2E
{
    internal sealed class Program
    {
        private const int ExitSuccess = 0;
        private const int ExitConversionFailed = 1;
        private const int ExitOneNoteUnavailable = 2;
        private const int ExitInvalidArguments = 3;
        private const int ExitPdfFailed = 4;

        [STAThread]
        private static int Main(string[] args)
        {
            var inputOption = new Option<FileInfo>("--input", "-i")
            {
                Description = "Markdown file path"
            };

            var markdownOption = new Option<string>("--markdown", "-m")
            {
                Description = "Inline markdown string"
            };

            var outputOption = new Option<DirectoryInfo>("--output", "-o")
            {
                Description = "Output directory (required)",
                Required = true
            };

            var cleanupOption = new Option<bool>("--cleanup")
            {
                Description = "Clean up test page and notebook after conversion (default: true)",
                DefaultValueFactory = _ => true
            };

            var convertCommand = new Command("convert", "Convert markdown and export results");
            convertCommand.Options.Add(inputOption);
            convertCommand.Options.Add(markdownOption);
            convertCommand.Options.Add(outputOption);
            convertCommand.Options.Add(cleanupOption);

            convertCommand.SetAction(async (parseResult, cancellationToken) =>
            {
                var input = parseResult.GetValue(inputOption);
                var markdown = parseResult.GetValue(markdownOption);
                var output = parseResult.GetValue(outputOption);
                var cleanup = parseResult.GetValue(cleanupOption);

                return await RunConvert(input, markdown, output, cleanup).ConfigureAwait(false);
            });

            var rootCommand = new RootCommand("TeXShift E2E Test Runner");
            rootCommand.Subcommands.Add(convertCommand);

            return rootCommand.Parse(args).Invoke();
        }

        private static async Task<int> RunConvert(FileInfo input, string markdown, DirectoryInfo output, bool cleanup)
        {
            if (output == null)
            {
                return EmitArgumentError("输出目录不能为空。", null, null, 0);
            }

            string markdownContent;
            if (input != null && !string.IsNullOrWhiteSpace(markdown))
            {
                return EmitArgumentError("请仅提供 --input 或 --markdown 其一。", null, output, 0);
            }

            if (input != null)
            {
                if (!input.Exists)
                {
                    return EmitArgumentError($"Markdown 文件不存在: {input.FullName}", null, output, 0);
                }
                markdownContent = File.ReadAllText(input.FullName);
            }
            else
            {
                if (string.IsNullOrWhiteSpace(markdown))
                {
                    return EmitArgumentError("请提供 --input 或 --markdown。", null, output, 0);
                }
                markdownContent = markdown;
            }

            string testName = input != null
                ? Path.GetFileNameWithoutExtension(input.Name)
                : $"inline_{DateTime.Now:yyyyMMdd_HHmmss}";

            if (!output.Exists)
            {
                output.Create();
            }

            var stopwatch = Stopwatch.StartNew();
            var startUtc = DateTime.UtcNow;
            TestPageManager pageManager = null;
            ServiceContainer serviceContainer = null;
            string pageId = null;

            try
            {
                pageManager = await TestPageManager.CreateAsync().ConfigureAwait(false);
                serviceContainer = new ServiceContainer();

                pageId = await pageManager.CreateTestPageAsync(testName, markdownContent).ConfigureAwait(false);

                var orchestrator = serviceContainer.CreateConversionOrchestrator(pageManager.OneNoteApp);
                var options = new ConversionOptions
                {
                    WriteDebugFiles = true,
                    ExportPdf = true,
                    OutputDirectory = output.FullName
                };

                var result = await orchestrator.ExecuteAsync(options).ConfigureAwait(false);
                stopwatch.Stop();

                var files = CollectOutputFiles(output, startUtc, result?.PdfPath);

                if (result == null || !result.Success)
                {
                    string errorMessage = result?.Error?.Message ?? "转换失败。";
                    int exitCode = IsPdfExportFailure(result) ? ExitPdfFailed : ExitConversionFailed;
                    EmitError(errorMessage, result?.Error);
                    EmitResult(new CliResult
                    {
                        Status = "failure",
                        TestName = testName,
                        OutputDirectory = output.FullName,
                        Files = files,
                        DurationMs = stopwatch.ElapsedMilliseconds,
                        Error = errorMessage
                    });
                    return exitCode;
                }

                EmitResult(new CliResult
                {
                    Status = "success",
                    TestName = testName,
                    OutputDirectory = output.FullName,
                    Files = files,
                    DurationMs = stopwatch.ElapsedMilliseconds
                });
                return ExitSuccess;
            }
            catch (COMException comEx)
            {
                stopwatch.Stop();
                string message = $"OneNote 不可用: {comEx.Message}";
                EmitError(message, comEx);
                EmitResult(new CliResult
                {
                    Status = "failure",
                    TestName = testName,
                    OutputDirectory = output.FullName,
                    Files = CollectOutputFiles(output, startUtc, null),
                    DurationMs = stopwatch.ElapsedMilliseconds,
                    Error = message
                });
                return ExitOneNoteUnavailable;
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                string message = $"转换异常: {ex.Message}";
                EmitError(message, ex);
                EmitResult(new CliResult
                {
                    Status = "failure",
                    TestName = testName,
                    OutputDirectory = output.FullName,
                    Files = CollectOutputFiles(output, startUtc, null),
                    DurationMs = stopwatch.ElapsedMilliseconds,
                    Error = message
                });
                return ExitConversionFailed;
            }
            finally
            {
                if (cleanup && pageId != null && pageManager != null)
                {
                    try
                    {
                        await pageManager.DeletePageAsync(pageId).ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        EmitError("清理测试页面失败。", ex);
                    }

                    try
                    {
                        await pageManager.CleanupTestResourcesAsync().ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        EmitError("清理测试笔记本失败。", ex);
                    }
                }

                serviceContainer?.Dispose();
                pageManager?.Dispose();
            }
        }

        private static int EmitArgumentError(string message, string testName, DirectoryInfo output, long durationMs)
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

        private static void EmitError(string message, Exception ex)
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

        private static void EmitResult(CliResult result)
        {
            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true
            };

            string payload = JsonSerializer.Serialize(result, options);
            Console.Out.WriteLine(payload);
        }

        private static bool IsPdfExportFailure(ConversionResult result)
        {
            if (result == null || result.Success)
            {
                return false;
            }

            if (result.Error is IOException)
            {
                return true;
            }

            return result.Error != null &&
                   result.Error.Message.IndexOf("PDF export", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static OutputFiles CollectOutputFiles(DirectoryInfo output, DateTime startUtc, string pdfPath)
        {
            if (output == null || !output.Exists)
            {
                return null;
            }

            return new OutputFiles
            {
                Markdown = FindNewestFile(output, "01_Input_Markdown_*.md", startUtc),
                OriginalXml = FindNewestFile(output, "02_Original_XML_*.xml", startUtc),
                ConvertedXml = FindNewestFile(output, "03_Converted_XML_*.xml", startUtc),
                FinalXml = FindNewestFile(output, "04_Final_Page_XML_*.xml", startUtc),
                Pdf = !string.IsNullOrWhiteSpace(pdfPath)
                    ? Path.GetFileName(pdfPath)
                    : FindNewestFile(output, "TeXShift_Export_*.pdf", startUtc)
            };
        }

        private static string FindNewestFile(DirectoryInfo directory, string pattern, DateTime startUtc)
        {
            var candidates = directory.GetFiles(pattern)
                .Where(file => file.LastWriteTimeUtc >= startUtc.AddSeconds(-2))
                .OrderByDescending(file => file.LastWriteTimeUtc)
                .FirstOrDefault();

            return candidates?.Name;
        }

        private sealed class CliResult
        {
            public string Status { get; set; }
            public string TestName { get; set; }
            public string OutputDirectory { get; set; }
            public OutputFiles Files { get; set; }
            public long DurationMs { get; set; }
            public string Error { get; set; }
        }

        private sealed class OutputFiles
        {
            public string Markdown { get; set; }
            public string OriginalXml { get; set; }
            public string ConvertedXml { get; set; }
            public string FinalXml { get; set; }
            public string Pdf { get; set; }
        }
    }
}
