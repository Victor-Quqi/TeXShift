using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using TeXShift.Core.Services;
using TeXShift.Tests.E2E.Models;

namespace TeXShift.Tests.E2E.Commands
{
    internal static class ConvertCommand
    {
        private static readonly Encoding Utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

        public static async Task<int> RunAsync(FileInfo input, string markdown, DirectoryInfo output, bool cleanup)
        {
            if (output == null)
            {
                return CommandHelpers.EmitArgumentError("输出目录不能为空。", null, null, 0);
            }

            string markdownContent;
            if (input != null && !string.IsNullOrWhiteSpace(markdown))
            {
                return CommandHelpers.EmitArgumentError("请仅提供 --input 或 --markdown 其一。", null, output, 0);
            }

            if (input != null)
            {
                if (!input.Exists)
                {
                    return CommandHelpers.EmitArgumentError($"Markdown 文件不存在: {input.FullName}", null, output, 0);
                }
                markdownContent = File.ReadAllText(input.FullName);
            }
            else
            {
                if (string.IsNullOrWhiteSpace(markdown))
                {
                    return CommandHelpers.EmitArgumentError("请提供 --input 或 --markdown。", null, output, 0);
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
            var lifecycleStopwatch = Stopwatch.StartNew();
            var lifecycleEntries = new List<LifecycleEntry>();
            var startUtc = DateTime.UtcNow;
            TestPageManager pageManager = null;
            ServiceContainer serviceContainer = null;
            string pageId = null;

            try
            {
                var stepStopwatch = Stopwatch.StartNew();
                try
                {
                    pageManager = await TestPageManager.CreateAsync().ConfigureAwait(false);
                }
                finally
                {
                    AddLifecycleEntry(
                        lifecycleEntries,
                        "E2E.AttachOrLaunchOneNote",
                        stepStopwatch,
                        pageManager == null ? string.Empty : pageManager.LaunchedOneNoteProcess ? "launched" : "attached");
                }

                serviceContainer = new ServiceContainer();

                // Save current page so we can navigate back after test
                stepStopwatch = Stopwatch.StartNew();
                try
                {
                    await pageManager.SaveCurrentPageAsync().ConfigureAwait(false);
                }
                finally
                {
                    AddLifecycleEntry(lifecycleEntries, "E2E.SaveCurrentPage", stepStopwatch);
                }

                stepStopwatch = Stopwatch.StartNew();
                try
                {
                    pageId = await pageManager.CreateTestPageAsync(testName, markdownContent).ConfigureAwait(false);
                }
                finally
                {
                    AddLifecycleEntry(lifecycleEntries, "E2E.CreateTestPage", stepStopwatch);
                }

                var orchestrator = serviceContainer.CreateConversionOrchestrator(pageManager.OneNoteApp);
                var options = new ConversionOptions
                {
                    WriteDebugFiles = true,
                    ExportPdf = true,
                    OutputDirectory = output.FullName
                };

                ConversionResult result;
                stepStopwatch = Stopwatch.StartNew();
                try
                {
                    result = await orchestrator.ExecuteAsync(options).ConfigureAwait(false);
                }
                finally
                {
                    AddLifecycleEntry(
                        lifecycleEntries,
                        "E2E.ConversionSession",
                        stepStopwatch,
                        "orchestrator ExecuteAsync");
                }

                stopwatch.Stop();

                var files = CommandHelpers.CollectOutputFiles(output, startUtc, result?.PdfPath);

                if (result == null || !result.Success)
                {
                    string errorMessage = result?.Error?.Message ?? "转换失败。";
                    int exitCode = IsPdfExportFailure(result) ? CommandHelpers.ExitPdfFailed : CommandHelpers.ExitConversionFailed;
                    CommandHelpers.EmitError(errorMessage, result?.Error);
                    CommandHelpers.EmitResult(new CliResult
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

                CommandHelpers.EmitResult(new CliResult
                {
                    Status = "success",
                    TestName = testName,
                    OutputDirectory = output.FullName,
                    Files = files,
                    DurationMs = stopwatch.ElapsedMilliseconds
                });
                return CommandHelpers.ExitSuccess;
            }
            catch (COMException comEx)
            {
                stopwatch.Stop();
                string message = $"OneNote 不可用: {comEx.Message}";
                CommandHelpers.EmitError(message, comEx);
                CommandHelpers.EmitResult(new CliResult
                {
                    Status = "failure",
                    TestName = testName,
                    OutputDirectory = output.FullName,
                    Files = CommandHelpers.CollectOutputFiles(output, startUtc, null),
                    DurationMs = stopwatch.ElapsedMilliseconds,
                    Error = message
                });
                return CommandHelpers.ExitOneNoteUnavailable;
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                string message = $"转换异常: {ex.Message}";
                CommandHelpers.EmitError(message, ex);
                CommandHelpers.EmitResult(new CliResult
                {
                    Status = "failure",
                    TestName = testName,
                    OutputDirectory = output.FullName,
                    Files = CommandHelpers.CollectOutputFiles(output, startUtc, null),
                    DurationMs = stopwatch.ElapsedMilliseconds,
                    Error = message
                });
                return CommandHelpers.ExitConversionFailed;
            }
            finally
            {
                try
                {
                    // Restore first so the view and restore pointer never sit on the Quick Notes fallback while artifacts are deleted.
                    if (pageManager != null)
                    {
                        var stepStopwatch = Stopwatch.StartNew();
                        try
                        {
                            await pageManager.RestoreOriginalPageAsync().ConfigureAwait(false);
                        }
                        catch (Exception ex)
                        {
                            CommandHelpers.EmitError("返回原始页面失败。", ex);
                        }
                        finally
                        {
                            AddLifecycleEntry(lifecycleEntries, "E2E.RestoreOriginalPage", stepStopwatch);
                        }
                    }

                    if (cleanup && pageId != null && pageManager != null)
                    {
                        var stepStopwatch = Stopwatch.StartNew();
                        try
                        {
                            await pageManager.DeletePageAsync(pageId).ConfigureAwait(false);
                        }
                        catch (Exception ex)
                        {
                            CommandHelpers.EmitError("清理测试页面失败。", ex);
                        }
                        finally
                        {
                            AddLifecycleEntry(lifecycleEntries, "E2E.DeleteTestPage", stepStopwatch);
                        }

                        stepStopwatch = Stopwatch.StartNew();
                        try
                        {
                            await pageManager.CleanupTestResourcesAsync().ConfigureAwait(false);
                        }
                        catch (Exception ex)
                        {
                            CommandHelpers.EmitError("清理测试笔记本失败。", ex);
                        }
                        finally
                        {
                            AddLifecycleEntry(lifecycleEntries, "E2E.CleanupTestResources", stepStopwatch);
                        }
                    }

                    serviceContainer?.Dispose();
                    pageManager?.Dispose();
                }
                finally
                {
                    lifecycleStopwatch.Stop();
                    TryAppendLifecycleReport(output, startUtc, lifecycleStopwatch.ElapsedMilliseconds, lifecycleEntries);
                }
            }
        }

        private static void AddLifecycleEntry(
            ICollection<LifecycleEntry> entries,
            string step,
            Stopwatch stopwatch,
            string detail = "")
        {
            stopwatch.Stop();
            entries.Add(new LifecycleEntry
            {
                Step = step,
                DurationMs = stopwatch.ElapsedMilliseconds,
                Detail = detail ?? string.Empty
            });
        }

        private static void TryAppendLifecycleReport(
            DirectoryInfo output,
            DateTime startUtc,
            long totalMs,
            IEnumerable<LifecycleEntry> entries)
        {
            try
            {
                string perfFileName = CommandHelpers.FindNewestFile(output, "F06_Perf_*.txt", startUtc);
                string section = FormatLifecycleSection(totalMs, entries);

                if (string.IsNullOrWhiteSpace(perfFileName))
                {
                    string newPerfPath = Path.Combine(output.FullName, $"F06_Perf_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                    File.WriteAllText(newPerfPath, section, Utf8NoBom);
                    return;
                }

                string perfPath = Path.Combine(output.FullName, perfFileName);
                string existingReport = File.ReadAllText(perfPath);
                string separator = GetSectionSeparator(existingReport);
                File.AppendAllText(perfPath, separator + section, Utf8NoBom);
            }
            catch (Exception ex)
            {
                CommandHelpers.EmitError("追加 E2E 生命周期性能数据失败。", ex);
            }
        }

        private static string FormatLifecycleSection(long totalMs, IEnumerable<LifecycleEntry> entries)
        {
            var builder = new StringBuilder();
            builder.AppendLine("E2E Lifecycle");
            builder.AppendLine($"TotalMs\t{totalMs}");
            builder.AppendLine();
            builder.AppendLine("Depth\tStep\tDurationMs\tDetail");

            foreach (var entry in entries)
            {
                builder.Append("0\t");
                builder.Append(entry.Step);
                builder.Append('\t');
                builder.Append(entry.DurationMs);
                builder.Append('\t');
                builder.AppendLine(entry.Detail);
            }

            return builder.ToString();
        }

        private static string GetSectionSeparator(string existingReport)
        {
            if (string.IsNullOrEmpty(existingReport))
            {
                return string.Empty;
            }

            if (existingReport.EndsWith("\r\n\r\n", StringComparison.Ordinal) ||
                existingReport.EndsWith("\n\n", StringComparison.Ordinal))
            {
                return string.Empty;
            }

            if (existingReport.EndsWith("\r\n", StringComparison.Ordinal) ||
                existingReport.EndsWith("\n", StringComparison.Ordinal))
            {
                return Environment.NewLine;
            }

            return Environment.NewLine + Environment.NewLine;
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

        private sealed class LifecycleEntry
        {
            public string Step { get; set; }
            public long DurationMs { get; set; }
            public string Detail { get; set; }
        }
    }
}
