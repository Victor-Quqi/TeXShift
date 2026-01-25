using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using TeXShift.Core.Services;
using TeXShift.Tests.E2E.Models;

namespace TeXShift.Tests.E2E.Commands
{
    internal static class ConvertCommand
    {
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
            var startUtc = DateTime.UtcNow;
            TestPageManager pageManager = null;
            ServiceContainer serviceContainer = null;
            string pageId = null;

            try
            {
                pageManager = await TestPageManager.CreateAsync().ConfigureAwait(false);
                serviceContainer = new ServiceContainer();

                // Save current page so we can navigate back after test
                await pageManager.SaveCurrentPageAsync().ConfigureAwait(false);

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
                if (cleanup && pageId != null && pageManager != null)
                {
                    try
                    {
                        await pageManager.DeletePageAsync(pageId).ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        CommandHelpers.EmitError("清理测试页面失败。", ex);
                    }

                    try
                    {
                        await pageManager.CleanupTestResourcesAsync().ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        CommandHelpers.EmitError("清理测试笔记本失败。", ex);
                    }
                }

                // Navigate back to original page after cleanup
                if (pageManager != null)
                {
                    try
                    {
                        await pageManager.RestoreOriginalPageAsync().ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        CommandHelpers.EmitError("返回原始页面失败。", ex);
                    }
                }

                serviceContainer?.Dispose();
                pageManager?.Dispose();
            }
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
    }
}
