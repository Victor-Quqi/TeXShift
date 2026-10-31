using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using TeXShift.Tests.E2E.Models;

namespace TeXShift.Tests.E2E.Commands
{
    internal static class PageTestCommand
    {
        public static async Task<int> RunAsync(PageTestOptions options)
        {
            FileInfo input = options.Input;
            FileInfo page = options.Page;
            DirectoryInfo output = options.Output;
            if (output == null)
            {
                return CommandHelpers.EmitArgumentError("输出目录不能为空。", null, null, 0);
            }

            TeXShift.Core.Configuration.AppSettings settings = null;
            if (options.Settings != null)
            {
                if (options.Driver != ConversionDriver.Core)
                    return CommandHelpers.EmitArgumentError("--settings 只用于 --driver core。", null, output, 0);
                try
                {
                    using (var stream = options.Settings.OpenRead())
                        settings = (TeXShift.Core.Configuration.AppSettings)new System.Runtime.Serialization.Json.DataContractJsonSerializer(
                            typeof(TeXShift.Core.Configuration.AppSettings)).ReadObject(stream);
                }
                catch (Exception ex) when (ex is IOException || ex is System.Runtime.Serialization.SerializationException)
                {
                    return CommandHelpers.EmitArgumentError("设置文件读取失败: " + ex.Message, null, output, 0);
                }
            }

            bool hasMarkdown = !string.IsNullOrWhiteSpace(options.Markdown);
            int sourceCount = (input != null ? 1 : 0) + (hasMarkdown ? 1 : 0) + (page != null ? 1 : 0);
            if (sourceCount > 1)
            {
                return CommandHelpers.EmitArgumentError("请仅提供 --input、--markdown 或 --page 其一。", null, output, 0);
            }

            if (sourceCount == 0)
            {
                return CommandHelpers.EmitArgumentError("请提供 --input、--markdown 或 --page。", null, output, 0);
            }

            bool hasCaret = !string.IsNullOrEmpty(options.Caret);
            bool hasSelect = options.Select != null && options.Select.Length > 0;
            bool hasOutline = !string.IsNullOrEmpty(options.Outline);
            if ((hasCaret ? 1 : 0) + (hasSelect ? 1 : 0) + (hasOutline ? 1 : 0) > 1)
            {
                return CommandHelpers.EmitArgumentError("请仅提供 --caret、--select 或 --outline 其一。", null, output, 0);
            }

            PageSelection selection = hasOutline ? PageSelection.Outline(options.Outline) : hasCaret
                ? PageSelection.Caret(options.Caret)
                : hasSelect
                    ? PageSelection.Paragraphs(options.Select[0], options.Select[options.Select.Length - 1])
                    : PageSelection.Default;

            string markdownContent = null;
            if (input != null)
            {
                if (!input.Exists)
                {
                    return CommandHelpers.EmitArgumentError($"Markdown 文件不存在: {input.FullName}", null, output, 0);
                }
                markdownContent = File.ReadAllText(input.FullName);
            }
            else if (page != null)
            {
                if (!page.Exists)
                {
                    return CommandHelpers.EmitArgumentError($"页面文件不存在: {page.FullName}", null, output, 0);
                }
            }
            else
            {
                markdownContent = options.Markdown;
            }

            FileInfo sourceFile = input ?? page;
            string pasteMarkdown = null;
            if (options.PasteInput != null)
            {
                if (!options.PasteInput.Exists)
                    return CommandHelpers.EmitArgumentError("粘贴文件不存在。", null, output, 0);
                pasteMarkdown = File.ReadAllText(options.PasteInput.FullName);
                if (string.IsNullOrWhiteSpace(pasteMarkdown))
                    return CommandHelpers.EmitArgumentError("粘贴文件为空。", null, output, 0);
            }
            string changesTemplate = null;
            if (options.Changes != null)
            {
                if (!options.Changes.Exists)
                    return CommandHelpers.EmitArgumentError($"XML 文件不存在: {options.Changes.FullName}", null, output, 0);
                changesTemplate = File.ReadAllText(options.Changes.FullName);
            }
            string testName = sourceFile != null
                ? Path.GetFileNameWithoutExtension(sourceFile.Name)
                : $"inline_{DateTime.Now:yyyyMMdd_HHmmss}";

            if (!output.Exists)
            {
                output.Create();
            }

            var run = new E2ERunReport(testName, options.Cleanup);
            run.Operation = changesTemplate == null ? options.Operation : "write-xml";
            run.Driver = changesTemplate == null ? options.Driver.ToString().ToLowerInvariant() : "xml";
            var cliResult = new CliResult { Operation = run.Operation, Driver = run.Driver, TestName = testName, OutputDirectory = output.FullName };
            int exitCode = CommandHelpers.ExitSuccess;
            TestPageManager pageManager = null;
            RibbonDriver ribbon = null;
            string pdfPath = null;
            System.Windows.Forms.IDataObject originalClipboard = null;
            bool clipboardChanged = false;

            try
            {
                await run.StepAsync("AttachOrLaunchOneNote", async () =>
                {
                    pageManager = await TestPageManager.CreateAsync().ConfigureAwait(false);
                    run.StartMode = pageManager.LaunchedOneNoteProcess ? "launched" : "attached";
                    run.OneNotePid = pageManager.OneNoteProcessId;
                    if (changesTemplate == null && options.Driver == ConversionDriver.Ribbon)
                        ribbon = new RibbonDriver(pageManager.MainWindowHandle);
                }).ConfigureAwait(false);

                await run.StepAsync("SaveCurrentPage", async () =>
                {
                    await pageManager.SaveCurrentPageAsync().ConfigureAwait(false);
                    run.OriginalPageId = pageManager.OriginalPageId;
                }).ConfigureAwait(false);

                await run.StepAsync("CreateTestPage", async () =>
                {
                    run.PageId = page != null
                        ? await pageManager.CreateTestPageFromFileAsync(page.FullName, selection).ConfigureAwait(false)
                        : await pageManager.CreateTestPageAsync(testName, markdownContent, selection).ConfigureAwait(false);
                }).ConfigureAwait(false);

                var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                if (changesTemplate != null)
                {
                    await run.StepAsync("WriteXml", () => XmlWriteReplay.ExecuteAsync(
                        pageManager, run.PageId, changesTemplate, output.FullName)).ConfigureAwait(false);
                }
                else
                {
                    await run.StepAsync("CaptureInput", async () =>
                    {
                        File.WriteAllText(Path.Combine(output.FullName, $"F02_Original_XML_{stamp}.xml"),
                            await pageManager.GetPageSelectionXmlAsync(run.PageId).ConfigureAwait(false));
                        if (markdownContent != null)
                            File.WriteAllText(Path.Combine(output.FullName, $"F01_Input_Markdown_{stamp}.md"), markdownContent);
                    }).ConfigureAwait(false);
                    if (ribbon != null)
                    {
                        await run.StepAsync("PrepareRibbon", () => Task.Run(() =>
                            ribbon.Prepare(options.Operation))).ConfigureAwait(false);
                        if (pasteMarkdown != null)
                            await run.StepAsync("PrepareClipboard", () => pageManager.OneNoteDispatcher.InvokeAsync(app =>
                            {
                                var clipboard = System.Windows.Forms.Clipboard.GetDataObject();
                                if (clipboard != null)
                                {
                                    var snapshot = new System.Windows.Forms.DataObject();
                                    foreach (string format in clipboard.GetFormats(false))
                                    {
                                        object value = clipboard.GetData(format, false);
                                        if (value != null) snapshot.SetData(format, false, value);
                                    }
                                    originalClipboard = snapshot;
                                }
                                System.Windows.Forms.Clipboard.SetText(pasteMarkdown);
                                clipboardChanged = true;
                            })).ConfigureAwait(false);
                        await run.StepAsync("InvokeRibbon", () => ribbon.InvokeAsync(options.Operation, run.PageId, output.FullName)).ConfigureAwait(false);
                    }
                    else
                        await run.StepAsync("InvokeCore", () => CoreConversionDriver.InvokeAsync(
                            options.Operation, pageManager.OneNoteDispatcher, output.FullName, settings, pasteMarkdown)).ConfigureAwait(false);
                    await run.StepAsync("CaptureResult", async () =>
                    {
                        File.WriteAllText(Path.Combine(output.FullName, $"F04_Final_Page_XML_Basic_{stamp}.xml"),
                            await pageManager.GetPageContentBasicXmlAsync(run.PageId).ConfigureAwait(false));
                        File.WriteAllText(Path.Combine(output.FullName, $"F08_Final_Selection_{stamp}.xml"),
                            await pageManager.GetPageSelectionXmlAsync(run.PageId).ConfigureAwait(false));
                    }).ConfigureAwait(false);
                }
                await run.StepAsync("ExportPdf", async () =>
                {
                    pdfPath = Path.Combine(output.FullName, $"F05_Export_{stamp}.pdf");
                    try
                    {
                        if (!await pageManager.ExportPageAsync(run.PageId, pdfPath,
                            Microsoft.Office.Interop.OneNote.PublishFormat.pfPDF).ConfigureAwait(false))
                            throw new IOException("PDF export failed.");
                    }
                    catch { exitCode = CommandHelpers.ExitPdfFailed; throw; }
                }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                if (exitCode == CommandHelpers.ExitSuccess)
                    exitCode = ex is COMException ? CommandHelpers.ExitOneNoteUnavailable : CommandHelpers.ExitConversionFailed;
                cliResult.Error = ex.Message;
                CommandHelpers.EmitError("E2E 执行失败。", ex);
                if (pageManager?.TestPageId != null && ribbon?.CommandPending != true)
                    await run.CleanupAsync("CaptureFailureSelection", async () =>
                        File.WriteAllText(Path.Combine(output.FullName, "F08_Failed_Selection.xml"),
                            await pageManager.GetPageSelectionXmlAsync(pageManager.TestPageId).ConfigureAwait(false))).ConfigureAwait(false);
                if (ribbon != null && !ribbon.CommandPending)
                    await run.CleanupAsync("CaptureFailureUI", () => Task.Run(() =>
                        File.WriteAllText(Path.Combine(output.FullName, "F10_UI.xml"), ribbon.Snapshot().ToString()))).ConfigureAwait(false);
            }
            finally
            {
                if (clipboardChanged && pageManager != null)
                    await run.CleanupAsync("RestoreClipboard", () => pageManager.OneNoteDispatcher.InvokeAsync(app =>
                    {
                        if (System.Windows.Forms.Clipboard.GetText() != pasteMarkdown) return;
                        if (originalClipboard == null) System.Windows.Forms.Clipboard.Clear();
                        else System.Windows.Forms.Clipboard.SetDataObject(originalClipboard, true);
                    })).ConfigureAwait(false);
                if (pageManager != null && ribbon?.CommandPending != true)
                {
                    run.PageId = run.PageId ?? pageManager.TestPageId;
                    bool restored = await run.CleanupAsync("RestoreOriginalPage", pageManager.RestoreOriginalPageAsync).ConfigureAwait(false);
                    if (options.Cleanup && restored)
                        await run.CleanupAsync("DeleteTestResources", pageManager.CleanupTestResourcesAsync).ConfigureAwait(false);
                    else
                        run.Skip("DeleteTestResources", options.Cleanup ? "restore-failed" : "retained");
                }
                if (ribbon?.CommandPending == true)
                {
                    run.Skip("RestoreOriginalPage", "command-pending");
                    run.Skip("DeleteTestResources", "command-pending");
                    await run.CleanupAsync("WaitForCommand", () => throw new TimeoutException(
                        "Command completion is unknown; the test page and OneNote session were retained.")).ConfigureAwait(false);
                }
                else if (ribbon != null)
                    await run.CleanupAsync("RestoreRibbon", () => Task.Run(() => ribbon.Restore())).ConfigureAwait(false);
                if (pageManager != null)
                    await run.CleanupAsync("DisposeOneNote", () => { pageManager.Dispose(); return Task.CompletedTask; }).ConfigureAwait(false);
            }

            exitCode = run.Finish(exitCode);
            cliResult.Status = exitCode == CommandHelpers.ExitSuccess ? "success" : "failure";
            cliResult.CleanupStatus = run.CleanupStatus;
            cliResult.DurationMs = run.DurationMs;
            if (exitCode == CommandHelpers.ExitCleanupFailed)
                cliResult.Error = "转换完成，但 E2E 收尾失败。";
            cliResult.Files = CommandHelpers.CollectOutputFiles(output, run.StartedUtc, pdfPath);
            try { cliResult.DiagnosticsFile = CommandHelpers.WriteJsonFile(output, "F07_E2E_Run", run); }
            catch (Exception ex)
            {
                CommandHelpers.EmitError("写入 E2E 诊断文件失败。", ex);
                cliResult.Status = "failure";
                cliResult.Error = string.Join(" ", cliResult.Error, ex.Message).Trim();
                if (exitCode == CommandHelpers.ExitSuccess) exitCode = CommandHelpers.ExitDiagnosticsFailed;
            }
            CommandHelpers.EmitResult(cliResult);
            return exitCode;
        }

    }
}
