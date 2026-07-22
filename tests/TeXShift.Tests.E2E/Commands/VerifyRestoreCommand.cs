using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace TeXShift.Tests.E2E.Commands
{
    internal static class VerifyRestoreCommand
    {
        private const int RestorePointerPersistDelayMs = 3000;
        private const string InlineMarkdown = "restore pointer check";

        public static async Task<int> RunAsync(string pages)
        {
            string[] pageIds = (pages ?? string.Empty)
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(pageId => pageId.Trim())
                .Where(pageId => pageId.Length > 0)
                .ToArray();

            if (pageIds.Length == 0)
            {
                Console.Error.WriteLine("At least one page ID is required in --pages.");
                return CommandHelpers.ExitInvalidArguments;
            }

            int passed = 0;

            try
            {
                for (int index = 0; index < pageIds.Length; index++)
                {
                    string targetPageId = pageIds[index];
                    string landedPageId = null;
                    bool cyclePassed = false;

                    try
                    {
                        TestPageManager.CloseAllOneNoteProcesses();
                        await SeedRestorePointerAsync(targetPageId).ConfigureAwait(false);

                        int convertExitCode = await RunConvertAsync().ConfigureAwait(false);
                        landedPageId = await ReadCurrentlyViewedPageIdAsync().ConfigureAwait(false);
                        cyclePassed = convertExitCode == CommandHelpers.ExitSuccess
                            && string.Equals(landedPageId, targetPageId, StringComparison.OrdinalIgnoreCase);
                    }
                    catch (Exception ex)
                    {
                        CommandHelpers.EmitError($"Restore cycle {index + 1} failed: {ex.Message}", ex);
                    }

                    if (cyclePassed)
                    {
                        passed++;
                    }

                    Console.Out.WriteLine($"cycle {index + 1}: target={targetPageId} landed={landedPageId ?? "<none>"} {(cyclePassed ? "PASS" : "FAIL")}");
                }
            }
            finally
            {
                TestPageManager.CloseAllOneNoteProcesses();
            }

            Console.Out.WriteLine($"{passed}/{pageIds.Length} PASS");
            return passed == pageIds.Length
                ? CommandHelpers.ExitSuccess
                : CommandHelpers.ExitConversionFailed;
        }

        private static async Task SeedRestorePointerAsync(string targetPageId)
        {
            TestPageManager pageManager = null;

            try
            {
                pageManager = await TestPageManager.CreateAsync().ConfigureAwait(false);
                await pageManager.NavigateToWhenReadyAsync(targetPageId).ConfigureAwait(false);
                await Task.Delay(RestorePointerPersistDelayMs).ConfigureAwait(false);
            }
            finally
            {
                pageManager?.Dispose();
                TestPageManager.CloseAllOneNoteProcesses();
            }
        }

        private static async Task<int> RunConvertAsync()
        {
            var output = new DirectoryInfo(Path.Combine(
                Path.GetTempPath(),
                "TeXShift_E2E_VerifyRestore_" + Guid.NewGuid().ToString("N")));
            TextWriter originalOutput = Console.Out;

            try
            {
                Console.SetOut(TextWriter.Null);
                return await ConvertCommand.RunAsync(null, InlineMarkdown, output, cleanup: true).ConfigureAwait(false);
            }
            finally
            {
                Console.SetOut(originalOutput);

                if (Directory.Exists(output.FullName))
                {
                    try
                    {
                        output.Delete(recursive: true);
                    }
                    catch (Exception ex)
                    {
                        CommandHelpers.EmitError($"Failed to remove temporary output directory: {output.FullName}", ex);
                    }
                }
            }
        }

        private static async Task<string> ReadCurrentlyViewedPageIdAsync()
        {
            TestPageManager pageManager = null;

            try
            {
                pageManager = await TestPageManager.CreateAsync().ConfigureAwait(false);
                pageManager.CloseLaunchedOneNoteOnDispose();
                try
                {
                    await pageManager.WaitForOneNoteReadyAsync().ConfigureAwait(false);
                }
                catch (TimeoutException)
                {
                    return null;
                }

                string hierarchyXml = await pageManager.GetHierarchyPagesXmlAsync().ConfigureAwait(false);
                var document = XDocument.Parse(hierarchyXml);
                XNamespace oneNoteNamespace = document.Root?.Name.Namespace ?? XNamespace.None;
                var currentPage = document.Descendants(oneNoteNamespace + "Page")
                    .FirstOrDefault(page => string.Equals(
                        (string)page.Attribute("isCurrentlyViewed"),
                        "true",
                        StringComparison.OrdinalIgnoreCase));

                return (string)currentPage?.Attribute("ID");
            }
            finally
            {
                pageManager?.Dispose();
            }
        }
    }
}
