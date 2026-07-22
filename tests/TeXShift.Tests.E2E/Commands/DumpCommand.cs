using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace TeXShift.Tests.E2E.Commands
{
    internal static class DumpCommand
    {
        private static readonly Encoding Utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

        public static async Task<int> RunAsync(bool hierarchy, string pageId, FileInfo output)
        {
            bool hasPageId = !string.IsNullOrWhiteSpace(pageId);
            if (hierarchy == hasPageId)
            {
                Console.Error.WriteLine("Specify exactly one of --hierarchy or --page.");
                return CommandHelpers.ExitInvalidArguments;
            }

            TestPageManager pageManager = null;

            try
            {
                pageManager = await TestPageManager.CreateAsync().ConfigureAwait(false);
                string xml = hierarchy
                    ? await pageManager.GetHierarchyPagesXmlAsync().ConfigureAwait(false)
                    : await pageManager.GetPageContentBasicXmlAsync(pageId).ConfigureAwait(false);

                if (output == null)
                {
                    Console.Out.Write(xml);
                }
                else
                {
                    output.Directory?.Create();
                    File.WriteAllText(output.FullName, xml, Utf8NoBom);
                }

                return CommandHelpers.ExitSuccess;
            }
            catch (Exception ex)
            {
                CommandHelpers.EmitError($"Dump failed: {ex.Message}", ex);
                return CommandHelpers.ExitOneNoteUnavailable;
            }
            finally
            {
                if (pageManager != null)
                {
                    pageManager.CloseLaunchedOneNoteOnDispose();
                    pageManager.Dispose();
                }
            }
        }
    }
}
