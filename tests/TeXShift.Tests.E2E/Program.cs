using System;
using System.CommandLine;
using System.IO;
using System.Threading;
using TeXShift.Tests.E2E.Commands;

namespace TeXShift.Tests.E2E
{
    internal sealed class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            var rootCommand = new RootCommand("TeXShift E2E Test Runner");

            rootCommand.Subcommands.Add(CreatePageTestCommand());
            rootCommand.Subcommands.Add(CreatePageTestCommand(reverse: true));
            rootCommand.Subcommands.Add(CreatePageTestCommand(paste: true));
            rootCommand.Subcommands.Add(CreatePageTestCommand(replayWrite: true));
            rootCommand.Subcommands.Add(CreateReverseXmlCommand());
            rootCommand.Subcommands.Add(CreateDumpCommand());
            rootCommand.Subcommands.Add(CreateVerifyRestoreCommand());

            // OneNote selection is shared desktop state; concurrent runners cannot isolate it.
            using (var mutex = new Mutex(false, @"Local\TeXShift.E2E"))
            {
                bool acquired;
                try { acquired = mutex.WaitOne(0); }
                catch (AbandonedMutexException) { acquired = true; }
                if (!acquired)
                {
                    Console.Error.WriteLine("Another TeXShift E2E runner is using OneNote.");
                    return CommandHelpers.ExitInvalidArguments;
                }
                try { return rootCommand.Parse(args).Invoke(); }
                finally { mutex.ReleaseMutex(); }
            }
        }

        private static Command CreatePageTestCommand(bool replayWrite = false, bool reverse = false, bool paste = false)
        {
            var inputOption = new Option<FileInfo>("--input", "-i")
            {
                Description = "Markdown file path"
            };

            var markdownOption = new Option<string>("--markdown", "-m")
            {
                Description = "Inline markdown string"
            };

            var pageOption = new Option<FileInfo>("--page")
            {
                Description = "Page to reproduce for page-specific failures: page XML (from dump --page) or a section file exported from OneNote (.one, first page)"
            };

            var caretOption = new Option<string>("--caret")
            {
                Description = "Put the caret in the paragraph containing this text; its whole Outline is converted (Cursor mode). Default: the first paragraph, a new empty one on Markdown pages"
            };

            var selectOption = new Option<string[]>("--select")
            {
                Description = "Select from the paragraph containing the first text through the one containing the second (Selection mode); one text selects a single paragraph",
                Arity = new ArgumentArity(1, 2),
                AllowMultipleArgumentsPerToken = true
            };

            var outlineOption = new Option<string>("--outline")
            {
                Description = "Select the whole text box containing this paragraph text"
            };

            var emptyOutlineOption = new Option<bool>("--empty-outline")
            {
                Description = "Paste into a new empty text box on the disposable test page"
            };
            var verifyUndoOption = new Option<bool>("--verify-undo")
            {
                Description = "Verify that one native Undo restores the original page after a Ribbon paste"
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

            var changesOption = new Option<FileInfo>("--changes")
            {
                Description = "Page-changes XML template. Tokens: {{pageId}}, {{outlineId}}, {{oe:0}} (zero-based OE in the selected outline). Submitted directly to OneNote.",
                Required = true
            };
            var command = replayWrite
                ? new Command("write-xml", "Replay an XML update against a disposable test page")
                : new Command(paste ? "paste-convert" : reverse ? "reverse" : "convert", "Invoke the installed TeXShift Ribbon command and export the resulting page");
            var pasteInputOption = new Option<FileInfo>("--paste-input")
            {
                Description = "Markdown to paste into the prepared page; Ribbon mode temporarily uses the clipboard",
                Required = true
            };
            if (paste)
            {
                command.Options.Add(pasteInputOption);
                command.Options.Add(emptyOutlineOption);
                command.Options.Add(verifyUndoOption);
            }
            var driverOption = new Option<ConversionDriver>("--driver")
            {
                Description = "ribbon: installed add-in and dialogs (activates Ribbon); core: pipeline with optional --settings, without Ribbon interaction",
                DefaultValueFactory = _ => ConversionDriver.Ribbon
            };
            var settingsOption = new Option<FileInfo>("--settings") {
                Description = "AppSettings JSON for --driver core. Does not change the installed add-in settings."
            };
            if (!replayWrite) { command.Options.Add(driverOption); command.Options.Add(settingsOption); }
            if (replayWrite) command.Options.Add(changesOption);
            command.Options.Add(inputOption);
            command.Options.Add(markdownOption);
            command.Options.Add(pageOption);
            command.Options.Add(caretOption);
            command.Options.Add(outlineOption);
            command.Options.Add(selectOption);
            command.Options.Add(outputOption);
            command.Options.Add(cleanupOption);

            command.SetAction(async (parseResult, cancellationToken) =>
            {
                var options = new PageTestOptions
                {
                    Operation = paste ? "paste-convert" : reverse ? "reverse" : "convert",
                    PasteInput = paste ? parseResult.GetValue(pasteInputOption) : null,
                    Driver = replayWrite ? ConversionDriver.Core : parseResult.GetValue(driverOption),
                    Settings = replayWrite ? null : parseResult.GetValue(settingsOption),
                    Input = parseResult.GetValue(inputOption),
                    Markdown = parseResult.GetValue(markdownOption),
                    Page = parseResult.GetValue(pageOption),
                    Changes = replayWrite ? parseResult.GetValue(changesOption) : null,
                    Caret = parseResult.GetValue(caretOption),
                    Outline = parseResult.GetValue(outlineOption),
                    EmptyOutline = paste && parseResult.GetValue(emptyOutlineOption),
                    VerifyUndo = paste && parseResult.GetValue(verifyUndoOption),
                    Select = parseResult.GetValue(selectOption),
                    Output = parseResult.GetValue(outputOption),
                    Cleanup = parseResult.GetValue(cleanupOption)
                };

                return await PageTestCommand.RunAsync(options).ConfigureAwait(false);
            });

            return command;
        }

        private static Command CreateReverseXmlCommand()
        {
            var xmlInputOption = new Option<string>("--input-xml", "-x")
            {
                Description = "OneNote XML file path (e.g., F03_Converted_XML_*.xml, F04_Final_Page_XML_Basic_*.xml). Supports * and ? wildcards (newest match will be used).",
                Required = true
            };

            var outputOption = new Option<DirectoryInfo>("--output", "-o")
            {
                Description = "Output directory (required)",
                Required = true
            };

            var strictOption = new Option<bool>("--strict", "-s")
            {
                Description = "Use strict recognition (only TeXShift-generated formats)",
                DefaultValueFactory = _ => false
            };

            var ignoreMetaOption = new Option<bool>("--ignore-meta")
            {
                Description = "Ignore TeXShift source metadata and exercise rendered OneNote XML parsing",
                DefaultValueFactory = _ => false
            };

            var command = new Command("reverse-xml", "Convert OneNote XML to Markdown (best-effort)");
            command.Options.Add(xmlInputOption);
            command.Options.Add(outputOption);
            command.Options.Add(strictOption);
            command.Options.Add(ignoreMetaOption);

            command.SetAction(async (parseResult, cancellationToken) =>
            {
                var inputXml = parseResult.GetValue(xmlInputOption);
                var output = parseResult.GetValue(outputOption);
                var strict = parseResult.GetValue(strictOption);
                var ignoreMeta = parseResult.GetValue(ignoreMetaOption);

                return await ReverseXmlCommand.RunAsync(inputXml, output, strict, ignoreMeta).ConfigureAwait(false);
            });

            return command;
        }

        private static Command CreateDumpCommand()
        {
            var hierarchyOption = new Option<bool>("--hierarchy")
            {
                Description = "Print the full OneNote page hierarchy XML"
            };

            var pageOption = new Option<string>("--page")
            {
                Description = "OneNote page ID to inspect or export"
            };

            var outputOption = new Option<FileInfo>("--output", "-o")
            {
                Description = "Output file; required for one/pdf. XML defaults to stdout."
            };
            var formatOption = new Option<DumpFormat>("--format")
            {
                Description = "xml: basic page/hierarchy XML; one: native page snapshot; pdf: rendered page",
                DefaultValueFactory = _ => DumpFormat.Xml
            };

            var command = new Command("dump", "Inspect or export OneNote content");
            command.Options.Add(hierarchyOption);
            command.Options.Add(pageOption);
            command.Options.Add(outputOption);
            command.Options.Add(formatOption);

            command.SetAction(async (parseResult, cancellationToken) =>
            {
                var hierarchy = parseResult.GetValue(hierarchyOption);
                var page = parseResult.GetValue(pageOption);
                var output = parseResult.GetValue(outputOption);

                return await DumpCommand.RunAsync(hierarchy, page, output, parseResult.GetValue(formatOption)).ConfigureAwait(false);
            });

            return command;
        }

        private static Command CreateVerifyRestoreCommand()
        {
            var pagesOption = new Option<string>("--pages")
            {
                Description = "Comma-separated OneNote page IDs",
                Required = true
            };

            var command = new Command(
                "verify-restore",
                "Repeatedly closes and relaunches the running OneNote. Run deliberately on a dev machine; never run in CI.");
            command.Options.Add(pagesOption);

            command.SetAction(async (parseResult, cancellationToken) =>
            {
                var pages = parseResult.GetValue(pagesOption);
                return await VerifyRestoreCommand.RunAsync(pages).ConfigureAwait(false);
            });

            return command;
        }
    }
}
