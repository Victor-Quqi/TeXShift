using System;
using System.CommandLine;
using System.IO;
using TeXShift.Tests.E2E.Commands;

namespace TeXShift.Tests.E2E
{
    internal sealed class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            var rootCommand = new RootCommand("TeXShift E2E Test Runner");

            rootCommand.Subcommands.Add(CreateConvertCommand());
            rootCommand.Subcommands.Add(CreateReverseXmlCommand());
            rootCommand.Subcommands.Add(CreateDumpCommand());
            rootCommand.Subcommands.Add(CreateVerifyRestoreCommand());

            return rootCommand.Parse(args).Invoke();
        }

        private static Command CreateConvertCommand()
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

            var foregroundOption = new Option<bool>("--foreground")
            {
                Description = "Make OneNote the active window before converting, as when a user clicks the ribbon (takes keyboard focus during the run)"
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

            var command = new Command("convert", "Convert markdown and export results");
            command.Options.Add(inputOption);
            command.Options.Add(markdownOption);
            command.Options.Add(pageOption);
            command.Options.Add(caretOption);
            command.Options.Add(selectOption);
            command.Options.Add(foregroundOption);
            command.Options.Add(outputOption);
            command.Options.Add(cleanupOption);

            command.SetAction(async (parseResult, cancellationToken) =>
            {
                var options = new ConvertCommandOptions
                {
                    Input = parseResult.GetValue(inputOption),
                    Markdown = parseResult.GetValue(markdownOption),
                    Page = parseResult.GetValue(pageOption),
                    Caret = parseResult.GetValue(caretOption),
                    Select = parseResult.GetValue(selectOption),
                    Foreground = parseResult.GetValue(foregroundOption),
                    Output = parseResult.GetValue(outputOption),
                    Cleanup = parseResult.GetValue(cleanupOption)
                };

                return await ConvertCommand.RunAsync(options).ConfigureAwait(false);
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
                Description = "Print basic XML for the specified OneNote page ID"
            };

            var outputOption = new Option<FileInfo>("--output", "-o")
            {
                Description = "Write XML to a UTF-8 file instead of stdout"
            };

            var command = new Command("dump", "Inspect OneNote XML without changing content");
            command.Options.Add(hierarchyOption);
            command.Options.Add(pageOption);
            command.Options.Add(outputOption);

            command.SetAction(async (parseResult, cancellationToken) =>
            {
                var hierarchy = parseResult.GetValue(hierarchyOption);
                var page = parseResult.GetValue(pageOption);
                var output = parseResult.GetValue(outputOption);

                return await DumpCommand.RunAsync(hierarchy, page, output).ConfigureAwait(false);
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
