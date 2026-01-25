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
            command.Options.Add(outputOption);
            command.Options.Add(cleanupOption);

            command.SetAction(async (parseResult, cancellationToken) =>
            {
                var input = parseResult.GetValue(inputOption);
                var markdown = parseResult.GetValue(markdownOption);
                var output = parseResult.GetValue(outputOption);
                var cleanup = parseResult.GetValue(cleanupOption);

                return await ConvertCommand.RunAsync(input, markdown, output, cleanup).ConfigureAwait(false);
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

            var command = new Command("reverse-xml", "Convert OneNote XML to Markdown (best-effort)");
            command.Options.Add(xmlInputOption);
            command.Options.Add(outputOption);
            command.Options.Add(strictOption);

            command.SetAction(async (parseResult, cancellationToken) =>
            {
                var inputXml = parseResult.GetValue(xmlInputOption);
                var output = parseResult.GetValue(outputOption);
                var strict = parseResult.GetValue(strictOption);

                return await ReverseXmlCommand.RunAsync(inputXml, output, strict).ConfigureAwait(false);
            });

            return command;
        }
    }
}
