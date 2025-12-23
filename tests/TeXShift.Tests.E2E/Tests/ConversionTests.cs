using System;
using System.Linq;
using System.Xml.Linq;
using TeXShift.Core.Services;
using TeXShift.Tests.E2E.Fixtures;
using TeXShift.Tests.E2E.Helpers;
using Xunit;

namespace TeXShift.Tests.E2E.Tests
{
    [Collection("OneNoteE2E")]
    public sealed class ConversionTests : IDisposable
    {
        private readonly OneNoteFixture _fixture;
        private readonly ServiceContainer _services;

        public ConversionTests(OneNoteFixture fixture)
        {
            _fixture = fixture;
            _services = new ServiceContainer();
        }

        [Fact]
        public void HeadingConversion_WritesHeading()
        {
            const string headingText = "Heading Test";
            var markdown = "# " + headingText;
            var outline = Convert(markdown);
            var pageId = _fixture.CreateTestPage("E2E Heading Conversion");

            _fixture.UpdatePageOutline(pageId, outline);

            var doc = _fixture.GetPageContent(pageId);
            var ns = doc.Root?.Name.Namespace ?? XNamespace.None;

            Assert.Contains(doc.Descendants(ns + "T"), t => t.Value.Contains(headingText));
            Assert.Contains(doc.Descendants(ns + "OE"), oe => oe.Attribute("quickStyleIndex")?.Value == "0");
        }

        [Fact]
        public void CodeBlockConversion_WritesCodeTable()
        {
            const string codeSnippet = "var number = 1;";
            var markdown = "```csharp\n" + codeSnippet + "\n```";
            var outline = Convert(markdown);
            var pageId = _fixture.CreateTestPage("E2E CodeBlock Conversion");

            _fixture.UpdatePageOutline(pageId, outline);

            var doc = _fixture.GetPageContent(pageId);
            var ns = doc.Root?.Name.Namespace ?? XNamespace.None;

            Assert.NotEmpty(doc.Descendants(ns + "Table"));
            Assert.Contains(doc.Descendants(ns + "T"), t => t.Value.Contains("var") && t.Value.Contains("number"));
        }

        [Fact]
        public void TestExampleFile_ConvertsIndentedCodeBlock()
        {
            var markdown = TestDataLoader.LoadMarkdown("indented_code_block_test.md");
            var outline = Convert(markdown);
            var pageId = _fixture.CreateTestPage("E2E Test Example");

            _fixture.UpdatePageOutline(pageId, outline);

            var doc = _fixture.GetPageContent(pageId);
            var ns = doc.Root?.Name.Namespace ?? XNamespace.None;

            Assert.Contains(doc.Descendants(ns + "T"), t => t.Value.Contains("Indented Code Block Test"));
            Assert.Contains(doc.Descendants(ns + "T"), t => t.Value.Contains("Console") || t.Value.Contains("WriteLine"));
        }

        public void Dispose()
        {
            _services.Dispose();
        }

        private XElement Convert(string markdown)
        {
            var converter = _services.CreateMarkdownConverter();
            return converter.ConvertToOneNoteXmlAsync(markdown).GetAwaiter().GetResult();
        }
    }
}
