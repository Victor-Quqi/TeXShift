using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TeXShift.Core.Markdown;
using TeXShift.Core.Math;
using TeXShift.Core.Mermaid;
using TeXShift.Core.OneNote;
using TeXShift.Core.Services;

namespace TeXShift.Core.Tests
{
    [TestClass]
    public class InlineHtmlRenderingTests
    {
        private static readonly XNamespace OneNoteNamespace = OneNoteXml.Namespace;

        [TestMethod]
        public async Task BrTagSpellingsProduceSoftLineBreaks()
        {
            var markdownCases = new[]
            {
                "before<br>after",
                "before<br/>after",
                "before<BR />after",
                "before<br data-kind='soft'>after"
            };

            foreach (var markdown in markdownCases)
            {
                var outline = await ConvertAsync(markdown);
                Assert.AreEqual("before\nafter", GetTextValues(outline).Single(), markdown);
            }
        }

        [TestMethod]
        public async Task ConsecutiveLeadingAndTrailingBrTagsArePreserved()
        {
            var outline = await ConvertAsync("<br>before<br><br/>after<br />");

            Assert.AreEqual("\nbefore\n\nafter\n", GetTextValues(outline).Single());
        }

        [TestMethod]
        public async Task BrTagsComposeWithRichTextAndBlockContainers()
        {
            var markdown = string.Join("\n\n", new[]
            {
                "# heading-a<br>heading-b",
                "**bold-a<br>bold-b**",
                "[link-a<br>link-b](https://example.com)",
                "- list-a<br>list-b",
                "> quote-a<br>quote-b"
            });

            var outline = await ConvertAsync(markdown);
            var textValues = GetTextValues(outline);

            Assert.IsTrue(textValues.Any(value => value.Contains("heading-a\nheading-b")));
            Assert.IsTrue(textValues.Any(value => value.Contains("bold-a\nbold-b")));
            Assert.IsTrue(textValues.Any(value => value.Contains("link-a\nlink-b")));
            Assert.IsTrue(textValues.Any(value => value == "list-a\nlist-b"));
            Assert.IsTrue(textValues.Any(value => value == "quote-a\nquote-b"));
        }

        [TestMethod]
        public async Task BrTagsRenderInsideTableTextAndFormatting()
        {
            var markdown = string.Join("\n", new[]
            {
                "| Plain | Styled |",
                "| --- | --- |",
                "| cell-a<br>cell-b | **bold-a<br>bold-b** |"
            });

            var outline = await ConvertAsync(markdown);
            var textValues = GetTextValues(outline);

            Assert.IsTrue(textValues.Any(value => value == "cell-a\ncell-b"));
            Assert.IsTrue(textValues.Any(value => value.Contains("bold-a\nbold-b")));
        }

        [TestMethod]
        public async Task StandaloneBrHtmlBlockProducesSoftLineBreak()
        {
            var outline = await ConvertAsync("before\n\n<br>\n\nafter");

            CollectionAssert.AreEqual(
                new[] { "before", "\n", "after" },
                GetTextValues(outline).ToArray());
        }

        [TestMethod]
        public async Task EncodedEscapedAndMalformedBrRemainLiteralText()
        {
            var encoded = await ConvertAsync("before&lt;br&gt;after");
            var encodedCustomTag = await ConvertAsync("before&lt;special&gt;after");
            var escaped = await ConvertAsync(@"before\<br>after");
            var malformed = await ConvertAsync("before<br after");

            Assert.AreEqual("before&lt;br&gt;after", GetTextValues(encoded).Single());
            Assert.AreEqual("before&lt;special&gt;after", GetTextValues(encodedCustomTag).Single());
            Assert.AreEqual("before&lt;br&gt;after", GetTextValues(escaped).Single());
            Assert.AreEqual("before&lt;br after", GetTextValues(malformed).Single());
        }

        [TestMethod]
        public async Task SimilarAndClosingHtmlTagsDoNotProduceLineBreaks()
        {
            var similar = await ConvertAsync("before<bracket>after");
            var closing = await ConvertAsync("before</br>after");

            Assert.IsFalse(GetTextValues(similar).Single().Contains("\n"));
            Assert.IsFalse(GetTextValues(closing).Single().Contains("\n"));
        }

        [TestMethod]
        public async Task BrInsideCodeRemainsLiteralText()
        {
            var markdown = string.Join("\n\n", new[]
            {
                "`inline-a<br>inline-b`",
                "```html\nblock-a<br>block-b\n```"
            });

            var outline = await ConvertAsync(markdown);
            var textValues = GetTextValues(outline);

            Assert.AreEqual("inline-a<br>inline-b", GetVisibleHtmlText(textValues[0]));
            Assert.IsFalse(textValues[0].Contains("\n"));
            Assert.AreEqual("block-a<br>block-b", GetVisibleHtmlText(textValues[1]));
            Assert.IsFalse(textValues[1].Contains("\n"));
        }

        [TestMethod]
        public async Task EncodedBrInsideCodeKeepsTheEntitySourceVisible()
        {
            var markdown = string.Join("\n\n", new[]
            {
                "`&lt;br&gt;`",
                "```html\n&lt;br&gt;\n```"
            });

            var outline = await ConvertAsync(markdown);
            var textValues = GetTextValues(outline);

            Assert.AreEqual("&lt;br&gt;", GetVisibleHtmlText(textValues[0]));
            Assert.AreEqual("&lt;br&gt;", GetVisibleHtmlText(textValues[1]));
        }

        [TestMethod]
        public async Task BrInsideMathIsPassedToMathServiceUnchanged()
        {
            var mathService = new RecordingMathService();
            var markdown = "Inline $a<br>b$.\n\n$$\nc<br>d\n$$";

            await ConvertAsync(markdown, mathService: mathService);

            Assert.AreEqual(2, mathService.Calls.Count);
            Assert.AreEqual("a<br>b", mathService.Calls[0].Latex);
            Assert.IsFalse(mathService.Calls[0].DisplayMode);
            Assert.AreEqual("c<br>d", mathService.Calls[1].Latex);
            Assert.IsTrue(mathService.Calls[1].DisplayMode);
        }

        [TestMethod]
        public async Task BrInsideTableMathIsPassedToMathServiceUnchanged()
        {
            var mathService = new RecordingMathService();
            var markdown = string.Join("\n", new[]
            {
                "| Text | Formula |",
                "| --- | --- |",
                "| a<br>b | $c<br>d$ |"
            });

            var outline = await ConvertAsync(markdown, mathService: mathService);

            Assert.IsTrue(GetTextValues(outline).Any(value => value == "a\nb"));
            Assert.AreEqual(1, mathService.Calls.Count);
            Assert.AreEqual("c<br>d", mathService.Calls[0].Latex);
        }

        [TestMethod]
        public async Task BrInsideMermaidIsPassedToMermaidServiceUnchanged()
        {
            var mermaidService = new RecordingMermaidService();
            var markdown = "```mermaid\nflowchart LR\nA[\"a<br>b\"] --> B\n```";

            var outline = await ConvertAsync(markdown, mermaidService: mermaidService);

            Assert.AreEqual("flowchart LR\nA[\"a<br>b\"] --> B", mermaidService.Code);
            Assert.AreEqual(1, outline.Descendants(OneNoteNamespace + "Image").Count());
        }

        private static async Task<XElement> ConvertAsync(
            string markdown,
            IMathService mathService = null,
            IMermaidService mermaidService = null)
        {
            using (var services = new ServiceContainer())
            {
                var converter = new MarkdownToOneNoteConverter(
                    services.StyleConfig,
                    services.MarkdownPipeline,
                    mathService,
                    mermaidService);
                return await converter.ConvertToOneNoteXmlAsync(markdown);
            }
        }

        private static List<string> GetTextValues(XElement outline)
        {
            return outline
                .Descendants(OneNoteNamespace + "T")
                .Select(element => element.Value)
                .ToList();
        }

        private static string GetVisibleHtmlText(string richText)
        {
            var withoutTags = Regex.Replace(richText, "<[^>]+>", string.Empty);
            return WebUtility.HtmlDecode(withoutTags).Trim('\u00A0');
        }

        private sealed class RecordingMathService : IMathService
        {
            public bool IsInitialized => true;
            public List<MathCall> Calls { get; } = new List<MathCall>();

            public Task InitializeAsync()
            {
                return Task.CompletedTask;
            }

            public Task<string> LatexToMathMLAsync(string latex, bool displayMode)
            {
                Calls.Add(new MathCall { Latex = latex, DisplayMode = displayMode });
                return Task.FromResult("<math />");
            }

            public string WrapMathMLForOneNote(string mathml)
            {
                return mathml;
            }

            public void Dispose()
            {
            }
        }

        private sealed class MathCall
        {
            public string Latex { get; set; }
            public bool DisplayMode { get; set; }
        }

        private sealed class RecordingMermaidService : IMermaidService
        {
            public bool IsInitialized => true;
            public string Code { get; private set; }

            public Task InitializeAsync()
            {
                return Task.CompletedTask;
            }

            public Task<MermaidRenderResult> RenderToImageAsync(
                string mermaidCode,
                MermaidRenderOptions options = null)
            {
                Code = mermaidCode;
                return Task.FromResult(new MermaidRenderResult
                {
                    Success = true,
                    Base64PngData = "AQ==",
                    Width = 1,
                    Height = 1
                });
            }

            public void Dispose()
            {
            }
        }
    }
}
