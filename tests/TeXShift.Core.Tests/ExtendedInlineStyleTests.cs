using System;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TeXShift.Core.Configuration;
using TeXShift.Core.Markdown;
using TeXShift.Core.OneNote;
using TeXShift.Core.OneNoteToMarkdown;
using TeXShift.Core.Services;

namespace TeXShift.Core.Tests
{
    [TestClass]
    public class ExtendedInlineStyleTests
    {
        private static readonly XNamespace OneNoteNamespace = OneNoteXml.Namespace;

        [TestMethod]
        public async Task ForwardConversionRendersExtendedInlineStyles()
        {
            var outline = await ConvertForwardAsync(
                "==marked== ++underlined++ ^super^ ~sub~ ~~strike~~");

            string html = GetRichText(outline);
            StringAssert.Contains(html, "<span style='background-color:#FFFF00'>marked</span>");
            StringAssert.Contains(html, "<span style='text-decoration:underline'>underlined</span>");
            StringAssert.Contains(html, "<sup>super</sup>");
            StringAssert.Contains(html, "<sub>sub</sub>");
            StringAssert.Contains(html, "<span style='text-decoration:line-through'>strike</span>");
        }

        [TestMethod]
        public async Task ForwardConversionRendersInlineHtmlStyleTagsAndAliases()
        {
            const string source =
                "<MARK data-source='external'>marked</MARK> " +
                "<u>under-u</u> <ins class='added'>under-ins</ins> " +
                "<s>strike-s</s> <del cite='history'>strike-del</del> " +
                "<sup title='power'>super</sup> <sub>sub</sub>";

            var outline = await ConvertForwardAsync(source);
            string html = GetRichText(outline);

            StringAssert.Contains(html, "<span style='background-color:#FFFF00'>marked</span>");
            StringAssert.Contains(html, "<span style='text-decoration:underline'>under-u</span>");
            StringAssert.Contains(html, "<span style='text-decoration:underline'>under-ins</span>");
            StringAssert.Contains(html, "<span style='text-decoration:line-through'>strike-s</span>");
            StringAssert.Contains(html, "<span style='text-decoration:line-through'>strike-del</span>");
            StringAssert.Contains(html, "<sup>super</sup>");
            StringAssert.Contains(html, "<sub>sub</sub>");
            Assert.IsFalse(html.Contains("data-source"));
            Assert.IsFalse(html.Contains("class='added'"));
            Assert.IsFalse(html.Contains("cite='history'"));
            Assert.IsFalse(html.Contains("title='power'"));
        }

        [TestMethod]
        public async Task ForwardConversionRendersHtmlEmphasisTagsAndAliases()
        {
            const string source =
                "<STRONG data-source='html'>strong</STRONG> <b>bold</b> " +
                "<em>emphasis</em> <i title='note'>italic</i>";

            var outline = await ConvertForwardAsync(source);
            string html = GetRichText(outline);

            StringAssert.Contains(html, "<span style='font-weight:bold'>strong</span>");
            StringAssert.Contains(html, "<span style='font-weight:bold'>bold</span>");
            StringAssert.Contains(html, "<span style='font-style:italic'>emphasis</span>");
            StringAssert.Contains(html, "<span style='font-style:italic'>italic</span>");
            Assert.IsFalse(html.Contains("data-source"));
            Assert.IsFalse(html.Contains("title='note'"));
        }

        [TestMethod]
        public async Task HtmlEmphasisTagsRoundTripWithNestedContentAndCode()
        {
            const string source =
                "<strong>Bold <em>italic</em> and [link](https://example.com)</strong> " +
                "`<strong>literal</strong>`";
            var outline = await ConvertForwardAsync(source);
            RemoveTeXShiftMeta(outline);

            var reverse = new OneNoteToMarkdownConverter(new OneNoteStyleConfig());
            var result = await reverse.ConvertToMarkdownAsync(outline);

            Assert.AreEqual(
                "**Bold *italic* and [link](https://example.com)** `<strong>literal</strong>`",
                result.Markdown.Trim());

            var roundTrip = await ConvertForwardAsync(result.Markdown);
            string html = GetRichText(roundTrip);
            StringAssert.Contains(html, "font-weight:bold");
            StringAssert.Contains(html, "font-style:italic");
            StringAssert.Contains(html, "&lt;strong&gt;literal&lt;/strong&gt;");
        }

        [TestMethod]
        public async Task InlineHtmlStyleTagsComposeAndRemainLiteralInsideCode()
        {
            const string source =
                "<mark>**Important [link](https://example.com)** and `code`</mark> " +
                "<ins>*new*</ins> <del>old H<sub>2</sub>O</del> " +
                "`<mark>literal</mark>`";

            var outline = await ConvertForwardAsync(source);
            string html = GetRichText(outline);

            StringAssert.Contains(
                html,
                "<span style='background-color:#FFFF00'><span style='font-weight:bold'>Important <a href=\"https://example.com\">link</a></span> and ");
            StringAssert.Contains(
                html,
                "<span style='text-decoration:underline'><span style='font-style:italic'>new</span></span>");
            StringAssert.Contains(
                html,
                "<span style='text-decoration:line-through'>old H<sub>2</sub>O</span>");
            StringAssert.Contains(html, "&lt;mark&gt;literal&lt;/mark&gt;");
        }

        [TestMethod]
        public async Task ReverseConversionParsesExtendedInlineStyles()
        {
            const string html =
                "<span style=\"background-color:yellow\">marked</span> " +
                "<span style=\"text-decoration:underline\">underlined</span> " +
                "<sup>super</sup> <sub>sub</sub>";

            string markdown = await ConvertReverseAsync(html);

            Assert.AreEqual("==marked== ++underlined++ ^super^ ~sub~", markdown);
        }

        [TestMethod]
        public async Task ReverseConversionParsesCombinedDecorationsAndEquivalentForms()
        {
            const string html =
                "<span style=\"text-decoration-line:underline line-through\">both</span> " +
                "<u>under</u> <span style=\"vertical-align:super\">up</span> " +
                "<span style=\"vertical-align:sub\">down</span>";

            string markdown = await ConvertReverseAsync(html);

            Assert.AreEqual("++~~both~~++ ++under++ ^up^ ~down~", markdown);
        }

        [TestMethod]
        public async Task ReverseConversionParsesInlineHtmlStyleAliases()
        {
            const string html =
                "<mark>marked</mark> <u>under-u</u> <ins>under-ins</ins> " +
                "<s>strike-s</s> <del>strike-del</del> " +
                "<sup>super</sup> <sub>sub</sub>";

            string markdown = await ConvertReverseAsync(html);

            Assert.AreEqual(
                "==marked== ++under-u++ ++under-ins++ ~~strike-s~~ ~~strike-del~~ ^super^ ~sub~",
                markdown);
        }

        [TestMethod]
        public async Task MetadataFreeRoundTripCanonicalizesMixedInlineHtmlStyles()
        {
            const string source =
                "<mark>**Important [link](https://example.com)** and `code`</mark> " +
                "<ins>*new*</ins> <del>old H<sub>2</sub>O</del>";
            var outline = await ConvertForwardAsync(source);
            RemoveTeXShiftMeta(outline);

            var reverse = new OneNoteToMarkdownConverter(new OneNoteStyleConfig());
            var result = await reverse.ConvertToMarkdownAsync(outline);

            Assert.AreEqual(
                "==**Important [link](https://example.com)** and `code`== " +
                "++*new*++ ~~old H<sub>2</sub>O~~",
                result.Markdown.Trim());

            var roundTrip = await ConvertForwardAsync(result.Markdown);
            string html = GetRichText(roundTrip);
            StringAssert.Contains(html, "background-color:#FFFF00");
            StringAssert.Contains(html, "text-decoration:underline");
            StringAssert.Contains(html, "text-decoration:line-through");
            StringAssert.Contains(html, "<sub>2</sub>");
        }

        [TestMethod]
        public async Task StrictReverseOnlyTreatsCanonicalBackgroundAsHighlight()
        {
            const string html =
                "<span style=\"background-color:#00FF00\">green</span> " +
                "<span style=\"background-color:rgb(255, 255, 0)\">yellow</span>";
            var strictConfig = new OneNoteStyleConfig();
            strictConfig.SetReverseConversionOptions(false);

            string markdown = await ConvertReverseAsync(html, strictConfig);

            Assert.AreEqual("green ==yellow==", markdown);
        }

        [TestMethod]
        public async Task FuzzyReverseTreatsVisibleBackgroundAsHighlight()
        {
            const string html = "<span style=\"background-color:#00FF00\">green</span>";

            string markdown = await ConvertReverseAsync(html);

            Assert.AreEqual("==green==", markdown);
        }

        [TestMethod]
        public async Task MetadataFreeRoundTripPreservesMixedExtendedStyles()
        {
            const string source =
                "==**Important [link](https://example.com)** and `code`== " +
                "$x_i^2$ ^note^ H~2~O";
            var outline = await ConvertForwardAsync(source);
            RemoveTeXShiftMeta(outline);

            var reverse = new OneNoteToMarkdownConverter(new OneNoteStyleConfig());
            var result = await reverse.ConvertToMarkdownAsync(outline);
            string markdown = result.Markdown.Trim();

            StringAssert.Contains(markdown, "[link](https://example.com)");
            StringAssert.Contains(markdown, "`code`");
            StringAssert.Contains(markdown, "^note^");
            StringAssert.Contains(markdown, "H~2~O");

            using (var services = new ServiceContainer())
            {
                var parsed = Markdig.Markdown.Parse(markdown, services.MarkdownPipeline);
                var delimiters = parsed.Descendants<EmphasisInline>()
                    .Select(node => $"{node.DelimiterChar}:{node.DelimiterCount}")
                    .ToArray();

                CollectionAssert.Contains(delimiters, "=:2");
                CollectionAssert.Contains(delimiters, "*:2");
                CollectionAssert.Contains(delimiters, "^:1");
                CollectionAssert.Contains(delimiters, "~:1");
            }
        }

        [TestMethod]
        public async Task InlineCodeBackgroundIsNotRecoveredAsHighlight()
        {
            var outline = await ConvertForwardAsync("`code`");
            RemoveTeXShiftMeta(outline);

            var reverse = new OneNoteToMarkdownConverter(new OneNoteStyleConfig());
            var result = await reverse.ConvertToMarkdownAsync(outline);

            Assert.AreEqual("`code`", result.Markdown.Trim());
        }

        [TestMethod]
        public async Task ReverseConversionMergesOneNoteSplitInlineCodeRuns()
        {
            const string html =
                "<span style='background:yellow;mso-highlight:yellow'>before " +
                "<span style='font-family:Consolas;background:#F1F1F1'>&nbsp;</span>" +
                "<span style='font-family:Consolas;background:#F1F1F1'>code</span>" +
                "<span style='font-family:Consolas;background:#F1F1F1'>&nbsp;</span>" +
                " after</span>";

            string markdown = await ConvertReverseAsync(html);

            Assert.AreEqual("==before `code` after==", markdown);
        }

        [TestMethod]
        public async Task ReverseConversionKeepsOneNoteFlattenedStyleTransitionsParseable()
        {
            const string html =
                "<span style='font-weight:bold;background:yellow'>Highlighted bold</span>" +
                "<span style='background:yellow'> and</span> | " +
                "<span style='font-style:italic;text-decoration:underline'>Underlined italic</span>" +
                "<span style='text-decoration:underline'> and </span>" +
                "<span style='text-decoration:underline line-through'>old H</span>" +
                "<span style='text-decoration:underline line-through;vertical-align:sub'>2</span>" +
                "<span style='text-decoration:underline line-through'>O</span>";

            string markdown = await ConvertReverseAsync(html);

            Assert.AreEqual(
                "==**Highlighted bold** and== | ++*Underlined italic* and ~~old H<sub>2</sub>O~~++",
                markdown);
        }

        [TestMethod]
        public async Task ReverseConversionKeepsStyledWhitespaceAtDelimiterBoundariesParseable()
        {
            const string html =
                "<span style='font-weight:bold;background:yellow'>Review </span>" +
                "<a href='https://example.com/review'>" +
                "<span style='font-weight:bold;background:yellow'>the change</span></a>" +
                "<span style='background:yellow'> and </span>" +
                "<span style='font-family:Consolas;background:#F1F1F1'>&nbsp;code&nbsp;</span>";

            string markdown = await ConvertReverseAsync(html);

            Assert.AreEqual(
                "==**Review**== [==**the change**==](https://example.com/review) ==and== `code`",
                markdown);

            var roundTrip = await ConvertForwardAsync(markdown);
            string richText = GetRichText(roundTrip);
            StringAssert.Contains(richText, "background-color:#FFFF00'>and</span>");
            Assert.IsFalse(richText.Contains("==and=="));
        }

        [TestMethod]
        public async Task ReverseConversionCollapsesOneNoteFormattingNewlineAfterBr()
        {
            const string html =
                "<span style='background:yellow'>first<br />\nsecond</span>";

            string markdown = await ConvertReverseAsync(html);

            Assert.AreEqual("==first\nsecond==", markdown);
            var roundTrip = await ConvertForwardAsync(markdown);
            string richText = GetRichText(roundTrip);
            StringAssert.Contains(richText, "background-color:#FFFF00");
            StringAssert.Contains(richText, "first\nsecond");
        }

        [TestMethod]
        public async Task StrikeAndSubscriptCombinationProducesParseableMarkdown()
        {
            string markdown = await ConvertReverseAsync(
                "<span style=\"text-decoration:line-through\"><sub>old</sub></span>");

            Assert.AreEqual("~~<sub>old</sub>~~", markdown);

            using (var services = new ServiceContainer())
            {
                var parsed = Markdig.Markdown.Parse(markdown, services.MarkdownPipeline);
                var delimiters = parsed.Descendants<EmphasisInline>()
                    .Select(node => $"{node.DelimiterChar}:{node.DelimiterCount}")
                    .ToArray();

                CollectionAssert.AreEqual(new[] { "~:2" }, delimiters);
                CollectionAssert.AreEqual(
                    new[] { "<sub>", "</sub>" },
                    parsed.Descendants<HtmlInline>().Select(node => node.Tag).ToArray());
            }

            var roundTrip = await ConvertForwardAsync(markdown);
            string html = GetRichText(roundTrip);
            StringAssert.Contains(html, "text-decoration:line-through");
            StringAssert.Contains(html, "<sub>old</sub>");
        }

        private static async Task<XElement> ConvertForwardAsync(string markdown)
        {
            using (var services = new ServiceContainer())
            {
                var converter = new MarkdownToOneNoteConverter(
                    services.StyleConfig,
                    services.MarkdownPipeline,
                    mathService: null,
                    mermaidService: null);
                return await converter.ConvertToOneNoteXmlAsync(markdown);
            }
        }

        private static async Task<string> ConvertReverseAsync(
            string html,
            OneNoteStyleConfig styleConfig = null)
        {
            var page = new XElement(
                OneNoteNamespace + "Page",
                new XElement(
                    OneNoteNamespace + "Outline",
                    new XElement(
                        OneNoteNamespace + "OEChildren",
                        new XElement(
                            OneNoteNamespace + "OE",
                            new XElement(OneNoteNamespace + "T", new XCData(html))))));

            var converter = new OneNoteToMarkdownConverter(styleConfig ?? new OneNoteStyleConfig());
            var result = await converter.ConvertToMarkdownAsync(page);
            return result.Markdown.Trim();
        }

        private static string GetRichText(XElement outline)
        {
            return outline.Descendants(OneNoteNamespace + "T").Single().Value;
        }

        private static void RemoveTeXShiftMeta(XElement outline)
        {
            var meta = outline.Elements(OneNoteNamespace + "Meta")
                .Where(element => ((string)element.Attribute("name") ?? string.Empty)
                    .StartsWith("texshift-", StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var element in meta)
            {
                element.Remove();
            }
        }
    }
}
