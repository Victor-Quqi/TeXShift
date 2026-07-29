using System;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TeXShift.Core.Configuration;
using TeXShift.Core.Markdown;
using TeXShift.Core.OneNote;
using TeXShift.Core.OneNoteToMarkdown;
using TeXShift.Core.Services;

namespace TeXShift.Core.Tests
{
    [TestClass]
    public class TextColorTests
    {
        private static readonly XNamespace OneNoteNamespace = OneNoteXml.Namespace;

        [TestMethod]
        public async Task ForwardConversionNormalizesSupportedCssColorForms()
        {
            const string source =
                "<span style=\"color:#f00\">hex</span> " +
                "<span style='color:rgb(0, 128, 255)'>rgb</span> " +
                "<span style=\"color:hsl(120, 100%, 25%)\">hsl</span> " +
                "<span style=\"color:hsl(0.5turn 100% 50%)\">turn</span> " +
                "<span style=\"color:hsl(3.141592653589793rad 100% 50%)\">rad</span> " +
                "<span style=\"color:rgba(100% 0% 0% / 100%)\">rgba</span> " +
                "<span style=\"color:grey\">named</span>";

            var outline = await ConvertForwardAsync(source);
            string html = GetRichText(outline);

            StringAssert.Contains(html, "<span style='color:#FF0000'>hex</span>");
            StringAssert.Contains(html, "<span style='color:#0080FF'>rgb</span>");
            StringAssert.Contains(html, "<span style='color:#008000'>hsl</span>");
            StringAssert.Contains(html, "<span style='color:#00FFFF'>turn</span>");
            StringAssert.Contains(html, "<span style='color:#00FFFF'>rad</span>");
            StringAssert.Contains(html, "<span style='color:#FF0000'>rgba</span>");
            StringAssert.Contains(html, "<span style='color:#808080'>named</span>");
        }

        [TestMethod]
        public async Task ForwardConversionStripsUnrelatedSpanAttributesAndStyles()
        {
            const string source =
                "<span class='external' onclick='alert(1)' " +
                "style=\"font-weight:bold;color:hsl(240 100% 50%);text-decoration:underline\">blue</span>";

            var outline = await ConvertForwardAsync(source);
            string html = GetRichText(outline);

            Assert.AreEqual("<span style='color:#0000FF'>blue</span>", html);
        }

        [TestMethod]
        public async Task SanitizerStillUnwrapsOneNoteFormattingSpans()
        {
            const string source =
                "<span lang='en-US' style='font-family:Calibri;font-weight:bold'>plain</span>";

            var outline = await ConvertForwardAsync(source);

            Assert.AreEqual("plain", GetRichText(outline));
        }

        [TestMethod]
        public async Task SanitizerPreservesColorInsideOneNoteFormattingSpans()
        {
            const string source =
                "<span lang='en-US' style='font-family:Calibri'>" +
                "before <span class='user' style='font-weight:bold;color:#0f0'>green</span> after" +
                "</span>";

            var outline = await ConvertForwardAsync(source);

            Assert.AreEqual(
                "before <span style='color:#00FF00'>green</span> after",
                GetRichText(outline));
        }

        [TestMethod]
        public async Task TransparentColorsAreRejectedWithoutLosingText()
        {
            const string source =
                "<span style=\"color:rgba(255, 0, 0, 0.5)\">rgba</span> " +
                "<span style=\"color:hsla(0, 100%, 50%, 50%)\">hsla</span>";

            var outline = await ConvertForwardAsync(source);

            Assert.AreEqual("rgba hsla", GetRichText(outline));
        }

        [TestMethod]
        public async Task ColorSpansRemainLiteralInsideCode()
        {
            const string source =
                "`<span style=\"color:hsl(120 100% 25%)\">inline</span>`\n\n" +
                "```html\n<span style=\"color:#f00\">block</span>\n```";

            var outline = await ConvertForwardAsync(source);
            var richText = outline.Descendants(OneNoteNamespace + "T")
                .Select(element => element.Value)
                .ToArray();

            Assert.AreEqual(2, richText.Length);
            Assert.AreEqual(
                "<span style=\"color:hsl(120 100% 25%)\">inline</span>",
                GetVisibleHtmlText(richText[0]));
            Assert.AreEqual(
                "<span style=\"color:#f00\">block</span>",
                GetVisibleHtmlText(richText[1]));
        }

        [TestMethod]
        public async Task ReverseConversionCanonicalizesOneNoteTextColors()
        {
            const string html =
                "<span style='color:#f00'>red</span> " +
                "<span style='color:rgb(0, 128, 255)'>blue</span> " +
                "<span style='color:hsl(120 100% 25%)'>green</span>";

            string markdown = await ConvertReverseAsync(html);

            Assert.AreEqual(
                "<span style=\"color:#FF0000\">red</span> " +
                "<span style=\"color:#0080FF\">blue</span> " +
                "<span style=\"color:#008000\">green</span>",
                markdown);
        }

        [TestMethod]
        public async Task MetadataFreeRoundTripPreservesColorWithMixedInlineContent()
        {
            const string source =
                "<span style=\"color:hsl(4 70% 52%)\">" +
                "**Red [link](https://example.com)** and `code`, " +
                "==highlight==, H~2~O</span>";
            var outline = await ConvertForwardAsync(source);
            RemoveTeXShiftMeta(outline);

            var reverse = new OneNoteToMarkdownConverter(new OneNoteStyleConfig());
            var result = await reverse.ConvertToMarkdownAsync(outline);
            string markdown = result.Markdown.Trim();

            Assert.AreEqual(
                "<span style=\"color:#DA3A2F\">" +
                "**Red [link](https://example.com)** and `code`, " +
                "==highlight==, H~2~O</span>",
                markdown);

            var roundTrip = await ConvertForwardAsync(markdown);
            string html = GetRichText(roundTrip);
            StringAssert.Contains(html, "color:#DA3A2F");
            StringAssert.Contains(html, "font-weight:bold");
            StringAssert.Contains(
                html,
                "<a href=\"https://example.com\" style='color:#DA3A2F'>link</a>");
            StringAssert.Contains(html, "background-color:#FFFF00");
            StringAssert.Contains(html, "<sub>2</sub>");
        }

        [TestMethod]
        public async Task ReverseConversionKeepsOneNoteFlattenedColorRunsParseable()
        {
            const string html =
                "<span style='font-weight:bold;color:#D32F2F'>Review </span>" +
                "<a href='https://example.com/review'>" +
                "<span style='font-weight:bold;color:#D32F2F'>the change</span></a>" +
                "<span style='color:#D32F2F'> and </span>" +
                "<span style='font-family:Consolas;background:#F1F1F1;color:#D32F2F'>" +
                "&nbsp;code&nbsp;</span>";

            string markdown = await ConvertReverseAsync(html);

            Assert.AreEqual(
                "<span style=\"color:#D32F2F\">**Review**</span> " +
                "[<span style=\"color:#D32F2F\">**the change**</span>]" +
                "(https://example.com/review) " +
                "<span style=\"color:#D32F2F\">and `code`</span>",
                markdown);

            var roundTrip = await ConvertForwardAsync(markdown);
            string richText = GetRichText(roundTrip);
            Assert.IsFalse(richText.Contains("style=\"color:"));
            StringAssert.Contains(richText, "color:#D32F2F");
            StringAssert.Contains(richText, "https://example.com/review");
            StringAssert.Contains(richText, "font-family:Consolas");
        }

        [TestMethod]
        public async Task ForwardConversionUsesOeColorForColoredHyperlinks()
        {
            const string source =
                "plain <span style=\"color:#FA0000\">red " +
                "[example](https://example.com)</span> plain";

            var outline = await ConvertForwardAsync(source);
            var oe = outline.Descendants(OneNoteNamespace + "OE").Single();
            string html = GetRichText(outline);

            Assert.AreEqual("color:#FA0000", (string)oe.Attribute("style"));
            StringAssert.StartsWith(html, "<span style='color:#000000'>");
            StringAssert.Contains(
                html,
                "<a href=\"https://example.com\" style='color:#FA0000'>example</a>");
        }

        [TestMethod]
        public async Task ReverseConversionReadsOeLevelHyperlinkColor()
        {
            var page = new XElement(
                OneNoteNamespace + "Page",
                new XElement(
                    OneNoteNamespace + "Outline",
                    new XElement(
                        OneNoteNamespace + "OEChildren",
                        new XElement(
                            OneNoteNamespace + "OE",
                            new XAttribute(
                                "style",
                                "font-family:微软雅黑;font-size:11.0pt;color:#FA0000"),
                            new XElement(
                                OneNoteNamespace + "T",
                                new XCData("<a href=\"https://example.com\">example</a>"))))));

            var converter = new OneNoteToMarkdownConverter(new OneNoteStyleConfig());
            var result = await converter.ConvertToMarkdownAsync(page);

            Assert.AreEqual(
                "<span style=\"color:#FA0000\">[example](https://example.com)</span>",
                result.Markdown.Trim());
        }

        [TestMethod]
        public async Task ReverseConversionReadsQuickStyleHyperlinkColor()
        {
            var page = new XElement(
                OneNoteNamespace + "Page",
                new XElement(
                    OneNoteNamespace + "QuickStyleDef",
                    new XAttribute("index", "7"),
                    new XAttribute("name", "TeXShiftColor7"),
                    new XAttribute("fontColor", "#248F24")),
                new XElement(
                    OneNoteNamespace + "Outline",
                    new XElement(
                        OneNoteNamespace + "OEChildren",
                        new XElement(
                            OneNoteNamespace + "OE",
                            new XAttribute("quickStyleIndex", "7"),
                            new XElement(
                                OneNoteNamespace + "T",
                                new XCData("<a href=\"https://example.com\">example</a>"))))));

            var converter = new OneNoteToMarkdownConverter(new OneNoteStyleConfig());
            var result = await converter.ConvertToMarkdownAsync(page);

            Assert.AreEqual(
                "<span style=\"color:#248F24\">[example](https://example.com)</span>",
                result.Markdown.Trim());
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

        private static string GetRichText(XElement outline)
        {
            return outline.Descendants(OneNoteNamespace + "T").Single().Value;
        }

        private static string GetVisibleHtmlText(string richText)
        {
            string withoutTags = Regex.Replace(richText, "<[^>]+>", string.Empty);
            return WebUtility.HtmlDecode(withoutTags).Trim('\u00A0');
        }

        private static async Task<string> ConvertReverseAsync(string html)
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

            var converter = new OneNoteToMarkdownConverter(new OneNoteStyleConfig());
            var result = await converter.ConvertToMarkdownAsync(page);
            return result.Markdown.Trim();
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
