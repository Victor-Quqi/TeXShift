using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TeXShift.Core.Configuration;
using TeXShift.Core.OneNoteToMarkdown;
using TeXShift.Core.Services;

namespace TeXShift.Core.Tests
{
    [TestClass]
    public class ConversionSmokeTests
    {
        private static readonly XNamespace OneNoteNamespace =
            "http://schemas.microsoft.com/office/onenote/2013/onenote";

        [TestMethod]
        public async Task ForwardConversionProducesOneNoteOutline()
        {
            using (var services = new ServiceContainer())
            {
                var converter = services.CreateMarkdownConverter();
                var outline = await converter.ConvertToOneNoteXmlAsync("# Heading\n\nBody");

                Assert.AreEqual(OneNoteNamespace + "Outline", outline.Name);
                Assert.IsTrue(outline.Descendants(OneNoteNamespace + "T").Any());
                StringAssert.Contains(outline.ToString(SaveOptions.DisableFormatting), "Heading");
                StringAssert.Contains(outline.ToString(SaveOptions.DisableFormatting), "Body");
            }
        }

        [TestMethod]
        public async Task ReverseConversionParsesInlineFormatting()
        {
            var page = new XElement(
                OneNoteNamespace + "Page",
                new XElement(
                    OneNoteNamespace + "Outline",
                    new XElement(
                        OneNoteNamespace + "OEChildren",
                        new XElement(
                            OneNoteNamespace + "OE",
                            new XElement(
                                OneNoteNamespace + "T",
                                "Hello <span style=\"font-weight:bold\">world</span>")))));

            var converter = new OneNoteToMarkdownConverter(new OneNoteStyleConfig());
            var result = await converter.ConvertToMarkdownAsync(page);

            Assert.AreEqual("Hello **world**", result.Markdown.Trim());
            Assert.AreEqual(0, result.Warnings.Count);
        }
    }
}
