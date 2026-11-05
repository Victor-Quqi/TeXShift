using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TeXShift.Core.Configuration;
using TeXShift.Core.Errors;
using TeXShift.Core.OneNote;

namespace TeXShift.Core.Tests
{
    [TestClass]
    public class PasteConversionTests
    {
        private static readonly XNamespace Ns = OneNoteXml.Namespace;

        [TestMethod]
        public void InlinePastePreservesBothSidesAndOtherParagraphs()
        {
            var page = Page(Paragraph("target", Text("left"), Selected(""), Text("right")), Paragraph("keep", Text("**untouched**")));
            string original = page.ToString();
            var result = OneNotePasteInsertion.CreateOutline(page, Converted(Paragraph(null, Text("<b>new</b>"))));
            Assert.AreEqual("left<b>new</b>right", string.Concat(result.Descendants(Ns + "OE").First().Elements(Ns + "T").Select(t => t.Value)));
            Assert.AreEqual("**untouched**", result.Descendants(Ns + "OE").Last().Element(Ns + "T").Value);
            Assert.AreEqual("keep", (string)result.Descendants(Ns + "OE").Last().Attribute("objectID"));
            Assert.AreEqual("right", result.Descendants(Ns + "T").Single(t => t.Attribute("selected") != null).ElementsAfterSelf().Single().Value);
            Assert.AreEqual(original, page.ToString());
        }

        [TestMethod]
        public void ReplacesOnlySelectedTextAcrossParagraphs()
        {
            var page = Page(Paragraph("a", Text("before"), Selected("remove")), Paragraph("b", Selected("remove too"), Text("after")));
            var result = OneNotePasteInsertion.CreateOutline(page, Converted(Paragraph(null, Text("new"))));
            Assert.AreEqual("beforenewafter", string.Concat(result.Descendants(Ns + "T").Select(t => t.Value)));
            Assert.AreEqual(1, result.Descendants(Ns + "OE").Count());
        }

        [TestMethod]
        public void BlockPastePreservesPrefixSuffixAndNativeObjects()
        {
            var image = new XElement(Ns + "OE", new XAttribute("objectID", "image"), new XElement(Ns + "Image", new XElement(Ns + "Data", "bytes")));
            var heading = Paragraph(null, Text("heading"));
            heading.SetAttributeValue("quickStyleIndex", 1);
            var result = OneNotePasteInsertion.CreateOutline(Page(Paragraph("a", Text("before"), Selected(""), Text("after")), image), Converted(heading));
            CollectionAssert.AreEqual(new[] { "before", "heading", "", "after" }, result.Descendants(Ns + "T").Select(t => t.Value).ToArray());
            Assert.AreEqual(image.ToString(), result.Descendants(Ns + "OE").Last().ToString());
            Assert.AreEqual(1, result.Descendants().Attributes("selected").Count());
        }

        [TestMethod]
        public void CanPasteInEmptyOutlineAndTableCell()
        {
            var empty = Page(Paragraph("a", Selected("")));
            Assert.AreEqual("new", string.Concat(OneNotePasteInsertion.CreateOutline(empty, Converted(Paragraph(null, Text("new")))).Descendants(Ns + "T").Select(t => t.Value)));
            var table = new XElement(Ns + "OE", new XElement(Ns + "Table", new XElement(Ns + "Row",
                new XElement(Ns + "Cell", new XElement(Ns + "OEChildren", Paragraph("cell", Selected("")))))));
            var result = OneNotePasteInsertion.CreateOutline(Page(table), Converted(Paragraph(null, Text("new"))));
            Assert.AreEqual("new", string.Concat(result.Descendants(Ns + "Cell").Single().Descendants(Ns + "T").Select(t => t.Value)));
        }

        [TestMethod]
        public void BinaryReadsAreLimitedToRetainedObjectsInTheUpdatedOutline()
        {
            var image = new XElement(Ns + "OE", new XElement(Ns + "Image",
                new XElement(Ns + "CallbackID", new XAttribute("callbackID", "retained"))));
            var page = Page(Paragraph("target", Selected("")), image);
            page.Add(new XElement(Ns + "InkDrawing", new XElement(Ns + "CallbackID", new XAttribute("callbackID", "ink"))));
            page.Add(Converted(new XElement(Ns + "OE", new XElement(Ns + "Image",
                new XElement(Ns + "CallbackID", new XAttribute("callbackID", "other-outline"))))));
            string original = page.ToString();
            var newImage = new XElement(Ns + "OE", new XElement(Ns + "Image", new XElement(Ns + "Data", "new-bytes")));
            var result = OneNotePasteInsertion.CreateOutline(page, Converted(newImage));
            var reads = new List<string>();
            OneNotePasteInsertion.MaterializeBinaryContent(result, id => { reads.Add(id); return "retained-bytes"; });
            CollectionAssert.AreEqual(new[] { "retained" }, reads);
            CollectionAssert.AreEqual(new[] { "new-bytes", "retained-bytes" }, result.Descendants(Ns + "Data").Select(e => e.Value).ToArray());
            Assert.IsFalse(result.Descendants(Ns + "CallbackID").Any());
            Assert.AreEqual(original, page.ToString());
        }

        [TestMethod]
        public void NestedContentSurvivesCaretInsertion()
        {
            var paragraph = Paragraph("a", Text("before"), Selected(""));
            paragraph.Add(new XElement(Ns + "OEChildren", Paragraph("child", Text("nested"))));
            var result = OneNotePasteInsertion.CreateOutline(Page(paragraph), Converted(Paragraph(null, Text("new"))));
            Assert.AreEqual("nested", result.Descendants(Ns + "OE").Single(e => (string)e.Attribute("objectID") == "child").Element(Ns + "T").Value);
        }

        [TestMethod]
        public void MissingBodySelectionDoesNotOverwriteExistingContent()
        {
            var page = Page(Paragraph("a", Text("keep")));
            string original = page.ToString();
            Assert.ThrowsExactly<ContentReadException>(() => OneNotePasteInsertion.CreateOutline(page, Converted(Paragraph(null, Text("new")))));
            Assert.AreEqual(original, page.ToString());
        }

        [TestMethod]
        public void NewPageGetsAnOutlineAndCaret()
        {
            var page = new XElement(Ns + "Page", new XAttribute("ID", "page"));
            var result = OneNotePasteInsertion.CreateOutline(page, Converted(Paragraph(null, Text("new"))));
            Assert.IsNull(result.Attribute("objectID"));
            Assert.AreEqual("new", result.Descendants(Ns + "T").First().Value);
            Assert.AreEqual("all", (string)result.Descendants(Ns + "T").Last().Attribute("selected"));
        }

        [TestMethod]
        public void PasteSettingDefaultsOffAndRoundTrips()
        {
            Assert.IsFalse(AppSettings.CreateDefault().ShowPasteConvertButton);
            var serializer = new DataContractJsonSerializer(typeof(AppSettings));
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes("{}")))
                Assert.IsFalse(((AppSettings)serializer.ReadObject(stream)).ShowPasteConvertButton);
            using (var stream = new MemoryStream())
            {
                serializer.WriteObject(stream, new AppSettings { ShowPasteConvertButton = true });
                stream.Position = 0;
                Assert.IsTrue(((AppSettings)serializer.ReadObject(stream)).ShowPasteConvertButton);
            }
        }

        [TestMethod]
        public void EditCheckIgnoresCaretAndTimestampButDetectsTextAndStyleChanges()
        {
            var original = Converted(Paragraph("a", Text("left"), Selected(""), Text("right")));
            original.SetAttributeValue("lastModifiedTime", "2026-10-31T00:00:00Z");
            var current = Converted(Paragraph("a", Text("leftright")));
            current.SetAttributeValue("lastModifiedTime", "2026-10-31T00:00:01Z");
            Assert.IsTrue(OneNotePasteInsertion.ContentEquals(original, current));
            current.Descendants(Ns + "T").Single().Value = "edited";
            Assert.IsFalse(OneNotePasteInsertion.ContentEquals(original, current));
            current.Descendants(Ns + "T").Single().Value = "leftright";
            current.Descendants(Ns + "OE").Single().SetAttributeValue("style", "font-size:20pt");
            Assert.IsFalse(OneNotePasteInsertion.ContentEquals(original, current));
        }

        private static XElement Text(string value) => new XElement(Ns + "T", new XCData(value));
        private static XElement Selected(string value) => new XElement(Ns + "T", new XAttribute("selected", "all"), new XCData(value));
        private static XElement Paragraph(string id, params XElement[] texts) => new XElement(Ns + "OE", id == null ? null : new XAttribute("objectID", id), texts);
        private static XElement Converted(params XElement[] paragraphs) => new XElement(Ns + "Outline", new XElement(Ns + "OEChildren", paragraphs));
        private static XElement Page(params XElement[] paragraphs) => new XElement(Ns + "Page", new XAttribute("ID", "page"),
            new XElement(Ns + "Outline", new XAttribute("objectID", "outline"), new XElement(Ns + "OEChildren", paragraphs)));
    }
}
