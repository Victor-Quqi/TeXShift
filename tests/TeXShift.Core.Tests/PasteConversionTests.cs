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
            Assert.AreEqual("target", (string)result.Descendants(Ns + "OE").First().Attribute("objectID"));
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
            Assert.AreEqual("a", (string)result.Descendants(Ns + "OE").Single().Attribute("objectID"));
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
            Assert.AreEqual("a", (string)result.Descendants(Ns + "OE").First().Attribute("objectID"));
            Assert.AreEqual(1, result.Descendants(Ns + "OE").Count(e => (string)e.Attribute("objectID") == "a"));
        }

        [TestMethod]
        public void CanPasteInEmptyOutlineAndTableCell()
        {
            var empty = Page(Paragraph("a", Selected("")));
            var emptyResult = OneNotePasteInsertion.CreateOutline(empty, Converted(Paragraph(null, Text("new"))));
            Assert.AreEqual("new", string.Concat(emptyResult.Descendants(Ns + "T").Select(t => t.Value)));
            Assert.AreEqual("a", (string)emptyResult.Descendants(Ns + "OE").Single().Attribute("objectID"));
            var table = new XElement(Ns + "OE", new XElement(Ns + "Table", new XElement(Ns + "Row",
                new XElement(Ns + "Cell", new XElement(Ns + "OEChildren", Paragraph("cell", Selected("")))))));
            var result = OneNotePasteInsertion.CreateOutline(Page(table), Converted(Paragraph(null, Text("new"))));
            Assert.AreEqual("new", string.Concat(result.Descendants(Ns + "Cell").Single().Descendants(Ns + "T").Select(t => t.Value)));
            Assert.AreEqual("cell", (string)result.Descendants(Ns + "Cell").Single().Descendants(Ns + "OE").Single().Attribute("objectID"));
        }

        [TestMethod]
        public void BlockPasteIntoEmptyParagraphRetainsItsIdentityOnTheFirstBlock()
        {
            var result = OneNotePasteInsertion.CreateOutline(Page(Paragraph("empty", Selected(""))),
                Converted(Paragraph(null, Text("first")), Paragraph(null, Text("second"))));
            var paragraphs = result.Descendants(Ns + "OE").ToList();
            Assert.AreEqual(3, paragraphs.Count);
            Assert.AreEqual("empty", (string)paragraphs.First().Attribute("objectID"));
            Assert.IsTrue(paragraphs.Skip(1).All(e => e.Attribute("objectID") == null));
            Assert.AreEqual("all", (string)paragraphs.Last().Element(Ns + "T").Attribute("selected"));
        }

        [TestMethod]
        public void BlockPasteAcrossParagraphsRetainsBothSurvivingBoundaryIdentities()
        {
            var page = Page(Paragraph("a", Text("before"), Selected("remove")),
                Paragraph("b", Selected("remove too"), Text("after")));
            var result = OneNotePasteInsertion.CreateOutline(page,
                Converted(Paragraph(null, Text("first")), Paragraph(null, Text("second"))));
            CollectionAssert.AreEqual(new[] { "a", "b" }, result.Descendants(Ns + "OE")
                .Attributes("objectID").Select(a => a.Value).ToArray());
        }

        [TestMethod]
        public void StructuralBlockPasteRetainsLeadingParagraphAndFinalCaret()
        {
            foreach (string block in new[] { "Table", "Image" })
            foreach (string selection in new[] { "", "replace me" })
            {
                var converted = Converted(new XElement(Ns + "OE", new XElement(Ns + block)));
                var result = OneNotePasteInsertion.CreateOutline(Page(Paragraph("target", Selected(selection), Text("suffix"))), converted);
                var paragraphs = result.Element(Ns + "OEChildren").Elements(Ns + "OE").ToList();
                Assert.AreEqual(3, paragraphs.Count);
                Assert.AreEqual("target", (string)paragraphs[0].Attribute("objectID"));
                Assert.AreEqual("", paragraphs[0].Element(Ns + "T").Value);
                Assert.IsNotNull(paragraphs[1].Element(Ns + block));
                Assert.IsTrue(paragraphs.Skip(1).All(e => e.Attribute("objectID") == null));
                Assert.AreEqual("all", (string)paragraphs[2].Element(Ns + "T").Attribute("selected"));
                Assert.AreEqual("suffix", string.Concat(paragraphs[2].Elements(Ns + "T").Select(t => t.Value)));
            }
        }

        [TestMethod]
        public void StructuralBlockPasteUsesExistingPrefixWithoutAddingABlankParagraph()
        {
            var result = OneNotePasteInsertion.CreateOutline(Page(Paragraph("target", Text("prefix"), Selected(""), Text("suffix"))),
                Converted(new XElement(Ns + "OE", new XElement(Ns + "Table"))));
            var paragraphs = result.Element(Ns + "OEChildren").Elements(Ns + "OE").ToList();
            Assert.AreEqual(3, paragraphs.Count);
            Assert.AreEqual("prefix", paragraphs[0].Element(Ns + "T").Value);
            Assert.IsNotNull(paragraphs[1].Element(Ns + "Table"));
            Assert.AreEqual("suffix", string.Concat(paragraphs[2].Elements(Ns + "T").Select(t => t.Value)));
            CollectionAssert.AreEqual(new[] { "target" }, paragraphs.Attributes("objectID").Select(a => a.Value).ToArray());
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
