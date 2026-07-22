using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TeXShift.Core.OneNote;
using TeXShift.Core.Services;

namespace TeXShift.Core.Tests
{
    [TestClass]
    public class OrderedListNumberingTests
    {
        private static readonly XNamespace OneNoteNamespace = OneNoteXml.Namespace;

        [TestMethod]
        public async Task InterruptedListRestartsAtWrittenNumber()
        {
            using (var services = new ServiceContainer())
            {
                var converter = services.CreateMarkdownConverter();
                var outline = await converter.ConvertToOneNoteXmlAsync("1. a\n\nparagraph\n\n2. b");
                var numbers = outline.Descendants(OneNoteNamespace + "Number").ToList();

                Assert.AreEqual(2, numbers.Count);
                Assert.IsNull(numbers[0].Attribute("restartNumberingAt"));
                Assert.AreEqual("2", (string)numbers[1].Attribute("restartNumberingAt"));
            }
        }

        [TestMethod]
        public async Task ListStartingAtThreeRestartsFirstItemOnly()
        {
            using (var services = new ServiceContainer())
            {
                var converter = services.CreateMarkdownConverter();
                var outline = await converter.ConvertToOneNoteXmlAsync("3. x\n4. y");
                var numbers = outline.Descendants(OneNoteNamespace + "Number").ToList();

                Assert.AreEqual(2, numbers.Count);
                Assert.AreEqual("3", (string)numbers[0].Attribute("restartNumberingAt"));
                Assert.IsNull(numbers[1].Attribute("restartNumberingAt"));
            }
        }

        [TestMethod]
        public async Task ListStartingAtOneHasNoRestartAttribute()
        {
            using (var services = new ServiceContainer())
            {
                var converter = services.CreateMarkdownConverter();
                var outline = await converter.ConvertToOneNoteXmlAsync("1. a\n2. b");
                var numbers = outline.Descendants(OneNoteNamespace + "Number").ToList();

                Assert.AreEqual(2, numbers.Count);
                Assert.IsFalse(numbers.Any(number => number.Attribute("restartNumberingAt") != null));
            }
        }
    }
}
