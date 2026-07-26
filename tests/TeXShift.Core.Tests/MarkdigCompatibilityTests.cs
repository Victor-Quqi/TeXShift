using System.Linq;
using Markdig;
using Markdig.Extensions.Mathematics;
using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TeXShift.Core.Services;

namespace TeXShift.Core.Tests
{
    [TestClass]
    public class MarkdigCompatibilityTests
    {
        [TestMethod]
        public void ConfiguredPipelineParsesExtendedEmphasisNodes()
        {
            using (var services = new ServiceContainer())
            {
                var document = Markdig.Markdown.Parse(
                    "==marked== ++inserted++ ^superscript^ ~subscript~ ~~strikethrough~~",
                    services.MarkdownPipeline);

                var delimiters = document
                    .Descendants<EmphasisInline>()
                    .Select(emphasis => $"{emphasis.DelimiterChar}:{emphasis.DelimiterCount}")
                    .ToArray();

                CollectionAssert.AreEqual(
                    new[] { "=:2", "+:2", "^:1", "~:1", "~:2" },
                    delimiters);
            }
        }

        [TestMethod]
        public void MathematicsTakesPrecedenceOverEmphasisExtras()
        {
            using (var services = new ServiceContainer())
            {
                var document = Markdig.Markdown.Parse(
                    "Inline $x_{total}^2 + y_i$ and ^text^.",
                    services.MarkdownPipeline);

                var math = document.Descendants<MathInline>().Single();
                var emphasis = document.Descendants<EmphasisInline>().Single();

                Assert.AreEqual("x_{total}^2 + y_i", math.Content.ToString());
                Assert.AreEqual('^', emphasis.DelimiterChar);
                Assert.AreEqual(1, emphasis.DelimiterCount);
            }
        }

        [TestMethod]
        public void HtmlCodeAndMathKeepSeparateInlineNodes()
        {
            using (var services = new ServiceContainer())
            {
                var document = Markdig.Markdown.Parse(
                    "before<br>after `x<br>y` $a<br>b$",
                    services.MarkdownPipeline);

                Assert.AreEqual("<br>", document.Descendants<HtmlInline>().Single().Tag);
                Assert.AreEqual("x<br>y", document.Descendants<CodeInline>().Single().Content);
                Assert.AreEqual("a<br>b", document.Descendants<MathInline>().Single().Content.ToString());
            }
        }

        [TestMethod]
        public void BoldApproximatePricesDoNotCreatePhantomTableCells()
        {
            const string markdown =
                "| Component | Per query | Per 1,000 queries |\n" +
                "| --- | --- | --- |\n" +
                "| Embedding | ~$0.00001 | ~$0.01 |\n" +
                "| **Total** | **~$0.0015** | **~$1.50** |";

            using (var services = new ServiceContainer())
            {
                var document = Markdig.Markdown.Parse(markdown, services.MarkdownPipeline);
                var rows = document
                    .Descendants<TableRow>()
                    .ToList();

                Assert.AreEqual(3, rows.Count);
                Assert.IsTrue(rows.All(row => row.Count == 3));
            }
        }
    }
}
