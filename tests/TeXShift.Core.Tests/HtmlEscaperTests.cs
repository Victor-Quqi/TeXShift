using Microsoft.VisualStudio.TestTools.UnitTesting;
using TeXShift.Core.Utils;

namespace TeXShift.Core.Tests
{
    [TestClass]
    public class HtmlEscaperTests
    {
        [TestMethod]
        public void EscapeEncodesHtmlSpecialCharacters()
        {
            Assert.AreEqual(
                "&lt;a href=&quot;x&quot;&gt;Tom &amp; Jerry&#39;s&lt;/a&gt;",
                HtmlEscaper.Escape("<a href=\"x\">Tom & Jerry's</a>"));
        }

        [TestMethod]
        public void EscapePreservesNullAndEmptyValues()
        {
            Assert.IsNull(HtmlEscaper.Escape(null));
            Assert.AreEqual(string.Empty, HtmlEscaper.Escape(string.Empty));
        }
    }
}
