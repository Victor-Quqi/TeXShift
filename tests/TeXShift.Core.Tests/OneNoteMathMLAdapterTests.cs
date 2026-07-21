using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TeXShift.Core.Math;

namespace TeXShift.Core.Tests
{
    [TestClass]
    public class OneNoteMathMLAdapterTests
    {
        private const string ConditionalCommentStart = "<!--[if mathML]>";
        private const string TrailingSentinel = "<span lang='x-none'>\u200B</span>";

        [TestMethod]
        [DataRow("<math><mn>2</mn></math>")]
        [DataRow("<math><mi>x</mi><mo>+</mo><mn>1</mn></math>")]
        public void AdaptToOneNoteWrapsNonEmptyMathMlWithoutLeadingSentinel(string mathMl)
        {
            var result = AdaptToOneNote(mathMl);

            // A leading sentinel becomes the first plain run and makes a leading numeric token
            // inherit the body font instead of the math font during OneNote import.
            StringAssert.StartsWith(
                result,
                ConditionalCommentStart,
                "Adapted MathML must start with the conditional comment; a leading sentinel causes OneNote's numeric-token font regression.");
            StringAssert.EndsWith(
                result,
                TrailingSentinel,
                "Adapted MathML must retain the trailing zero-width-space sentinel.");

            var conditionalCommentEnd = result.IndexOf("<![endif]-->");
            Assert.IsTrue(
                conditionalCommentEnd > ConditionalCommentStart.Length,
                "Adapted MathML must contain a non-empty conditional comment.");
            StringAssert.Contains(
                result.Substring(ConditionalCommentStart.Length, conditionalCommentEnd - ConditionalCommentStart.Length),
                mathMl,
                "The input MathML must remain inside the conditional comment.");
        }

        [TestMethod]
        [DataRow(null)]
        [DataRow("")]
        [DataRow("   \t\r\n")]
        public void AdaptToOneNoteReturnsEmptyForNullEmptyOrWhitespace(string mathMl)
        {
            Assert.AreEqual(string.Empty, AdaptToOneNote(mathMl));
        }

        private static string AdaptToOneNote(string mathMl)
        {
            var adapterType = typeof(IMathService).Assembly.GetType(
                "TeXShift.Core.Math.OneNoteMathMLAdapter",
                throwOnError: true);
            var method = adapterType.GetMethod(
                "AdaptToOneNote",
                BindingFlags.Static | BindingFlags.Public);

            Assert.IsNotNull(method);
            return (string)method.Invoke(null, new object[] { mathMl });
        }
    }
}
