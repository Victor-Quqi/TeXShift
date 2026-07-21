using System;
using System.Reflection;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TeXShift.Core.OneNoteToMarkdown;

namespace TeXShift.Core.Tests
{
    [TestClass]
    public class TeXShiftMetaSignatureTests
    {
        private static readonly XNamespace OneNoteNamespace =
            "http://schemas.microsoft.com/office/onenote/2013/onenote";

        [TestMethod]
        [DataRow("{", "", "<mo>{</mo><mi>x</mi>")]
        [DataRow("", "|", "<mi>x</mi><mo>|</mo>")]
        public void OneSidedMfencedMatchesOneNoteExplicitFenceRewrite(
            string open,
            string close,
            string rewrittenContent)
        {
            var mfenced =
                $"<mfenced open=\"{open}\" close=\"{close}\"><mrow><mi>x</mi></mrow></mfenced>";

            var beforeOneNote = BuildOutline(mfenced);
            var afterOneNote = BuildOutline($"<mrow>{rewrittenContent}</mrow>");

            Assert.AreEqual(
                ComputeSignature(beforeOneNote),
                ComputeSignature(afterOneNote));
        }

        [TestMethod]
        public void MfencedWithoutAttributesUsesDefaultParentheses()
        {
            var beforeOneNote = BuildOutline("<mfenced><mi>x</mi></mfenced>");
            var afterOneNote = BuildOutline("<mrow><mo>(</mo><mi>x</mi><mo>)</mo></mrow>");

            Assert.AreEqual(
                ComputeSignature(beforeOneNote),
                ComputeSignature(afterOneNote));
        }

        private static XElement BuildOutline(string mathMlContent)
        {
            var richText =
                "<!--[if mathML]><math xmlns=\"http://www.w3.org/1998/Math/MathML\">" +
                mathMlContent +
                "</math><![endif]-->";

            return new XElement(
                OneNoteNamespace + "Outline",
                new XElement(
                    OneNoteNamespace + "OEChildren",
                    new XElement(
                        OneNoteNamespace + "OE",
                        new XElement(OneNoteNamespace + "T", new XCData(richText)))));
        }

        private static string ComputeSignature(XElement outline)
        {
            var assembly = typeof(OneNoteToMarkdownConverter).Assembly;
            var writerType = assembly.GetType(
                "TeXShift.Core.OneNoteMeta.TeXShiftMetaWriter",
                throwOnError: true);
            var method = writerType.GetMethod(
                "ComputeSignature",
                BindingFlags.Static | BindingFlags.NonPublic);

            Assert.IsNotNull(method);
            return (string)method.Invoke(null, new object[] { outline });
        }
    }
}
