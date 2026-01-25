using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using TeXShift.Core.Localization;
using TeXShift.Core.OneNoteToMarkdown.Abstractions;
using TeXShift.Core.OneNoteToMarkdown.Inlines;

namespace TeXShift.Core.OneNoteToMarkdown.Handlers
{
    /// <summary>
    /// Detects OneNote MathML blocks and emits a warning + placeholder.
    ///
    /// TeXShift does not embed the original LaTeX source into the MathML (OneNote strips it anyway),
    /// so without valid TeXShift meta we cannot reliably recover the original expression.
    /// </summary>
    internal sealed class MathElementHandler : IElementHandler
    {
        private static readonly Regex MathConditionalCommentRegex = new Regex(
            "<!--\\s*\\[if\\s+mathml\\s*\\]>.*?<!\\s*\\[endif\\s*\\]\\s*-->",
            RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

        public bool CanHandle(XElement element, IOneNoteConverterContext context)
        {
            if (element == null)
            {
                return false;
            }

            var ns = context.OneNoteNamespace;
            var tElements = element.Elements(ns + "T").ToList();
            if (tElements.Count == 0)
            {
                return false;
            }

            string html = string.Concat(tElements.Select(t => t.Value ?? string.Empty));
            if (html.IndexOf("mathML", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return false;
            }

            if (!MathConditionalCommentRegex.IsMatch(html))
            {
                return false;
            }

            // Only handle "pure" math OEs here. Inline math mixed with other text is handled
            // by InlineParser (it emits an inline placeholder and preserves surrounding text).
            var withoutMath = MathConditionalCommentRegex.Replace(html, string.Empty);
            var stripped = HtmlStripper.Strip(withoutMath)
                .Replace("\u200B", string.Empty) // OneNote sentinel
                .Trim();

            return stripped.Length == 0;
        }

        public IEnumerable<string> Handle(XElement element, IOneNoteConverterContext context)
        {
            context.AddWarning("TeXShift: MathML detected but original LaTeX source cannot be recovered without TeXShift meta; emitting placeholder.");
            yield return Resources.GetString("Reverse_MathSourceMissing");
        }
    }
}

