using System;

namespace TeXShift.Tests.E2E.Models
{
    /// <summary>
    /// Selection set up on the test page before converting, with paragraphs matched by the text they
    /// contain. The reader converts a caret's whole Outline (Cursor mode) and only the paragraphs of a
    /// selected range (Selection mode).
    /// </summary>
    internal sealed class PageSelection
    {
        private PageSelection(string fromText, string toText, bool wholeOutline = false, bool emptyOutline = false)
        {
            FromText = fromText;
            ToText = toText;
            WholeOutline = wholeOutline;
            IsEmptyOutline = emptyOutline;
        }

        /// <summary>
        /// Caret in the first paragraph of the first Outline: a new empty paragraph on Markdown pages.
        /// </summary>
        public static PageSelection Default { get; } = new PageSelection(null, null);
        public static PageSelection EmptyOutline { get; } = new PageSelection(null, null, emptyOutline: true);

        public string FromText { get; }

        public string ToText { get; }

        public bool IsDefault => FromText == null && !IsEmptyOutline;
        public bool IsEmptyOutline { get; }

        public bool IsRange => ToText != null;
        public bool WholeOutline { get; }

        public static PageSelection Outline(string paragraphText) => new PageSelection(RequireText(paragraphText), null, true);

        public static PageSelection Caret(string paragraphText)
        {
            return new PageSelection(RequireText(paragraphText), null);
        }

        public static PageSelection Paragraphs(string fromText, string toText)
        {
            return new PageSelection(RequireText(fromText), RequireText(toText));
        }

        private static string RequireText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                throw new ArgumentException("Paragraph text is required.", nameof(text));
            }

            return text;
        }
    }
}
