using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using TeXShift.Core.Configuration;

namespace TeXShift.Core.OneNoteToMarkdown.Inlines
{
    /// <summary>
    /// OneNote rich-text HTML -> Markdown inline parser.
    /// </summary>
    internal sealed partial class InlineParser
    {
        private readonly string _inlineCodeFontFamily;
        private readonly string _inlineCodeBackgroundColor;
        private readonly int _inlineCodePaddingCount;
        private readonly bool _tryRecognizeNonTeXShiftFormats;

        private static readonly Regex HrefAttrRegex = new Regex(
            "\\bhref\\s*=\\s*(\"(?<v>[^\"]*)\"|'(?<v>[^']*)'|(?<v>[^\\s>]+))",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex BackgroundColorRegex = new Regex(
            "background(?:-color)?\\s*:\\s*(?<v>#[0-9a-fA-F]{3,8})",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        [Flags]
        private enum InlineStyle
        {
            None = 0,
            Strike = 1 << 0,
            Bold = 1 << 1,
            Italic = 1 << 2,
        }

        private sealed class Frame
        {
            public string TagName { get; }
            public InlineStyle Style { get; }
            public bool IsInlineCode { get; }
            public bool InlineCodeHasMonospace { get; }
            public string Href { get; }
            public int ContentStartIndex { get; }

            public Frame(
                string tagName,
                InlineStyle style = InlineStyle.None,
                bool isInlineCode = false,
                string href = null,
                bool inlineCodeHasMonospace = false,
                int contentStartIndex = -1)
            {
                TagName = tagName ?? string.Empty;
                Style = style;
                IsInlineCode = isInlineCode;
                InlineCodeHasMonospace = inlineCodeHasMonospace;
                Href = href;
                ContentStartIndex = contentStartIndex;
            }
        }

        public InlineParser(OneNoteStyleConfig styleConfig)
        {
            var inlineCode = styleConfig?.GetInlineCodeStyle();
            _inlineCodeFontFamily = inlineCode?.FontFamily ?? string.Empty;
            _inlineCodeBackgroundColor = inlineCode?.BackgroundColor ?? string.Empty;
            _inlineCodePaddingCount = inlineCode?.PaddingCount ?? 0;
            _tryRecognizeNonTeXShiftFormats = styleConfig?.TryRecognizeNonTeXShiftFormats ?? true;
        }
    }
}
