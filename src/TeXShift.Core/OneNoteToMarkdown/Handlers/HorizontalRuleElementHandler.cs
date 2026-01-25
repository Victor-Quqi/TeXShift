using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TeXShift.Core.Configuration;
using TeXShift.Core.OneNoteToMarkdown.Abstractions;
using TeXShift.Core.OneNoteToMarkdown.Inlines;

namespace TeXShift.Core.OneNoteToMarkdown.Handlers
{
    /// <summary>
    /// Converts OneNote horizontal rules (character or image) to Markdown rules.
    /// </summary>
    internal sealed class HorizontalRuleElementHandler : IElementHandler
    {
        public bool CanHandle(XElement element, IOneNoteConverterContext context)
        {
            if (element == null || context == null)
            {
                return false;
            }

            if (!IsCentered(element))
            {
                return false;
            }

            var ns = context.OneNoteNamespace;
            var styleConfig = context.StyleConfig.GetHorizontalRuleStyle();

            var image = element.Element(ns + "Image");
            if (image != null)
            {
                var alt = image.Attribute("alt")?.Value;
                if (!string.IsNullOrEmpty(alt) &&
                    alt.Equals("mermaid", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                return IsHorizontalRuleImage(image, ns);
            }

            var t = element.Element(ns + "T");
            if (t != null)
            {
                return IsHorizontalRuleText(t.Value, styleConfig);
            }

            return false;
        }

        public IEnumerable<string> Handle(XElement element, IOneNoteConverterContext context)
        {
            yield return "---";
        }

        private static bool IsCentered(XElement element)
        {
            var alignment = element.Attribute("alignment")?.Value;
            return string.Equals(alignment, "center", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsHorizontalRuleText(string html, OneNoteStyleConfig.HorizontalRuleConfig styleConfig)
        {
            var text = HtmlStripper.Strip(html ?? string.Empty);
            text = text.Replace(" ", string.Empty).Replace("\u00A0", string.Empty).Trim();
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            var ruleChar = styleConfig.Character;
            if (ruleChar == '\0')
            {
                return false;
            }

            if (text.Any(c => c != ruleChar))
            {
                return false;
            }

            return text.Length >= System.Math.Max(3, styleConfig.CharacterLength);
        }

        private static bool IsHorizontalRuleImage(XElement image, XNamespace ns)
        {
            var data = image.Element(ns + "Data")?.Value;
            if (!TryGetPngDimensions(data, out int width, out int height))
            {
                return false;
            }

            if (height <= 4 && width >= System.Math.Max(50, height * 20))
            {
                return true;
            }

            return false;
        }

        private static bool TryGetPngDimensions(string base64, out int width, out int height)
        {
            width = 0;
            height = 0;

            if (string.IsNullOrWhiteSpace(base64))
            {
                return false;
            }

            byte[] data;
            try
            {
                data = Convert.FromBase64String(base64);
            }
            catch
            {
                return false;
            }

            if (data.Length < 24)
            {
                return false;
            }

            // PNG signature
            if (data[0] != 0x89 || data[1] != 0x50 || data[2] != 0x4E || data[3] != 0x47)
            {
                return false;
            }

            width = ReadBigEndianInt(data, 16);
            height = ReadBigEndianInt(data, 20);

            return width > 0 && height > 0;
        }

        private static int ReadBigEndianInt(byte[] data, int offset)
        {
            if (data == null || data.Length < offset + 4)
            {
                return 0;
            }

            return (data[offset] << 24)
                | (data[offset + 1] << 16)
                | (data[offset + 2] << 8)
                | data[offset + 3];
        }
    }
}
