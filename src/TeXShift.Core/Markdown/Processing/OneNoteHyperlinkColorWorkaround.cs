using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using TeXShift.Core.Utils;

namespace TeXShift.Core.Markdown.Processing
{
    internal static class OneNoteHyperlinkColorWorkaround
    {
        private const string DefaultTextColor = "#000000";

        private static readonly Regex AnchorStartTagRegex = new Regex(
            "<a\\b(?<attrs>(?:[^>\"']|\"[^\"]*\"|'[^']*')*)>",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static void Apply(XElement outline, XNamespace oneNoteNamespace)
        {
            if (outline == null)
            {
                return;
            }

            foreach (var oe in outline.Descendants(oneNoteNamespace + "OE"))
            {
                ApplyToOe(oe, oneNoteNamespace);
            }
        }

        private static void ApplyToOe(XElement oe, XNamespace oneNoteNamespace)
        {
            var textElements = oe.Elements(oneNoteNamespace + "T").ToList();
            if (textElements.Count == 0 ||
                !TryGetExplicitLinkColor(oe, oneNoteNamespace, out string linkColor))
            {
                return;
            }

            string existingStyle = (string)oe.Attribute("style") ?? string.Empty;
            string inheritedColor = CssColorParser.TryGetColorFromStyle(existingStyle, out string parsedColor)
                ? parsedColor
                : DefaultTextColor;

            oe.SetAttributeValue("style", SetColorProperty(existingStyle, linkColor));
            if (string.Equals(inheritedColor, linkColor, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            foreach (var textElement in textElements)
            {
                string html = textElement.Value ?? string.Empty;
                textElement.ReplaceNodes(new XCData(
                    "<span style='color:" + inheritedColor + "'>" + html + "</span>"));
            }
        }

        internal static bool TryGetExplicitLinkColor(
            XElement oe,
            XNamespace oneNoteNamespace,
            out string commonColor)
        {
            commonColor = null;
            bool foundLink = false;

            if (oe == null)
            {
                return false;
            }

            foreach (var textElement in oe.Elements(oneNoteNamespace + "T"))
            {
                string html = textElement.Value ?? string.Empty;
                foreach (Match match in AnchorStartTagRegex.Matches(html))
                {
                    foundLink = true;
                    if (!CssColorParser.TryGetColorFromAttributes(
                            match.Groups["attrs"].Value,
                            out string linkColor))
                    {
                        return false;
                    }

                    if (commonColor == null)
                    {
                        commonColor = linkColor;
                    }
                    else if (!string.Equals(
                        commonColor,
                        linkColor,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }
                }
            }

            return foundLink && commonColor != null;
        }

        private static string SetColorProperty(string style, string color)
        {
            var declarations = (style ?? string.Empty)
                .Split(';')
                .Select(value => value.Trim())
                .Where(value => value.Length > 0)
                .Where(value =>
                {
                    int colon = value.IndexOf(':');
                    return colon <= 0 || !string.Equals(
                        value.Substring(0, colon).Trim(),
                        "color",
                        StringComparison.OrdinalIgnoreCase);
                })
                .ToList();

            declarations.Add("color:" + color);
            return string.Join(";", declarations);
        }
    }
}
