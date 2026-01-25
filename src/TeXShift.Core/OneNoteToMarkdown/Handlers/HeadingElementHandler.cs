using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using TeXShift.Core.OneNoteToMarkdown.Abstractions;
using TeXShift.Core.OneNoteToMarkdown.Inlines;

namespace TeXShift.Core.OneNoteToMarkdown.Handlers
{
    /// <summary>
    /// Converts OneNote heading OEs to Markdown headings.
    /// </summary>
    internal sealed class HeadingElementHandler : IElementHandler
    {
        private static readonly Regex FontSizeRegex = new Regex(
            "font-size\\s*:\\s*(?<size>[0-9.]+)pt",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public bool CanHandle(XElement element, IOneNoteConverterContext context)
        {
            return GetHeadingLevel(element, context) > 0;
        }

        public IEnumerable<string> Handle(XElement element, IOneNoteConverterContext context)
        {
            int level = GetHeadingLevel(element, context);
            if (level <= 0)
            {
                yield break;
            }

            var ns = context.OneNoteNamespace;
            var tElements = element.Elements(ns + "T").ToList();
            if (tElements.Count == 0)
            {
                yield break;
            }

            // OneNote may inject empty <one:T> nodes (e.g., selection/cursor artifacts). Ignore them here.
            var parts = new List<string>(tElements.Count);
            foreach (var t in tElements)
            {
                string html = t?.Value ?? string.Empty;
                if (string.IsNullOrWhiteSpace(html))
                {
                    continue;
                }

                if (TryUnwrapHeadingSpan(html, out string innerHtml))
                {
                    html = innerHtml;
                }

                // Headings already carry structure via '#', so suppress bold markers (they're not recoverable anyway).
                parts.Add(context.ParseInlineHtml(html, InlineParseMode.Heading));
            }

            string content = string.Join("\n", parts).Trim();
            content = content.Replace("\r", string.Empty).Replace("\n", " ").Trim();

            string prefix = new string('#', System.Math.Min(6, System.Math.Max(1, level)));
            yield return $"{prefix} {content}".TrimEnd();
        }

        private int GetHeadingLevel(XElement element, IOneNoteConverterContext context)
        {
            if (element == null)
            {
                return 0;
            }

            var ns = context.OneNoteNamespace;

            // Never treat list items as headings; OneNote lists are represented as OEs with <one:List>/<one:Tag>.
            if (IsListItem(element, ns))
            {
                return 0;
            }

            var tElements = element.Elements(ns + "T").ToList();
            if (tElements.Count == 0)
            {
                return 0;
            }

            string combinedHtml = string.Concat(tElements.Select(t => t.Value ?? string.Empty));

            double fontSize;
            if (!TryExtractHeadingFontSizeFromOeStyle(element.Attribute("style")?.Value, out fontSize) &&
                !tElements.Any(t => TryExtractHeadingFontSizeFromOeStyle(t.Attribute("style")?.Value, out fontSize)) &&
                !TryExtractHeadingFontSize(combinedHtml, out fontSize))
            {
                return 0;
            }

            int level = MapFontSizeToHeadingLevel(fontSize, context);

            // Always try to recognize TeXShift-generated headings first (even in fuzzy mode).
            // This ensures TeXShift H6 (11pt) can still be recognized by spacing/font rules.
            if (IsTeXShiftHeading(element, combinedHtml, fontSize, level, context))
            {
                return level;
            }

            if (!context.StyleConfig.TryRecognizeNonTeXShiftFormats)
            {
                return 0;
            }

            // Fuzzy mode: treat sufficiently large, bold text as a heading.
            // This intentionally avoids quickStyleIndex, which is not a stable indicator of headings.
            if (!IsBold(combinedHtml) || fontSize < 14.0)
            {
                return 0;
            }

            return level;
        }

        private int MapFontSizeToHeadingLevel(double fontSize, IOneNoteConverterContext context)
        {
            int bestLevel = 1;
            double bestDelta = double.MaxValue;

            for (int level = 1; level <= 6; level++)
            {
                var cfg = context.StyleConfig.GetHeadingFont(level);
                double delta = System.Math.Abs(cfg.FontSize - fontSize);
                if (delta < bestDelta)
                {
                    bestDelta = delta;
                    bestLevel = level;
                }
            }

            return bestLevel;
        }

        private bool TryUnwrapHeadingSpan(string html, out string innerHtml)
        {
            innerHtml = html;
            if (string.IsNullOrEmpty(html))
            {
                return false;
            }

            var match = HtmlRegexes.OuterSpan.Match(html);
            if (!match.Success)
            {
                return false;
            }

            string attrs = match.Groups["attrs"].Value ?? string.Empty;
            string style = GetStyle(attrs);
            if (string.IsNullOrEmpty(style))
            {
                return false;
            }

            if (style.IndexOf("font-size", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return false;
            }

            innerHtml = match.Groups["inner"].Value ?? string.Empty;
            return true;
        }

        private static bool IsListItem(XElement element, XNamespace ns)
        {
            if (element == null)
            {
                return false;
            }

            if (element.Element(ns + "Tag") != null)
            {
                return true;
            }

            var list = element.Element(ns + "List");
            if (list == null)
            {
                return false;
            }

            return list.Element(ns + "Bullet") != null || list.Element(ns + "Number") != null;
        }

        private bool IsTeXShiftHeading(
            XElement element,
            string combinedHtml,
            double fontSize,
            int level,
            IOneNoteConverterContext context)
        {
            var expectedFont = context.StyleConfig.GetHeadingFont(level);
            var expectedSpacing = context.StyleConfig.GetHeadingSpacing(level);

            if (!IsClose(fontSize, expectedFont.FontSize, 0.25))
            {
                return false;
            }

            if (!TryParseDoubleInvariant(element.Attribute("spaceBefore")?.Value, out double spaceBefore) ||
                !TryParseDoubleInvariant(element.Attribute("spaceAfter")?.Value, out double spaceAfter))
            {
                return false;
            }

            if (!IsClose(spaceBefore, expectedSpacing.SpaceBefore, 0.6) ||
                !IsClose(spaceAfter, expectedSpacing.SpaceAfter, 0.6))
            {
                return false;
            }

            if (expectedFont.IsBold && !IsBold(combinedHtml))
            {
                return false;
            }

            return true;
        }

        private bool TryExtractHeadingFontSize(string html, out double fontSize)
        {
            fontSize = 0.0;
            if (string.IsNullOrEmpty(html))
            {
                return false;
            }

            var match = FontSizeRegex.Match(html);
            if (!match.Success)
            {
                return false;
            }

            return double.TryParse(
                match.Groups["size"].Value,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out fontSize);
        }

        private bool TryExtractHeadingFontSizeFromOeStyle(string oeStyle, out double fontSize)
        {
            fontSize = 0.0;
            if (string.IsNullOrEmpty(oeStyle))
            {
                return false;
            }

            var match = FontSizeRegex.Match(oeStyle);
            if (!match.Success)
            {
                return false;
            }

            return double.TryParse(
                match.Groups["size"].Value,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out fontSize);
        }

        private bool IsBold(string html)
        {
            if (string.IsNullOrEmpty(html))
            {
                return false;
            }

            return html.IndexOf("font-weight:bold", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool TryParseDoubleInvariant(string text, out double value)
        {
            return double.TryParse(
                text,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out value);
        }

        private static bool IsClose(double actual, double expected, double tolerance)
        {
            return System.Math.Abs(actual - expected) <= tolerance;
        }

        private string GetStyle(string attrs)
        {
            if (string.IsNullOrEmpty(attrs))
            {
                return null;
            }

            var match = HtmlRegexes.StyleAttr.Match(attrs);
            return match.Success ? match.Groups["v"].Value : null;
        }
    }
}
