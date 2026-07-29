using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using TeXShift.Core.Configuration;
using TeXShift.Core.OneNote;
using TeXShift.Core.OneNoteMeta;
using TeXShift.Core.OneNoteToMarkdown.Abstractions;
using TeXShift.Core.OneNoteToMarkdown.Handlers;
using TeXShift.Core.OneNoteToMarkdown.Inlines;
using TeXShift.Core.Utils;

namespace TeXShift.Core.OneNoteToMarkdown
{
    /// <summary>
    /// Converts OneNote XML to Markdown by dispatching elements to specialized handlers
    /// and parsing inline formatting.
    /// </summary>
    public class OneNoteToMarkdownConverter : IOneNoteToMarkdownConverter, IOneNoteConverterContext
    {
        private readonly InlineParser _inlineParser;
        private readonly List<string> _warnings = new List<string>();
        private readonly List<IElementHandler> _handlers;

        // Explicit IOneNoteConverterContext
        public XNamespace OneNoteNamespace { get; } = OneNoteXml.Namespace;
        public OneNoteStyleConfig StyleConfig { get; }
        public int CurrentIndentLevel { get; set; }
        public int CurrentListIndex { get; set; }
        public bool CurrentListIsOrdered { get; set; }

        public OneNoteToMarkdownConverter(OneNoteStyleConfig styleConfig)
        {
            StyleConfig = styleConfig ?? throw new ArgumentNullException(nameof(styleConfig));
            _inlineParser = new InlineParser(styleConfig);
            _handlers = new List<IElementHandler>
            {
                new HeadingElementHandler(),
                new ListElementHandler(),
                new CodeBlockElementHandler(),
                new MathElementHandler(),
                new QuoteBlockElementHandler(),
                new TableElementHandler(),
                new MermaidElementHandler(),
                new HorizontalRuleElementHandler(),
                new ImageElementHandler(),
                new ParagraphElementHandler()
            };
        }

        public Task<ReverseConversionResult> ConvertToMarkdownAsync(XElement oneNoteElement)
        {
            // Keep reverse conversion off the caller thread (often UI/COM) while remaining CPU-bound.
            return Task.Run(() => ConvertToMarkdownCore(oneNoteElement));
        }

        string IOneNoteConverterContext.ParseInlineHtml(
            string oneNoteHtml,
            XElement containingOe,
            InlineParseMode mode)
        {
            string html = oneNoteHtml ?? string.Empty;
            string oeStyle = (string)containingOe?.Attribute("style") ?? string.Empty;
            string textColor = null;
            if (!CssColorParser.TryGetColorFromStyle(oeStyle, out textColor))
            {
                textColor = GetQuickStyleTextColor(containingOe);
            }

            if (!string.IsNullOrEmpty(textColor))
            {
                html = "<span style='color:" + textColor + "'>" + html + "</span>";
            }

            return _inlineParser.Parse(html, mode);
        }

        private string GetQuickStyleTextColor(XElement containingOe)
        {
            string index = (string)containingOe?.Attribute("quickStyleIndex");
            if (string.IsNullOrEmpty(index))
            {
                return null;
            }

            var page = containingOe.Ancestors(OneNoteNamespace + "Page").FirstOrDefault();
            var quickStyle = page?.Elements(OneNoteNamespace + "QuickStyleDef")
                .FirstOrDefault(element => string.Equals(
                    (string)element.Attribute("index"),
                    index,
                    StringComparison.Ordinal));

            return CssColorParser.TryNormalize(
                (string)quickStyle?.Attribute("fontColor"),
                out string color)
                ? color
                : null;
        }

        string IOneNoteConverterContext.ConvertOeChildrenToMarkdown(XElement oeChildren)
        {
            if (oeChildren == null)
            {
                return string.Empty;
            }

            int savedIndent = CurrentIndentLevel;
            int savedIndex = CurrentListIndex;
            bool savedOrdered = CurrentListIsOrdered;

            var blocks = ConvertOeChildren(oeChildren);
            var sb = new StringBuilder();
            AppendBlocks(sb, blocks);

            CurrentIndentLevel = savedIndent;
            CurrentListIndex = savedIndex;
            CurrentListIsOrdered = savedOrdered;

            return sb.ToString();
        }

        void IOneNoteConverterContext.AddWarning(string warning)
        {
            if (!string.IsNullOrWhiteSpace(warning))
            {
                _warnings.Add(warning);
            }
        }

        private ReverseConversionResult ConvertToMarkdownCore(XElement oneNoteElement)
        {
            _warnings.Clear();

            var markdown = new StringBuilder();

            foreach (var outline in EnumerateOutlines(oneNoteElement))
            {
                if (TryAppendMetaSource(outline, markdown))
                {
                    continue;
                }

                var oeChildren = outline.Element(OneNoteNamespace + "OEChildren");
                if (oeChildren == null)
                {
                    continue;
                }

                CurrentIndentLevel = 0;
                CurrentListIndex = 0;
                CurrentListIsOrdered = false;

                var blocks = ConvertOeChildren(oeChildren);
                AppendBlocks(markdown, blocks);
            }

            var result = new ReverseConversionResult
            {
                Markdown = markdown.ToString()
            };
            result.Warnings.AddRange(_warnings);
            return result;
        }

        private bool TryAppendMetaSource(XElement outline, StringBuilder markdown)
        {
            var metaResult = TeXShiftMetaReader.ReadOutline(outline);
            if (metaResult == null || !metaResult.HasTeXShiftMeta)
            {
                return false;
            }

            if (metaResult.IsValid)
            {
                var source = NormalizeMetaSource(metaResult.Source ?? string.Empty);
                if (markdown.Length > 0 && !string.IsNullOrWhiteSpace(source))
                {
                    EnsureTrailingNewlines(markdown, 2);
                }

                // Preserve the stored source verbatim (including trailing newline, if present).
                markdown.Append(source);
                return true;
            }

            if (!string.IsNullOrWhiteSpace(metaResult.FailureReason))
            {
                _warnings.Add(metaResult.FailureReason);
            }

            return false;
        }

        private static string NormalizeMetaSource(string source)
        {
            if (string.IsNullOrEmpty(source))
            {
                return string.Empty;
            }

            // Defensive: in some cases the captured source may include an extra leading newline
            // introduced during copy/paste or OneNote extraction. Avoid outputting an artificial
            // blank first line.
            if (source.StartsWith("\r\n", StringComparison.Ordinal))
            {
                return source.Substring(2);
            }

            if (source[0] == '\n' || source[0] == '\r')
            {
                return source.Substring(1);
            }

            return source;
        }

        private IEnumerable<XElement> EnumerateOutlines(XElement element)
        {
            if (element == null)
            {
                yield break;
            }

            // Phase 1: accept common roots (Outline/Page) and fall back to "wrap as Outline".
            if (element.Name.LocalName == "Outline")
            {
                yield return element;
                yield break;
            }

            if (element.Name.LocalName == "Page")
            {
                foreach (var outline in element.Descendants(OneNoteNamespace + "Outline"))
                {
                    yield return outline;
                }
                yield break;
            }

            if (element.Name.LocalName == "OEChildren")
            {
                yield return new XElement(OneNoteNamespace + "Outline", element);
                yield break;
            }

            if (element.Name.LocalName == "OE")
            {
                yield return new XElement(OneNoteNamespace + "Outline",
                    new XElement(OneNoteNamespace + "OEChildren", element));
                yield break;
            }

            // Best-effort: treat the passed element as a container.
            yield return new XElement(OneNoteNamespace + "Outline",
                new XElement(OneNoteNamespace + "OEChildren", element));
        }

        private sealed class BlockRecord
        {
            public string Text { get; }
            public bool IsListItem { get; }
            public bool IsBlank { get; }

            public BlockRecord(string text, bool isListItem, bool isBlank = false)
            {
                Text = text ?? string.Empty;
                IsListItem = isListItem;
                IsBlank = isBlank;
            }
        }

        private List<BlockRecord> ConvertOeChildren(XElement oeChildren)
        {
            var results = new List<BlockRecord>();
            if (oeChildren == null)
            {
                return results;
            }

            bool prevWasList = false;
            bool prevIsOrdered = false;

            foreach (var oe in oeChildren.Elements(OneNoteNamespace + "OE"))
            {
                bool isListItem = IsListItem(oe);
                bool isBlank = !isListItem && IsBlankOe(oe);
                bool isOrdered = isListItem && IsOrderedListItem(oe);

                if (!isBlank)
                {
                    if (isListItem)
                    {
                        if (!prevWasList || prevIsOrdered != isOrdered)
                        {
                            CurrentListIndex = 0;
                            CurrentListIsOrdered = isOrdered;
                        }
                        prevWasList = true;
                        prevIsOrdered = isOrdered;
                    }
                    else
                    {
                        prevWasList = false;
                        prevIsOrdered = false;
                        CurrentListIndex = 0;
                        CurrentListIsOrdered = false;
                    }
                }

                if (isBlank)
                {
                    results.Add(new BlockRecord(string.Empty, false, true));
                }
                else
                {
                    string block = ConvertOe(oe);
                    if (!string.IsNullOrWhiteSpace(block))
                    {
                        results.Add(new BlockRecord(block, isListItem));
                    }
                }

                var child = oe.Element(OneNoteNamespace + "OEChildren");
                if (child != null && child.Elements(OneNoteNamespace + "OE").Any())
                {
                    bool childHasList = child.Elements(OneNoteNamespace + "OE").Any(IsListItem);

                    int savedIndent = CurrentIndentLevel;
                    int savedIndex = CurrentListIndex;
                    bool savedOrdered = CurrentListIsOrdered;

                    if (isListItem)
                    {
                        CurrentIndentLevel++;
                        results.AddRange(ConvertOeChildren(child));
                        CurrentIndentLevel = savedIndent;
                    }
                    else if (childHasList)
                    {
                        results.AddRange(ConvertOeChildren(child));
                    }
                    else
                    {
                        CurrentIndentLevel++;
                        results.AddRange(ConvertOeChildren(child));
                        CurrentIndentLevel = savedIndent;
                    }

                    CurrentListIndex = savedIndex;
                    CurrentListIsOrdered = savedOrdered;
                }
            }

            return results;
        }

        private void AppendBlocks(StringBuilder markdown, List<BlockRecord> blocks)
        {
            if (blocks == null || blocks.Count == 0)
            {
                return;
            }

            BlockRecord previous = null;
            foreach (var block in blocks)
            {
                if (block == null)
                {
                    continue;
                }

                if (block.IsBlank)
                {
                    if (markdown.Length > 0)
                    {
                        EnsureTrailingNewlines(markdown, 2);
                    }
                    previous = null;
                    continue;
                }

                if (string.IsNullOrWhiteSpace(block.Text))
                {
                    continue;
                }

                if (markdown.Length > 0)
                {
                    bool listToList = previous != null && previous.IsListItem && block.IsListItem;
                    EnsureTrailingNewlines(markdown, listToList ? 1 : 2);
                }

                markdown.Append(block.Text.TrimEnd());
                previous = block;
            }
        }

        private static void EnsureTrailingNewlines(StringBuilder markdown, int count)
        {
            if (markdown == null || markdown.Length == 0 || count <= 0)
            {
                return;
            }

            int newlineCount = 0;
            int i = markdown.Length - 1;
            while (i >= 0)
            {
                if (markdown[i] == '\n')
                {
                    newlineCount++;
                    i--;
                    if (i >= 0 && markdown[i] == '\r')
                    {
                        i--;
                    }
                    continue;
                }
                if (markdown[i] == '\r')
                {
                    newlineCount++;
                    i--;
                    continue;
                }
                break;
            }

            while (newlineCount < count)
            {
                markdown.AppendLine();
                newlineCount++;
            }
        }

        private string ConvertOe(XElement oe)
        {
            if (oe == null)
            {
                return string.Empty;
            }

            foreach (var handler in _handlers)
            {
                if (handler.CanHandle(oe, this))
                {
                    var parts = handler.Handle(oe, this).Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
                    if (parts.Count == 0)
                    {
                        return string.Empty;
                    }
                    return string.Join("\n", parts);
                }
            }

            return ConvertOeBestEffort(oe);
        }

        private string ConvertOeBestEffort(XElement oe)
        {
            if (oe == null)
            {
                return string.Empty;
            }

            var tElements = oe.Descendants(OneNoteNamespace + "T").ToList();
            if (tElements.Count == 0)
            {
                return string.Empty;
            }

            var lines = new List<string>(tElements.Count);
            foreach (var t in tElements)
            {
                string html = t.Value ?? string.Empty;
                string parsed = _inlineParser.Parse(html);
                lines.Add(parsed);
            }

            return string.Join("\n", lines);
        }

        private bool IsListItem(XElement element)
        {
            if (element == null)
            {
                return false;
            }

            if (element.Element(OneNoteNamespace + "Tag") != null)
            {
                return true;
            }

            var list = element.Element(OneNoteNamespace + "List");
            if (list == null)
            {
                return false;
            }

            return list.Element(OneNoteNamespace + "Bullet") != null || list.Element(OneNoteNamespace + "Number") != null;
        }

        private bool IsOrderedListItem(XElement element)
        {
            if (element == null)
            {
                return false;
            }

            var list = element.Element(OneNoteNamespace + "List");
            return list != null && list.Element(OneNoteNamespace + "Number") != null;
        }

        private bool IsBlankOe(XElement element)
        {
            if (element == null)
            {
                return false;
            }

            var ns = OneNoteNamespace;

            var childOes = element.Element(ns + "OEChildren");
            if (childOes != null && childOes.Elements(ns + "OE").Any())
            {
                return false;
            }

            foreach (var child in element.Elements())
            {
                var localName = child.Name.LocalName;
                if (!string.Equals(localName, "T", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(localName, "Meta", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            var tElements = element.Elements(ns + "T").ToList();
            if (tElements.Count == 0)
            {
                return true;
            }

            foreach (var t in tElements)
            {
                var html = t.Value ?? string.Empty;
                var stripped = HtmlStripper.Strip(html);
                if (!string.IsNullOrWhiteSpace(stripped))
                {
                    return false;
                }
            }

            return true;
        }
    }
}
