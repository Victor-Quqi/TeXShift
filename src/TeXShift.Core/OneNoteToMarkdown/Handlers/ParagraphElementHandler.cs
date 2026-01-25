using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TeXShift.Core.OneNoteToMarkdown.Abstractions;

namespace TeXShift.Core.OneNoteToMarkdown.Handlers
{
    /// <summary>
    /// Converts OneNote paragraph OEs to Markdown paragraphs.
    /// </summary>
    internal sealed class ParagraphElementHandler : IElementHandler
    {
        public bool CanHandle(XElement element, IOneNoteConverterContext context)
        {
            if (element == null)
            {
                return false;
            }

            var ns = context.OneNoteNamespace;
            return element.Elements(ns + "T").Any();
        }

        public IEnumerable<string> Handle(XElement element, IOneNoteConverterContext context)
        {
            var ns = context.OneNoteNamespace;
            var tElements = element.Elements(ns + "T").ToList();
            if (tElements.Count == 0)
            {
                yield break;
            }

            var lines = new List<string>(tElements.Count);
            foreach (var t in tElements)
            {
                string html = t.Value ?? string.Empty;
                string parsed = context.ParseInlineHtml(html);
                lines.Add(parsed);
            }

            yield return string.Join("\n", lines).TrimEnd();
        }
    }
}

