using System;
using System.Collections.Generic;
using System.Xml.Linq;
using TeXShift.Core.OneNoteToMarkdown.Abstractions;

namespace TeXShift.Core.OneNoteToMarkdown.Handlers
{
    /// <summary>
    /// Detects TeXShift-generated Mermaid images and emits a placeholder when the Mermaid source is missing.
    /// </summary>
    internal sealed class MermaidElementHandler : IElementHandler
    {
        public bool CanHandle(XElement element, IOneNoteConverterContext context)
        {
            if (element == null || context == null)
            {
                return false;
            }

            var ns = context.OneNoteNamespace;
            var image = element.Element(ns + "Image");
            if (image == null)
            {
                return false;
            }

            var alt = (string)image.Attribute("alt") ?? string.Empty;
            return string.Equals(alt, "mermaid", StringComparison.OrdinalIgnoreCase);
        }

        public IEnumerable<string> Handle(XElement element, IOneNoteConverterContext context)
        {
            context.AddWarning("TeXShift: Mermaid diagram detected but source is missing/expired; emitting placeholder.");
            yield return "[TeXShift: mermaid diagram omitted (source missing)]";
        }
    }
}
