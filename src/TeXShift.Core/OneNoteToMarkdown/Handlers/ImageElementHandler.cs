using System;
using System.Collections.Generic;
using System.Xml.Linq;
using TeXShift.Core.OneNoteToMarkdown.Abstractions;

namespace TeXShift.Core.OneNoteToMarkdown.Handlers
{
    /// <summary>
    /// Emits a placeholder for embedded OneNote images during reverse conversion.
    /// TeXShift currently does not persist image binaries as files, and inlining base64 data URLs can be extremely slow.
    /// </summary>
    internal sealed class ImageElementHandler : IElementHandler
    {
        public bool CanHandle(XElement element, IOneNoteConverterContext context)
        {
            if (element == null || context == null)
            {
                return false;
            }

            var ns = context.OneNoteNamespace;
            return element.Element(ns + "Image") != null;
        }

        public IEnumerable<string> Handle(XElement element, IOneNoteConverterContext context)
        {
            var ns = context.OneNoteNamespace;
            var image = element?.Element(ns + "Image");
            if (image == null)
            {
                yield break;
            }

            var format = (string)image.Attribute("format") ?? string.Empty;
            var alt = NormalizeAltText((string)image.Attribute("alt") ?? "image");

            context.AddWarning($"TeXShift: image detected; omitting binary payload for performance (alt={alt}, format={format}).");
            yield return $"[TeXShift: image omitted (alt={alt}, format={format})]";
        }

        private static string NormalizeAltText(string alt)
        {
            if (string.IsNullOrWhiteSpace(alt))
            {
                return "image";
            }

            // Keep Markdown well-formed.
            return alt.Replace("\r", " ").Replace("\n", " ").Replace("]", "\\]");
        }
    }
}
