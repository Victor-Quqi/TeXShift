using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TeXShift.Core.Errors;
using TeXShift.Core.Localization;
using TeXShift.Core.OneNoteMeta;

namespace TeXShift.Core.OneNote
{
    /// <summary>Inserts converted blocks at the captured text caret or selection.</summary>
    public static class OneNotePasteInsertion
    {
        private static readonly XNamespace Ns = OneNoteXml.Namespace;

        public static XElement CreateOutline(XElement page, XElement converted)
        {
            if (page == null) throw new ArgumentNullException(nameof(page));
            if (converted == null) throw new ArgumentNullException(nameof(converted));
            var copy = new XElement(page);
            var selected = copy.Descendants().Where(e => e.Attribute("selected") != null
                && !e.Descendants().Any(n => n.Attribute("selected") != null)).ToList();
            var inserted = converted.Element(Ns + "OEChildren")?.Elements(Ns + "OE")
                .Select(e => new XElement(e)).ToList() ?? new List<XElement>();
            if (inserted.Count == 0)
                throw new ContentWriteException(Resources.GetString("Error_NoValidTextContent"), "Clipboard conversion produced no blocks.");

            if (selected.Count == 0 && !copy.Elements(Ns + "Outline").Any())
            {
                var outline = new XElement(converted);
                outline.AddFirst(new XElement(Ns + "Position", new XAttribute("x", 36),
                    new XAttribute("y", 90), new XAttribute("z", 0)));
                PlaceCaretAtEnd(outline.Element(Ns + "OEChildren"));
                return outline;
            }

            var paragraphs = selected.Select(e => e.Name == Ns + "OE" ? e : e.Ancestors(Ns + "OE").FirstOrDefault())
                .Distinct().ToList();
            if (paragraphs.Count == 0 || paragraphs.Any(e => e == null)) throw InvalidSelection();
            var first = paragraphs[0];
            var last = paragraphs[paragraphs.Count - 1];
            var source = first.Ancestors(Ns + "Outline").FirstOrDefault();
            if (source == null || paragraphs.Any(e => e.Parent != first.Parent)) throw InvalidSelection();
            var range = first.Parent.Elements(Ns + "OE").SkipWhile(e => e != first)
                .TakeWhile(e => e != last).Concat(new[] { last }).ToList();
            // A text range spanning nested lists or table cells needs a structural selection editor.
            if (range.Count > 1 && range.Any(e => e.Element(Ns + "OEChildren") != null)) throw InvalidSelection();

            var prefix = BoundaryText(first, before: true);
            var suffix = BoundaryText(last, before: false);
            var replacement = new List<XElement>();
            if (inserted.Count == 1 && IsInlineParagraph(inserted[0]) && first.Element(Ns + "T") != null)
            {
                var merged = CopyParagraph(first, prefix.Concat(inserted[0].Elements(Ns + "T")));
                merged.Add(new XElement(Ns + "T", new XAttribute("selected", "all"), new XCData(string.Empty)));
                merged.Add(suffix);
                replacement.Add(merged);
            }
            else
            {
                if (prefix.Count > 0) replacement.Add(CopyParagraph(first, prefix));
                replacement.AddRange(inserted);
                // Keep the caret outside tables, images and lists after a block paste.
                var tail = CopyParagraph(last, new[] { new XElement(Ns + "T", new XAttribute("selected", "all"), new XCData(string.Empty)) }.Concat(suffix));
                replacement.Add(tail);
            }

            var nested = last.Element(Ns + "OEChildren");
            if (nested != null) replacement.Last().Add(new XElement(nested));
            foreach (var node in source.DescendantsAndSelf()) node.Attribute("selected")?.Remove();
            first.ReplaceWith(replacement);
            foreach (var node in range.Skip(1)) node.Remove();
            TeXShiftMetaWriter.RemoveExistingTeXShiftMeta(source, Ns);
            return source;
        }

        public static void MaterializeBinaryContent(XElement outline, Func<string, string> readBinary)
        {
            if (outline == null) throw new ArgumentNullException(nameof(outline));
            if (readBinary == null) throw new ArgumentNullException(nameof(readBinary));
            // UpdatePageContent replaces the whole outline and requires its retained binary data.
            foreach (var callback in outline.Descendants(Ns + "CallbackID").ToList())
                callback.ReplaceWith(new XElement(Ns + "Data", readBinary((string)callback.Attribute("callbackID"))));
        }

        private static List<XElement> BoundaryText(XElement paragraph, bool before)
        {
            var texts = paragraph.Elements(Ns + "T").ToList();
            int index = before ? texts.FindIndex(t => t.Attribute("selected") != null)
                : texts.FindLastIndex(t => t.Attribute("selected") != null);
            if (index < 0) return new List<XElement>();
            return (before ? texts.Take(index) : texts.Skip(index + 1))
                .Where(t => t.Value.Length > 0).Select(t => new XElement(t)).ToList();
        }

        public static bool ContentEquals(XElement original, XElement current)
        {
            if (original == null || current == null) return original == null && current == null;
            return XNode.DeepEquals(NormalizeForComparison(original), NormalizeForComparison(current));
        }

        private static XElement NormalizeForComparison(XElement outline)
        {
            var copy = new XElement(outline);
            foreach (var attribute in copy.DescendantsAndSelf().Attributes().Where(a => a.Name == "selected"
                || a.Name == "lastModifiedTime" || a.Name == "lastModifiedBy" || a.Name == "lastModifiedByInitials").ToList())
                attribute.Remove();
            foreach (var paragraph in copy.Descendants(Ns + "OE"))
            {
                var texts = paragraph.Elements(Ns + "T").ToList();
                if (texts.Count == 0) continue;
                texts[0].Value = string.Concat(texts.Select(t => t.Value));
                foreach (var text in texts.Skip(1)) text.Remove();
            }
            return copy;
        }

        private static XElement CopyParagraph(XElement original, IEnumerable<XElement> texts)
        {
            return new XElement(Ns + "OE",
                original.Attributes().Where(a => a.Name != "objectID" && a.Name != "selected"),
                original.Elements().Where(e => e.Name == Ns + "List" || e.Name == Ns + "Tag"), texts);
        }

        private static bool IsInlineParagraph(XElement paragraph) => paragraph.Elements().All(e => e.Name == Ns + "T")
            && paragraph.Attributes().All(a => a.Name == "spaceBefore" || a.Name == "spaceAfter" || a.Name == "spaceBetween");

        private static void PlaceCaretAtEnd(XElement children)
        {
            foreach (var node in children.DescendantsAndSelf()) node.Attribute("selected")?.Remove();
            children.Add(new XElement(Ns + "OE", new XElement(Ns + "T", new XAttribute("selected", "all"), new XCData(string.Empty))));
        }

        private static ContentReadException InvalidSelection() => new ContentReadException(
            Resources.GetString("Error_PasteConvert_InsertionPoint"), "Paste requires a text caret or a selection within one paragraph container.");
    }
}
