using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TeXShift.Core.OneNoteToMarkdown.Abstractions;

namespace TeXShift.Core.OneNoteToMarkdown.Handlers
{
    /// <summary>
    /// Converts OneNote list OEs (bullet/number/task) to Markdown list items.
    /// </summary>
    internal sealed class ListElementHandler : IElementHandler
    {
        public bool CanHandle(XElement element, IOneNoteConverterContext context)
        {
            return IsListItem(element, context);
        }

        public IEnumerable<string> Handle(XElement element, IOneNoteConverterContext context)
        {
            if (element == null)
            {
                yield break;
            }

            var ns = context.OneNoteNamespace;

            bool isTask = element.Element(ns + "Tag") != null;
            bool isOrdered = !isTask && IsOrderedList(element, context);

            string prefix = BuildPrefix(context, isOrdered, isTask, element);

            var tElements = element.Elements(ns + "T").ToList();
            var contentLines = new List<string>(tElements.Count);
            foreach (var t in tElements)
            {
                string html = t.Value ?? string.Empty;
                string parsed = context.ParseInlineHtml(html, element);
                contentLines.Add(parsed);
            }

            string content = contentLines.Count == 0 ? string.Empty : string.Join("\n", contentLines);
            yield return (prefix + content).TrimEnd();
        }

        private string BuildPrefix(IOneNoteConverterContext context, bool isOrdered, bool isTask, XElement element)
        {
            string indent = new string(' ', context.CurrentIndentLevel * 4);

            if (isTask)
            {
                bool completed = IsTaskCompleted(element, context);
                return indent + (completed ? "- [x] " : "- [ ] ");
            }

            if (isOrdered)
            {
                if (!context.CurrentListIsOrdered)
                {
                    context.CurrentListIsOrdered = true;
                    context.CurrentListIndex = 0;
                }

                context.CurrentListIndex++;
                return indent + context.CurrentListIndex.ToString() + ". ";
            }

            context.CurrentListIsOrdered = false;
            context.CurrentListIndex = 0;
            return indent + "- ";
        }

        private bool IsTaskCompleted(XElement element, IOneNoteConverterContext context)
        {
            if (element == null)
            {
                return false;
            }

            var tag = element.Element(context.OneNoteNamespace + "Tag");
            if (tag == null)
            {
                return false;
            }

            var completedAttr = tag.Attribute("completed")?.Value;
            if (string.IsNullOrWhiteSpace(completedAttr))
            {
                return false;
            }

            return completedAttr.Equals("true", StringComparison.OrdinalIgnoreCase);
        }

        private bool IsListItem(XElement element, IOneNoteConverterContext context)
        {
            if (element == null)
            {
                return false;
            }

            var ns = context.OneNoteNamespace;
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

        private bool IsOrderedList(XElement element, IOneNoteConverterContext context)
        {
            if (element == null)
            {
                return false;
            }

            var list = element.Element(context.OneNoteNamespace + "List");
            return list != null && list.Element(context.OneNoteNamespace + "Number") != null;
        }
    }
}
