using System.Collections.Generic;
using System.Xml.Linq;
using TeXShift.Core.OneNoteToMarkdown.Abstractions;

namespace TeXShift.Core.OneNoteToMarkdown.Handlers
{
    /// <summary>
    /// Defines the contract for a handler that converts a specific OneNote XML element pattern to Markdown.
    /// </summary>
    internal interface IElementHandler
    {
        bool CanHandle(XElement element, IOneNoteConverterContext context);

        IEnumerable<string> Handle(XElement element, IOneNoteConverterContext context);
    }
}

