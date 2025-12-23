using Markdig.Extensions.TaskLists;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TeXShift.Core.Utils;

namespace TeXShift.Core.Markdown
{
    /// <summary>
    /// Provides shared helper methods for handling images across different block handlers.
    /// Centralizes image detection, element creation, and fallback logic.
    /// </summary>
    internal static class ImageElementHelper
    {
        /// <summary>
        /// Checks if a paragraph contains only a single image and returns it.
        /// Filters out whitespace-only literals and optionally TaskList checkboxes.
        /// </summary>
        /// <param name="paragraph">The paragraph to check.</param>
        /// <param name="filterTaskList">Whether to filter out TaskList inline elements.</param>
        /// <returns>The image LinkInline if found, null otherwise.</returns>
        public static LinkInline GetSingleImage(ParagraphBlock paragraph, bool filterTaskList = false)
        {
            if (paragraph?.Inline == null) return null;

            var inlines = paragraph.Inline.ToList();

            // Filter out whitespace-only literals and optionally TaskList checkboxes
            var meaningfulInlines = inlines.Where(i =>
                !(i is LiteralInline lit && string.IsNullOrWhiteSpace(lit.Content.ToString())) &&
                !(filterTaskList && i is TaskList)).ToList();

            if (meaningfulInlines.Count == 1 && meaningfulInlines[0] is LinkInline link && link.IsImage)
            {
                return link;
            }

            return null;
        }

        /// <summary>
        /// Creates an Image XML element from a LinkInline.
        /// Returns null if the image fails to load.
        /// </summary>
        /// <param name="imageLink">The image link to process.</param>
        /// <param name="ns">The OneNote XML namespace.</param>
        /// <returns>An Image XElement, or null if loading fails.</returns>
        public static XElement CreateImageElement(LinkInline imageLink, XNamespace ns)
        {
            var url = imageLink?.Url ?? "";
            var altText = GetAltText(imageLink);

            var result = ImageLoader.LoadImage(url);
            if (!result.Success)
            {
                return null;
            }

            var imageElement = new XElement(ns + "Image",
                new XAttribute("format", result.Format));

            if (!string.IsNullOrEmpty(altText))
            {
                imageElement.Add(new XAttribute("alt", altText));
            }

            imageElement.Add(new XElement(ns + "Data", result.Base64Data));
            return imageElement;
        }

        /// <summary>
        /// Creates a fallback link element when image loading fails.
        /// </summary>
        /// <param name="imageLink">The image link to create a fallback for.</param>
        /// <param name="ns">The OneNote XML namespace.</param>
        /// <returns>A T XElement containing an HTML link.</returns>
        public static XElement CreateImageFallback(LinkInline imageLink, XNamespace ns)
        {
            var url = imageLink?.Url ?? "";
            var altText = GetAltText(imageLink);

            return new XElement(ns + "T",
                new XCData($"<a href=\"{HtmlEscaper.Escape(url)}\">[🖼️{HtmlEscaper.Escape(altText)}]</a>"));
        }

        /// <summary>
        /// Creates an OE element containing either an embedded image or a fallback link.
        /// </summary>
        /// <param name="imageLink">The image link to process.</param>
        /// <param name="ns">The OneNote XML namespace.</param>
        /// <returns>An OE XElement with image or fallback content.</returns>
        public static XElement CreateImageOE(LinkInline imageLink, XNamespace ns)
        {
            var imageElement = CreateImageElement(imageLink, ns);
            if (imageElement != null)
            {
                return new XElement(ns + "OE", imageElement);
            }
            return new XElement(ns + "OE", CreateImageFallback(imageLink, ns));
        }

        /// <summary>
        /// Extracts alt text from an image link.
        /// </summary>
        /// <param name="imageLink">The image link.</param>
        /// <returns>The alt text, or "image" if not found.</returns>
        public static string GetAltText(LinkInline imageLink)
        {
            if (imageLink?.FirstChild is LiteralInline literal)
            {
                return literal.Content.ToString();
            }
            return "image";
        }
    }
}
