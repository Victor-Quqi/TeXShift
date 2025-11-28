using Markdig.Syntax;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Xml.Linq;
using TeXShift.Core.Configuration;
using TeXShift.Core.Markdown.Abstractions;

namespace TeXShift.Core.Markdown.Handlers
{
    internal class HorizontalRuleHandler : IBlockHandler
    {
        public IEnumerable<XElement> Handle(Block block, IMarkdownConverterContext context)
        {
            var styleConfig = context.StyleConfig.GetHorizontalRuleStyle();
            var ns = context.OneNoteNamespace;

            XElement oe;

            if (styleConfig.Mode == OneNoteStyleConfig.HorizontalRuleMode.Image)
            {
                var imageWidth = context.SourceOutlineWidth.HasValue
                    ? (int)(context.SourceOutlineWidth.Value * 4.2)
                    : styleConfig.InitialImageWidth;
                oe = CreateImageRule(ns, styleConfig, imageWidth);
            }
            else
            {
                oe = CreateCharacterRule(ns, styleConfig);
            }

            // Add spacing to the container element to ensure the "line" is visually separated.
            oe.Add(new XAttribute("spaceBefore", "8.0"));
            oe.Add(new XAttribute("spaceAfter", "8.0"));

            return new[] { oe };
        }

        private XElement CreateCharacterRule(XNamespace ns, OneNoteStyleConfig.HorizontalRuleConfig styleConfig)
        {
            var lineChars = new string(styleConfig.Character, styleConfig.CharacterLength);
            var htmlContent = $"<span style='font-size:8pt; color:{styleConfig.Color};'>{lineChars}</span>";

            return new XElement(ns + "OE",
                new XAttribute("alignment", "center"),
                new XElement(ns + "T",
                    new XCData(htmlContent)
                )
            );
        }

        private XElement CreateImageRule(XNamespace ns, OneNoteStyleConfig.HorizontalRuleConfig styleConfig, int imageWidth)
        {
            var base64Image = GenerateLineImageBase64(styleConfig.Color, imageWidth, 1);

            // Per OneNote's XML schema, size attributes are not allowed on the Image element.
            // The image's dimensions are determined by the image data itself.
            var imageElement = new XElement(ns + "Image",
                new XAttribute("format", "png"),
                new XElement(ns + "Data", base64Image)
            );

            return new XElement(ns + "OE",
                new XAttribute("alignment", "center"),
                imageElement
            );
        }

        private string GenerateLineImageBase64(string hexColor, int width, int height)
        {
            // Remove '#' if present
            if (hexColor.StartsWith("#"))
            {
                hexColor = hexColor.Substring(1);
            }

            // Parse ARGB from hex string
            int r = int.Parse(hexColor.Substring(0, 2), NumberStyles.HexNumber);
            int g = int.Parse(hexColor.Substring(2, 2), NumberStyles.HexNumber);
            int b = int.Parse(hexColor.Substring(4, 2), NumberStyles.HexNumber);
            var color = Color.FromArgb(r, g, b);

            using (var bmp = new Bitmap(width, height))
            {
                using (var graphics = Graphics.FromImage(bmp))
                {
                    graphics.Clear(color);
                }
                using (var ms = new MemoryStream())
                {
                    bmp.Save(ms, ImageFormat.Png);
                    return Convert.ToBase64String(ms.ToArray());
                }
            }
        }
    }
}