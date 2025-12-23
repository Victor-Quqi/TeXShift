using System.Collections.Generic;

namespace TeXShift.Core.Configuration
{
    /// <summary>
    /// Configuration for OneNote element styles and spacing.
    /// </summary>
    public class OneNoteStyleConfig
    {
        public class SpacingConfig
        {
            public double SpaceBefore { get; set; }
            public double SpaceAfter { get; set; }
            public double SpaceBetween { get; set; }

            public SpacingConfig(double before, double after, double between)
            {
                SpaceBefore = before;
                SpaceAfter = after;
                SpaceBetween = between;
            }
        }

        public class FontConfig
        {
            public double FontSize { get; set; }
            public bool IsBold { get; set; }

            public FontConfig(double fontSize, bool isBold = true)
            {
                FontSize = fontSize;
                IsBold = isBold;
            }
        }

        public class InlineCodeConfig
        {
            public string FontFamily { get; set; }
            public string BackgroundColor { get; set; }
            public string PaddingChar { get; set; }
            public int PaddingCount { get; set; }


            public InlineCodeConfig(string fontFamily, string backgroundColor, string paddingChar = "&nbsp;", int paddingCount = 1)
            {
                FontFamily = fontFamily;
                BackgroundColor = backgroundColor;
                PaddingChar = paddingChar;
                PaddingCount = paddingCount;
            }
        }

        public enum HorizontalRuleMode
        {
            Character,
            Image
        }

        public class HorizontalRuleConfig
        {
            public HorizontalRuleMode Mode { get; set; }
            public string Color { get; set; }
            public int CharacterLength { get; set; }
            public char Character { get; set; }
            public int InitialImageWidth { get; set; }

            public HorizontalRuleConfig(HorizontalRuleMode mode, string color, int charLength, char character, int initialImageWidth)
            {
                Mode = mode;
                Color = color;
                CharacterLength = charLength;
                Character = character;
                InitialImageWidth = initialImageWidth;
            }
        }

        public class QuoteBlockConfig
        {
            public string BackgroundColor { get; set; }
            public double BaseWidth { get; set; }
            public double WidthReduction { get; set; }

            public QuoteBlockConfig(string backgroundColor = "#E2F5FE", double baseWidth = 534.0, double widthReduction = 13.52)
            {
                BackgroundColor = backgroundColor;
                BaseWidth = baseWidth;
                WidthReduction = widthReduction;
            }
        }

        public class CodeBlockConfig
        {
            public string BackgroundColor { get; set; }
            public string DefaultTextColor { get; set; }
            public string FontFamily { get; set; }
            public double FontSize { get; set; }
            public double LineHeight { get; set; }
            public bool EnableSyntaxHighlight { get; set; }

            // GitHub Dark Theme Defaults
            public CodeBlockConfig(
                string backgroundColor = "#0D1117",
                string defaultTextColor = "#C9D1D9",
                string fontFamily = "Consolas",
                double fontSize = 11.0,
                double lineHeight = 16.0,
                bool enableSyntaxHighlight = true)
            {
                BackgroundColor = backgroundColor;
                DefaultTextColor = defaultTextColor;
                FontFamily = fontFamily;
                FontSize = fontSize;
                LineHeight = lineHeight;
                EnableSyntaxHighlight = enableSyntaxHighlight;
            }

            /// <summary>
            /// Generate the style attribute value for the OE element
            /// </summary>
            public string GetOEStyle()
            {
                return $"line-height:{LineHeight:F1}pt;font-family:'{FontFamily}';font-size:{FontSize:F1}pt;color:{DefaultTextColor}";
            }
        }

        /// <summary>
        /// Configuration for width reservation to prevent content from exceeding container boundaries.
        /// Uses conservative values to ensure tables and nested elements never stretch the text box.
        /// All values are in points and can be adjusted based on actual font/style combinations.
        /// </summary>
        public class WidthReservationConfig
        {
            /// <summary>
            /// List indent width (must match Indents configuration).
            /// </summary>
            public double ListIndent { get; set; }

            /// <summary>
            /// Conservative reservation for table border + cell padding + rendering overhead.
            /// </summary>
            public double TableSystemOverhead { get; set; }

            /// <summary>
            /// Width reservation for unordered list markers (bullets).
            /// </summary>
            public double UnorderedListMarker { get; set; }

            /// <summary>
            /// Width reservation for ordered list markers (numbers up to 3 digits).
            /// </summary>
            public double OrderedListMarker { get; set; }

            /// <summary>
            /// Left margin for quote blocks (visual spacing).
            /// </summary>
            public double QuoteBlockMargin { get; set; }

            /// <summary>
            /// Width reservation for task list checkboxes (future use).
            /// </summary>
            public double TaskListCheckbox { get; set; }

            public WidthReservationConfig(
                double listIndent = 22.0,
                double tableSystemOverhead = 15.0,
                double unorderedListMarker = 16.0,
                double orderedListMarker = 28.0,
                double quoteBlockMargin = 8.0,
                double taskListCheckbox = 20.0)
            {
                ListIndent = listIndent;
                TableSystemOverhead = tableSystemOverhead;
                UnorderedListMarker = unorderedListMarker;
                OrderedListMarker = orderedListMarker;
                QuoteBlockMargin = quoteBlockMargin;
                TaskListCheckbox = taskListCheckbox;
            }

            /// <summary>
            /// Total width consumed by a quote block table (system overhead + left margin).
            /// </summary>
            public double QuoteBlockTotalReservation => TableSystemOverhead + QuoteBlockMargin;

            /// <summary>
            /// Calculates total width consumed by a list item (indent + marker width).
            /// </summary>
            public double GetListItemReservation(bool isOrdered)
            {
                return ListIndent + (isOrdered ? OrderedListMarker : UnorderedListMarker);
            }
        }

        // Default spacing configurations
        private static readonly Dictionary<string, SpacingConfig> DefaultSpacing = new Dictionary<string, SpacingConfig>
        {
            { "h1", new SpacingConfig(17.6, 17.6, 22.0) },
            { "h2", new SpacingConfig(15.4, 15.4, 19.0) },
            { "h3", new SpacingConfig(13.2, 13.2, 16.0) },
            { "h4", new SpacingConfig(11.0, 11.0, 16.0) },
            { "h5", new SpacingConfig(9.0, 9.0, 16.0) },
            { "h6", new SpacingConfig(7.0, 7.0, 16.0) },
            { "paragraph", new SpacingConfig(5.5, 5.5, 16.0) },
            { "list", new SpacingConfig(5.5, 5.5, 16.0) },
            { "code", new SpacingConfig(5.5, 5.5, 16.0) },
            { "quote", new SpacingConfig(5.5, 5.5, 16.0) }
        };

        // Default font size configurations for headings
        private static readonly Dictionary<string, FontConfig> DefaultFonts = new Dictionary<string, FontConfig>
        {
            { "h1", new FontConfig(22.0, true) },    // 一级标题：22pt 粗体
            { "h2", new FontConfig(20.0, true) },    // 二级标题：20pt 粗体
            { "h3", new FontConfig(18.0, true) },    // 三级标题：18pt 粗体
            { "h4", new FontConfig(16.0, true) },    // 四级标题：16pt 粗体
            { "h5", new FontConfig(14.0, true) },    // 五级标题：14pt 粗体
            { "h6", new FontConfig(11.0, true) }     // 六级标题：11pt 粗体
       };

       // Default style for inline code
       private static readonly InlineCodeConfig DefaultInlineCodeStyle = new InlineCodeConfig("Consolas", "#F1F1F1", "&nbsp;", 1); // Default: 1 non-breaking space

       // Default style for horizontal rule
       private static readonly HorizontalRuleConfig DefaultHorizontalRuleStyle = new HorizontalRuleConfig(HorizontalRuleMode.Image, "#888888", 90, '─', 2325);

       // Default style for quote blocks
       private static readonly QuoteBlockConfig DefaultQuoteBlockStyle = new QuoteBlockConfig("#E8F5E9", 534.0, 13.52);

       // Default style for code blocks (GitHub Dark theme)
       private static readonly CodeBlockConfig DefaultCodeBlockStyle = new CodeBlockConfig();

       // Default width reservation configuration
       // Note: Small reservation values prioritize maximum table width (96%+ fill rate).
       // This allows tables to nearly fill the text box, accepting slight text box expansion (~6%).
       private static readonly WidthReservationConfig DefaultWidthReservation = new WidthReservationConfig(
           listIndent: 22.0,              // Must match DefaultIndents
           tableSystemOverhead: 4.0,      // Minimal border + cell padding (prioritizes width)
           unorderedListMarker: 4.0,     // Bullet symbols with spacing
           orderedListMarker: 7.0,       // Numbers up to 3 digits (e.g., "999. ")
           quoteBlockMargin: 4.0,         // Minimal left margin (prioritizes width)
           taskListCheckbox: 5.0);       // Future task list support

       // Default indent configurations for nested content, matching onemark for consistency.
       private static readonly Dictionary<int, double> DefaultIndents = new Dictionary<int, double>
        {
            { 1, 22.0 },
            { 2, 22.0 },
            { 3, 22.0 },
            { 4, 22.0 }
        };

        private Dictionary<string, SpacingConfig> _customSpacing;
        private Dictionary<string, FontConfig> _customFonts;
        private Dictionary<int, double> _customIndents;
        private InlineCodeConfig _customInlineCodeStyle;
        private HorizontalRuleConfig _customHorizontalRuleStyle;
        private QuoteBlockConfig _customQuoteBlockStyle;
        private CodeBlockConfig _customCodeBlockStyle;
        private WidthReservationConfig _customWidthReservation;

        public IReadOnlyDictionary<int, double> Indents => _customIndents;
        public WidthReservationConfig WidthReservation => _customWidthReservation;

        public OneNoteStyleConfig()
        {
            _customSpacing = new Dictionary<string, SpacingConfig>(DefaultSpacing);
            _customFonts = new Dictionary<string, FontConfig>(DefaultFonts);
            _customIndents = new Dictionary<int, double>(DefaultIndents);
            _customInlineCodeStyle = DefaultInlineCodeStyle;
            _customHorizontalRuleStyle = DefaultHorizontalRuleStyle;
            _customQuoteBlockStyle = DefaultQuoteBlockStyle;
            _customCodeBlockStyle = DefaultCodeBlockStyle;
            _customWidthReservation = DefaultWidthReservation;
        }

        /// <summary>
        /// Gets spacing configuration for a heading level (1-6).
        /// </summary>
        public SpacingConfig GetHeadingSpacing(int level)
        {
            string key = $"h{level}";
            return _customSpacing.ContainsKey(key)
                ? _customSpacing[key]
                : _customSpacing["h6"]; // Fallback to h6 for levels > 6
        }

        /// <summary>
        /// Gets font configuration for a heading level (1-6).
        /// </summary>
        public FontConfig GetHeadingFont(int level)
        {
            string key = $"h{level}";
            return _customFonts.ContainsKey(key)
                ? _customFonts[key]
                : _customFonts["h6"]; // Fallback to h6 for levels > 6
        }

        /// <summary>
        /// Gets spacing configuration for paragraphs.
        /// </summary>
        public SpacingConfig GetParagraphSpacing()
        {
            return _customSpacing["paragraph"];
        }

        /// <summary>
        /// Gets spacing configuration for list items.
        /// </summary>
        public SpacingConfig GetListSpacing()
        {
            return _customSpacing["list"];
        }

        /// <summary>
        /// Gets spacing configuration for code blocks.
        /// </summary>
        public SpacingConfig GetCodeSpacing()
        {
            return _customSpacing["code"];
        }

        /// <summary>
        /// Gets style configuration for inline code.
        /// </summary>
        public InlineCodeConfig GetInlineCodeStyle()
        {
            return _customInlineCodeStyle;
        }

        /// <summary>
        /// Gets style configuration for horizontal rules.
        /// </summary>
        public HorizontalRuleConfig GetHorizontalRuleStyle()
        {
            return _customHorizontalRuleStyle;
        }

        /// <summary>
        /// Gets spacing configuration for quote blocks.
        /// </summary>
        public SpacingConfig GetQuoteSpacing()
        {
            return _customSpacing["quote"];
        }

        /// <summary>
        /// Gets style configuration for quote blocks.
        /// </summary>
        public QuoteBlockConfig GetQuoteBlockStyle()
        {
            return _customQuoteBlockStyle;
        }

        /// <summary>
        /// Gets style configuration for code blocks.
        /// </summary>
        public CodeBlockConfig GetCodeBlockStyle()
        {
            return _customCodeBlockStyle;
        }

        /// <summary>
        /// Allows customization of spacing for a specific element type.
        /// </summary>
        public void SetSpacing(string elementType, double before, double after, double between)
        {
            _customSpacing[elementType] = new SpacingConfig(before, after, between);
        }

        /// <summary>
        /// Allows customization of font for a specific heading level.
        /// </summary>
        public void SetHeadingFont(int level, double fontSize, bool isBold = true)
        {
            string key = $"h{level}";
            _customFonts[key] = new FontConfig(fontSize, isBold);
        }

        /// <summary>
        /// Allows customization of indent for a specific level.
        /// </summary>
        public void SetIndent(int level, double indent)
        {
            _customIndents[level] = indent;
        }

        /// <summary>
        /// Allows customization of style for inline code.
        /// </summary>
        public void SetInlineCodeStyle(string fontFamily, string backgroundColor, string paddingChar = "&nbsp;", int paddingCount = 1)
        {
            _customInlineCodeStyle = new InlineCodeConfig(fontFamily, backgroundColor, paddingChar, paddingCount);
        }

        /// <summary>
        /// Allows customization of style for horizontal rules.
        /// </summary>
        public void SetHorizontalRuleStyle(HorizontalRuleMode mode, string color, int charLength, char character, int initialImageWidth)
        {
            _customHorizontalRuleStyle = new HorizontalRuleConfig(mode, color, charLength, character, initialImageWidth);
        }

        /// <summary>
        /// Allows customization of style for code blocks.
        /// </summary>
        public void SetCodeBlockStyle(string backgroundColor, string textColor, string fontFamily, double fontSize, double lineHeight, bool enableSyntaxHighlight)
        {
            _customCodeBlockStyle = new CodeBlockConfig(backgroundColor, textColor, fontFamily, fontSize, lineHeight, enableSyntaxHighlight);
        }

        /// <summary>
        /// Allows customization of style for quote blocks.
        /// </summary>
        public void SetQuoteBlockStyle(string backgroundColor, double baseWidth = 534.0, double widthReduction = 13.52)
        {
            _customQuoteBlockStyle = new QuoteBlockConfig(backgroundColor, baseWidth, widthReduction);
        }

        /// <summary>
        /// Resets all spacing and font configurations to default values.
        /// </summary>
        public void ResetToDefaults()
        {
            _customSpacing = new Dictionary<string, SpacingConfig>(DefaultSpacing);
            _customFonts = new Dictionary<string, FontConfig>(DefaultFonts);
            _customIndents = new Dictionary<int, double>(DefaultIndents);
            _customInlineCodeStyle = DefaultInlineCodeStyle;
            _customHorizontalRuleStyle = DefaultHorizontalRuleStyle;
            _customQuoteBlockStyle = DefaultQuoteBlockStyle;
            _customCodeBlockStyle = DefaultCodeBlockStyle;
        }
    }
}
