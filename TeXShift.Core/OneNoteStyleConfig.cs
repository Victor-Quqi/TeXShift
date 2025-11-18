using System.Collections.Generic;

namespace TeXShift.Core
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
            { "code", new SpacingConfig(5.5, 5.5, 16.0) }
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

        public IReadOnlyDictionary<int, double> Indents => _customIndents;
 
        public OneNoteStyleConfig()
        {
            _customSpacing = new Dictionary<string, SpacingConfig>(DefaultSpacing);
            _customFonts = new Dictionary<string, FontConfig>(DefaultFonts);
            _customIndents = new Dictionary<int, double>(DefaultIndents);
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
        /// Resets all spacing and font configurations to default values.
        /// </summary>
        public void ResetToDefaults()
        {
            _customSpacing = new Dictionary<string, SpacingConfig>(DefaultSpacing);
            _customFonts = new Dictionary<string, FontConfig>(DefaultFonts);
            _customIndents = new Dictionary<int, double>(DefaultIndents);
        }
    }
}
