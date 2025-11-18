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

        private Dictionary<string, SpacingConfig> _customSpacing;

        public OneNoteStyleConfig()
        {
            _customSpacing = new Dictionary<string, SpacingConfig>(DefaultSpacing);
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
        /// Resets all spacing to default values.
        /// </summary>
        public void ResetToDefaults()
        {
            _customSpacing = new Dictionary<string, SpacingConfig>(DefaultSpacing);
        }
    }
}
