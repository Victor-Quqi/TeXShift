using System.Runtime.Serialization;

namespace TeXShift.Core.Configuration
{
    /// <summary>
    /// Root configuration container for all TeXShift settings.
    /// Uses DataContract for JSON serialization with DataContractJsonSerializer.
    /// </summary>
    [DataContract]
    public class AppSettings
    {
        [DataMember]
        public DebugSettings Debug { get; set; }

        [DataMember]
        public CodeBlockStyleSettings CodeBlock { get; set; }

        [DataMember]
        public InlineCodeStyleSettings InlineCode { get; set; }

        [DataMember]
        public QuoteBlockStyleSettings QuoteBlock { get; set; }

        [DataMember]
        public HeadingStyleSettings Headings { get; set; }

        [DataMember]
        public LayoutSettings Layout { get; set; }

        [DataMember]
        public ImageSettings Image { get; set; }

        public AppSettings()
        {
            Debug = new DebugSettings();
            CodeBlock = new CodeBlockStyleSettings();
            InlineCode = new InlineCodeStyleSettings();
            QuoteBlock = new QuoteBlockStyleSettings();
            Headings = new HeadingStyleSettings();
            Layout = new LayoutSettings();
            Image = new ImageSettings();
        }

        /// <summary>
        /// Creates a new AppSettings instance with all default values.
        /// </summary>
        public static AppSettings CreateDefault()
        {
            return new AppSettings();
        }
    }

    /// <summary>
    /// Debug-related settings including debug button visibility.
    /// </summary>
    [DataContract]
    public class DebugSettings
    {
        /// <summary>
        /// Whether to show debug buttons in the Ribbon (调试转换, 查看XML).
        /// Default: false (hidden).
        /// </summary>
        [DataMember]
        public bool ShowDebugButtons { get; set; } = false;

        /// <summary>
        /// Custom output path for debug files.
        /// Empty string means use default (DebugOutput folder in project root).
        /// </summary>
        [DataMember]
        public string DebugOutputPath { get; set; } = "";
    }

    /// <summary>
    /// Code block styling settings.
    /// </summary>
    [DataContract]
    public class CodeBlockStyleSettings
    {
        /// <summary>
        /// Background color in hex format (e.g., "#0D1117").
        /// </summary>
        [DataMember]
        public string BackgroundColor { get; set; } = "#0D1117";

        /// <summary>
        /// Default text color in hex format (e.g., "#C9D1D9").
        /// </summary>
        [DataMember]
        public string TextColor { get; set; } = "#C9D1D9";

        /// <summary>
        /// Font family (e.g., "Consolas").
        /// </summary>
        [DataMember]
        public string FontFamily { get; set; } = "Consolas";

        /// <summary>
        /// Font size in points.
        /// </summary>
        [DataMember]
        public double FontSize { get; set; } = 11.0;

        /// <summary>
        /// Line height in points.
        /// </summary>
        [DataMember]
        public double LineHeight { get; set; } = 16.0;

        /// <summary>
        /// Whether to enable syntax highlighting.
        /// </summary>
        [DataMember]
        public bool EnableSyntaxHighlight { get; set; } = true;
    }

    /// <summary>
    /// Inline code styling settings.
    /// </summary>
    [DataContract]
    public class InlineCodeStyleSettings
    {
        /// <summary>
        /// Background color in hex format (e.g., "#F1F1F1").
        /// </summary>
        [DataMember]
        public string BackgroundColor { get; set; } = "#F1F1F1";

        /// <summary>
        /// Font family (e.g., "Consolas").
        /// </summary>
        [DataMember]
        public string FontFamily { get; set; } = "Consolas";
    }

    /// <summary>
    /// Quote block styling settings.
    /// </summary>
    [DataContract]
    public class QuoteBlockStyleSettings
    {
        /// <summary>
        /// Background color in hex format (e.g., "#E8F5E9").
        /// </summary>
        [DataMember]
        public string BackgroundColor { get; set; } = "#E8F5E9";
    }

    /// <summary>
    /// Heading style settings for H1-H6.
    /// </summary>
    [DataContract]
    public class HeadingStyleSettings
    {
        [DataMember]
        public double H1FontSize { get; set; } = 22.0;

        [DataMember]
        public double H2FontSize { get; set; } = 20.0;

        [DataMember]
        public double H3FontSize { get; set; } = 18.0;

        [DataMember]
        public double H4FontSize { get; set; } = 16.0;

        [DataMember]
        public double H5FontSize { get; set; } = 14.0;

        [DataMember]
        public double H6FontSize { get; set; } = 11.0;

        /// <summary>
        /// Gets font size for a specific heading level (1-6).
        /// </summary>
        public double GetFontSize(int level)
        {
            switch (level)
            {
                case 1: return H1FontSize;
                case 2: return H2FontSize;
                case 3: return H3FontSize;
                case 4: return H4FontSize;
                case 5: return H5FontSize;
                case 6: return H6FontSize;
                default: return H6FontSize;
            }
        }

        /// <summary>
        /// Sets font size for a specific heading level (1-6).
        /// </summary>
        public void SetFontSize(int level, double size)
        {
            switch (level)
            {
                case 1: H1FontSize = size; break;
                case 2: H2FontSize = size; break;
                case 3: H3FontSize = size; break;
                case 4: H4FontSize = size; break;
                case 5: H5FontSize = size; break;
                case 6: H6FontSize = size; break;
            }
        }
    }

    /// <summary>
    /// Layout settings for spacing and indentation.
    /// </summary>
    [DataContract]
    public class LayoutSettings
    {
        /// <summary>
        /// List indent width in points.
        /// </summary>
        [DataMember]
        public double ListIndent { get; set; } = 22.0;

        /// <summary>
        /// Default table column width in points.
        /// </summary>
        [DataMember]
        public double TableColumnWidth { get; set; } = 72.0;

        /// <summary>
        /// Paragraph spacing before in points.
        /// </summary>
        [DataMember]
        public double ParagraphSpaceBefore { get; set; } = 5.5;

        /// <summary>
        /// Paragraph spacing after in points.
        /// </summary>
        [DataMember]
        public double ParagraphSpaceAfter { get; set; } = 5.5;
    }

    /// <summary>
    /// Image loading settings.
    /// </summary>
    [DataContract]
    public class ImageSettings
    {
        /// <summary>
        /// Timeout for image downloads in seconds.
        /// </summary>
        [DataMember]
        public int DownloadTimeoutSeconds { get; set; } = 30;

        /// <summary>
        /// Maximum file size for images in bytes (default 10MB).
        /// </summary>
        [DataMember]
        public long MaxFileSizeBytes { get; set; } = 10 * 1024 * 1024;
    }
}
