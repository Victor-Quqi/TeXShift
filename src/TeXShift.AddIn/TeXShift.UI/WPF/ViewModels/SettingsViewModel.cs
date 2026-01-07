using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using System.Windows.Media;
using TeXShift.AddIn.Localization;
using TeXShift.Core.Configuration;
using TeXShift.Core.Localization;

namespace TeXShift.AddIn.UI.WPF.ViewModels
{
    /// <summary>
    /// ViewModel for the Settings Window.
    /// </summary>
    public class SettingsViewModel : ViewModelBase
    {
        #region Debug Settings

        private bool _showDebugButtons;
        public bool ShowDebugButtons
        {
            get => _showDebugButtons;
            set => SetProperty(ref _showDebugButtons, value);
        }

        private bool _exportPdf;
        public bool ExportPdf
        {
            get => _exportPdf;
            set => SetProperty(ref _exportPdf, value);
        }

        private string _debugOutputPath;
        public string DebugOutputPath
        {
            get => _debugOutputPath;
            set => SetProperty(ref _debugOutputPath, value);
        }

        #endregion

        #region Code Block Settings

        private string _codeBlockBackgroundColor;
        public string CodeBlockBackgroundColor
        {
            get => _codeBlockBackgroundColor;
            set => SetProperty(ref _codeBlockBackgroundColor, value);
        }

        private string _codeBlockTextColor;
        public string CodeBlockTextColor
        {
            get => _codeBlockTextColor;
            set => SetProperty(ref _codeBlockTextColor, value);
        }

        private string _codeBlockFontFamily;
        public string CodeBlockFontFamily
        {
            get => _codeBlockFontFamily;
            set => SetProperty(ref _codeBlockFontFamily, value);
        }

        private double _codeBlockFontSize;
        public double CodeBlockFontSize
        {
            get => _codeBlockFontSize;
            set => SetProperty(ref _codeBlockFontSize, value);
        }

        private double _codeBlockSpaceBetween;
        public double CodeBlockSpaceBetween
        {
            get => _codeBlockSpaceBetween;
            set => SetProperty(ref _codeBlockSpaceBetween, value);
        }

        private bool _enableSyntaxHighlight;
        public bool EnableSyntaxHighlight
        {
            get => _enableSyntaxHighlight;
            set => SetProperty(ref _enableSyntaxHighlight, value);
        }

        #endregion

        #region Inline Code Settings

        private string _inlineCodeBackgroundColor;
        public string InlineCodeBackgroundColor
        {
            get => _inlineCodeBackgroundColor;
            set => SetProperty(ref _inlineCodeBackgroundColor, value);
        }

        private string _inlineCodeFontFamily;
        public string InlineCodeFontFamily
        {
            get => _inlineCodeFontFamily;
            set => SetProperty(ref _inlineCodeFontFamily, value);
        }

        #endregion

        #region Quote Block Settings

        private string _quoteBlockBackgroundColor;
        public string QuoteBlockBackgroundColor
        {
            get => _quoteBlockBackgroundColor;
            set => SetProperty(ref _quoteBlockBackgroundColor, value);
        }

        #endregion

        #region Heading Settings

        private double _h1FontSize;
        public double H1FontSize
        {
            get => _h1FontSize;
            set => SetProperty(ref _h1FontSize, value);
        }

        private double _h2FontSize;
        public double H2FontSize
        {
            get => _h2FontSize;
            set => SetProperty(ref _h2FontSize, value);
        }

        private double _h3FontSize;
        public double H3FontSize
        {
            get => _h3FontSize;
            set => SetProperty(ref _h3FontSize, value);
        }

        private double _h4FontSize;
        public double H4FontSize
        {
            get => _h4FontSize;
            set => SetProperty(ref _h4FontSize, value);
        }

        private double _h5FontSize;
        public double H5FontSize
        {
            get => _h5FontSize;
            set => SetProperty(ref _h5FontSize, value);
        }

        private double _h6FontSize;
        public double H6FontSize
        {
            get => _h6FontSize;
            set => SetProperty(ref _h6FontSize, value);
        }

        #endregion

        #region Language Settings

        private LanguageOption _selectedLanguage;
        public LanguageOption SelectedLanguage
        {
            get => _selectedLanguage;
            set => SetProperty(ref _selectedLanguage, value);
        }

        public ObservableCollection<LanguageOption> AvailableLanguages { get; } =
            new ObservableCollection<LanguageOption>
            {
                new LanguageOption("auto", Resources.GetString("UI_Language_System")),
                new LanguageOption("zh-CN", Resources.GetString("UI_Language_Chinese")),
                new LanguageOption("en-US", Resources.GetString("UI_Language_English"))
            };

        public string LanguageSectionTitle { get; } = Resources.GetString("UI_LanguageTitle");
        public string LanguageLabel { get; } = Resources.GetString("UI_LanguageLabel");

        #endregion

        #region UI Localization

        public string WindowTitle => UIResources.GetString("Settings_Title");

        public string TabStyleHeader => UIResources.GetString("Settings_Tab_Style");
        public string TabCodeBlockHeader => UIResources.GetString("Settings_Tab_CodeBlock");
        public string TabDebugHeader => UIResources.GetString("Settings_Tab_Debug");

        public string QuoteBlockSectionTitle => UIResources.GetString("Settings_Section_QuoteBlock");
        public string HeadingSizeSectionTitle => UIResources.GetString("Settings_Section_HeadingSize");
        public string CodeBlockSectionTitle => UIResources.GetString("Settings_Section_CodeBlock");
        public string InlineCodeSectionTitle => UIResources.GetString("Settings_Section_InlineCode");

        public string BackgroundColorLabel => UIResources.GetString("Settings_Label_BackgroundColor");
        public string TextColorLabel => UIResources.GetString("Settings_Label_TextColor");
        public string FontLabel => UIResources.GetString("Settings_Label_Font");
        public string FontSizeLabel => UIResources.GetString("Settings_Label_FontSize");
        public string LineSpacingLabel => UIResources.GetString("Settings_Label_LineSpacing");

        public string SelectButtonText => UIResources.GetString("Settings_Button_Select");
        public string BrowseButtonText => UIResources.GetString("Settings_Button_Browse");
        public string ResetDefaultsButtonText => UIResources.GetString("Settings_Button_ResetDefaults");
        public string CancelButtonText => UIResources.GetString("Settings_Button_Cancel");
        public string OkButtonText => UIResources.GetString("Settings_Button_Ok");

        public string EnableSyntaxHighlightText => UIResources.GetString("Settings_Checkbox_EnableSyntaxHighlight");
        public string ShowDebugButtonsText => UIResources.GetString("Settings_Checkbox_ShowDebugButtons");
        public string ExportPdfText => UIResources.GetString("Settings_Checkbox_ExportPdf");
        public string DebugOutputPathDescription => UIResources.GetString("Settings_Description_DebugOutputPath");
        public string DefaultDebugOutputPath => TeXShift.Core.Logging.DebugLogger.ResolveDebugOutputFolder(null);

        #endregion

        #region Font Options

        public ObservableCollection<string> AvailableFonts { get; } =
            new ObservableCollection<string>(FontPresets.MonospaceFonts);

        #endregion

        #region Preserved Settings

        private LayoutSettings _layoutSettings;
        private ImageSettings _imageSettings;

        #endregion

        #region Commands

        public ICommand ResetCommand { get; }

        #endregion

        #region Color Properties for ColorPicker Binding

        private Color _codeBlockBgColor;
        public Color CodeBlockBgColor
        {
            get => _codeBlockBgColor;
            set
            {
                if (SetProperty(ref _codeBlockBgColor, value))
                {
                    CodeBlockBackgroundColor = $"#{value.R:X2}{value.G:X2}{value.B:X2}";
                }
            }
        }

        private Color _codeBlockTxtColor;
        public Color CodeBlockTxtColor
        {
            get => _codeBlockTxtColor;
            set
            {
                if (SetProperty(ref _codeBlockTxtColor, value))
                {
                    CodeBlockTextColor = $"#{value.R:X2}{value.G:X2}{value.B:X2}";
                }
            }
        }

        private Color _inlineCodeBgColor;
        public Color InlineCodeBgColor
        {
            get => _inlineCodeBgColor;
            set
            {
                if (SetProperty(ref _inlineCodeBgColor, value))
                {
                    InlineCodeBackgroundColor = $"#{value.R:X2}{value.G:X2}{value.B:X2}";
                }
            }
        }

        private Color _quoteBlockBgColor;
        public Color QuoteBlockBgColor
        {
            get => _quoteBlockBgColor;
            set
            {
                if (SetProperty(ref _quoteBlockBgColor, value))
                {
                    QuoteBlockBackgroundColor = $"#{value.R:X2}{value.G:X2}{value.B:X2}";
                }
            }
        }

        #endregion

        public SettingsViewModel()
        {
            ResetCommand = new RelayCommand(ResetToDefaults);
            LoadFromSettings(AppSettings.CreateDefault());
        }

        /// <summary>
        /// Load settings from AppSettings into ViewModel properties.
        /// </summary>
        public void LoadFromSettings(AppSettings settings)
        {
            if (settings == null)
                settings = AppSettings.CreateDefault();

            var debugSettings = settings.Debug ?? new DebugSettings();
            var codeBlockSettings = settings.CodeBlock ?? new CodeBlockStyleSettings();
            var inlineCodeSettings = settings.InlineCode ?? new InlineCodeStyleSettings();
            var quoteBlockSettings = settings.QuoteBlock ?? new QuoteBlockStyleSettings();
            var headingSettings = settings.Headings ?? new HeadingStyleSettings();

            var codeBlockFontFamily = string.IsNullOrWhiteSpace(codeBlockSettings.FontFamily) ? "Consolas" : codeBlockSettings.FontFamily;
            var inlineCodeFontFamily = string.IsNullOrWhiteSpace(inlineCodeSettings.FontFamily) ? "Consolas" : inlineCodeSettings.FontFamily;
            EnsureFontAvailable(codeBlockFontFamily);
            EnsureFontAvailable(inlineCodeFontFamily);

            // Debug
            ShowDebugButtons = debugSettings.ShowDebugButtons;
            ExportPdf = debugSettings.ExportPdf;
            DebugOutputPath = debugSettings.DebugOutputPath ?? string.Empty;

            // Code Block
            CodeBlockBackgroundColor = string.IsNullOrWhiteSpace(codeBlockSettings.BackgroundColor) ? "#0D1117" : codeBlockSettings.BackgroundColor;
            CodeBlockTextColor = string.IsNullOrWhiteSpace(codeBlockSettings.TextColor) ? "#C9D1D9" : codeBlockSettings.TextColor;
            CodeBlockFontFamily = codeBlockFontFamily;
            CodeBlockFontSize = codeBlockSettings.FontSize;
            CodeBlockSpaceBetween = codeBlockSettings.SpaceBetween;
            EnableSyntaxHighlight = codeBlockSettings.EnableSyntaxHighlight;

            // Inline Code
            InlineCodeBackgroundColor = string.IsNullOrWhiteSpace(inlineCodeSettings.BackgroundColor) ? "#F1F1F1" : inlineCodeSettings.BackgroundColor;
            InlineCodeFontFamily = inlineCodeFontFamily;

            // Quote Block
            QuoteBlockBackgroundColor = string.IsNullOrWhiteSpace(quoteBlockSettings.BackgroundColor) ? "#E8F5E9" : quoteBlockSettings.BackgroundColor;

            // Headings
            H1FontSize = headingSettings.H1FontSize;
            H2FontSize = headingSettings.H2FontSize;
            H3FontSize = headingSettings.H3FontSize;
            H4FontSize = headingSettings.H4FontSize;
            H5FontSize = headingSettings.H5FontSize;
            H6FontSize = headingSettings.H6FontSize;

            // Preserve Layout/Image settings
            _layoutSettings = CloneLayoutSettings(settings.Layout);
            _imageSettings = CloneImageSettings(settings.Image);

            // Language
            var languageCode = string.IsNullOrWhiteSpace(settings.Language) ? "auto" : settings.Language;
            SelectedLanguage = AvailableLanguages.FirstOrDefault(option =>
                string.Equals(option.Code, languageCode, StringComparison.OrdinalIgnoreCase))
                ?? AvailableLanguages.First();

            // Update Color properties for ColorPicker
            CodeBlockBgColor = HexToColor(CodeBlockBackgroundColor);
            CodeBlockTxtColor = HexToColor(CodeBlockTextColor);
            InlineCodeBgColor = HexToColor(InlineCodeBackgroundColor);
            QuoteBlockBgColor = HexToColor(QuoteBlockBackgroundColor);
        }

        /// <summary>
        /// Export ViewModel properties to AppSettings.
        /// </summary>
        public AppSettings ToSettings()
        {
            var codeBlockFontFamily = string.IsNullOrWhiteSpace(CodeBlockFontFamily) ? "Consolas" : CodeBlockFontFamily;
            var inlineCodeFontFamily = string.IsNullOrWhiteSpace(InlineCodeFontFamily) ? "Consolas" : InlineCodeFontFamily;
            var debugOutputPath = DebugOutputPath?.Trim() ?? string.Empty;
            var codeBlockBackgroundColor = string.IsNullOrWhiteSpace(CodeBlockBackgroundColor) ? "#0D1117" : CodeBlockBackgroundColor;
            var codeBlockTextColor = string.IsNullOrWhiteSpace(CodeBlockTextColor) ? "#C9D1D9" : CodeBlockTextColor;
            var inlineCodeBackgroundColor = string.IsNullOrWhiteSpace(InlineCodeBackgroundColor) ? "#F1F1F1" : InlineCodeBackgroundColor;
            var quoteBlockBackgroundColor = string.IsNullOrWhiteSpace(QuoteBlockBackgroundColor) ? "#E8F5E9" : QuoteBlockBackgroundColor;

            return new AppSettings
            {
                Debug = new DebugSettings
                {
                    ShowDebugButtons = ShowDebugButtons,
                    ExportPdf = ExportPdf,
                    DebugOutputPath = debugOutputPath
                },
                CodeBlock = new CodeBlockStyleSettings
                {
                    BackgroundColor = codeBlockBackgroundColor,
                    TextColor = codeBlockTextColor,
                    FontFamily = codeBlockFontFamily,
                    FontSize = CodeBlockFontSize,
                    SpaceBetween = CodeBlockSpaceBetween,
                    EnableSyntaxHighlight = EnableSyntaxHighlight
                },
                InlineCode = new InlineCodeStyleSettings
                {
                    BackgroundColor = inlineCodeBackgroundColor,
                    FontFamily = inlineCodeFontFamily
                },
                QuoteBlock = new QuoteBlockStyleSettings
                {
                    BackgroundColor = quoteBlockBackgroundColor
                },
                Headings = new HeadingStyleSettings
                {
                    H1FontSize = H1FontSize,
                    H2FontSize = H2FontSize,
                    H3FontSize = H3FontSize,
                    H4FontSize = H4FontSize,
                    H5FontSize = H5FontSize,
                    H6FontSize = H6FontSize
                },
                Layout = CloneLayoutSettings(_layoutSettings),
                Image = CloneImageSettings(_imageSettings),
                Language = SelectedLanguage?.Code == "auto" ? string.Empty : SelectedLanguage?.Code
            };
        }

        private void EnsureFontAvailable(string fontFamily)
        {
            if (string.IsNullOrWhiteSpace(fontFamily))
                return;

            foreach (var font in AvailableFonts)
            {
                if (string.Equals(font, fontFamily, StringComparison.OrdinalIgnoreCase))
                    return;
            }

            AvailableFonts.Add(fontFamily);
        }

        private static LayoutSettings CloneLayoutSettings(LayoutSettings source)
        {
            if (source == null)
                return new LayoutSettings();

            return new LayoutSettings
            {
                ListIndent = source.ListIndent,
                TableColumnWidth = source.TableColumnWidth,
                ParagraphSpaceBefore = source.ParagraphSpaceBefore,
                ParagraphSpaceAfter = source.ParagraphSpaceAfter
            };
        }

        private static ImageSettings CloneImageSettings(ImageSettings source)
        {
            if (source == null)
                return new ImageSettings();

            return new ImageSettings
            {
                DownloadTimeoutSeconds = source.DownloadTimeoutSeconds,
                MaxFileSizeBytes = source.MaxFileSizeBytes
            };
        }

        private void ResetToDefaults()
        {
            LoadFromSettings(AppSettings.CreateDefault());
        }

        private static Color HexToColor(string hex)
        {
            try
            {
                hex = hex?.TrimStart('#') ?? "FFFFFF";
                if (hex.Length == 6)
                {
                    return Color.FromRgb(
                        Convert.ToByte(hex.Substring(0, 2), 16),
                        Convert.ToByte(hex.Substring(2, 2), 16),
                        Convert.ToByte(hex.Substring(4, 2), 16));
                }
                if (hex.Length == 8)
                {
                    return Color.FromArgb(
                        Convert.ToByte(hex.Substring(0, 2), 16),
                        Convert.ToByte(hex.Substring(2, 2), 16),
                        Convert.ToByte(hex.Substring(4, 2), 16),
                        Convert.ToByte(hex.Substring(6, 2), 16));
                }
            }
            catch { }
            return Colors.White;
        }

        public class LanguageOption
        {
            public LanguageOption(string code, string displayName)
            {
                Code = code;
                DisplayName = displayName;
            }

            public string Code { get; }
            public string DisplayName { get; }
        }
    }
}
