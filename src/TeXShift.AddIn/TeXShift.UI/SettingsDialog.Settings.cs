using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using TeXShift.AddIn.Localization;
using TeXShift.Core.Configuration;

namespace TeXShift.AddIn.UI
{
    /// <summary>
    /// Settings load/save and helper methods for SettingsDialog.
    /// </summary>
    public partial class SettingsDialog
    {
        private void LoadSettingsToControls()
        {
            // Debug settings
            _showDebugButtonsCheckBox.Checked = _currentSettings.Debug.ShowDebugButtons;
            _exportPdfCheckBox.Checked = _currentSettings.Debug.ExportPdf;
            _debugOutputPathTextBox.Text = _currentSettings.Debug.DebugOutputPath;

            // Code block settings
            _codeBlockBgColorPanel.BackColor = ColorFromHex(_currentSettings.CodeBlock.BackgroundColor);
            _codeBlockTextColorPanel.BackColor = ColorFromHex(_currentSettings.CodeBlock.TextColor);
            SelectOrAddItem(_codeBlockFontComboBox, _currentSettings.CodeBlock.FontFamily);
            SafeSetNumericValue(_codeBlockFontSizeNumeric, (decimal)_currentSettings.CodeBlock.FontSize);
            SafeSetNumericValue(_codeBlockSpaceBetweenNumeric, (decimal)_currentSettings.CodeBlock.SpaceBetween);
            _enableSyntaxHighlightCheckBox.Checked = _currentSettings.CodeBlock.EnableSyntaxHighlight;

            // Inline code settings
            _inlineCodeBgColorPanel.BackColor = ColorFromHex(_currentSettings.InlineCode.BackgroundColor);
            SelectOrAddItem(_inlineCodeFontComboBox, _currentSettings.InlineCode.FontFamily);

            // Quote block settings
            _quoteBlockBgColorPanel.BackColor = ColorFromHex(_currentSettings.QuoteBlock.BackgroundColor);

            // Heading settings
            for (int i = 0; i < 6; i++)
            {
                SafeSetNumericValue(_headingFontSizeNumerics[i], (decimal)_currentSettings.Headings.GetFontSize(i + 1));
            }

            // Mermaid settings
            var mermaidTheme = string.IsNullOrWhiteSpace(_currentSettings.Mermaid?.Theme) ? "default" : _currentSettings.Mermaid.Theme;
            SelectOrAddItem(_mermaidThemeComboBox, mermaidTheme);
            SafeSetNumericValue(_mermaidMaxWidthNumeric, _currentSettings.Mermaid?.MaxWidth ?? 1920);
            SafeSetNumericValue(_mermaidMaxHeightNumeric, _currentSettings.Mermaid?.MaxHeight ?? 1080);

            // Language settings
            var languageCode = string.IsNullOrWhiteSpace(_currentSettings.Language) ? "auto" : _currentSettings.Language;
            for (int i = 0; i < _languageComboBox.Items.Count; i++)
            {
                if (_languageComboBox.Items[i] is LanguageItem item && string.Equals(item.Code, languageCode, StringComparison.OrdinalIgnoreCase))
                {
                    _languageComboBox.SelectedIndex = i;
                    break;
                }
            }
            if (_languageComboBox.SelectedIndex < 0 && _languageComboBox.Items.Count > 0)
                _languageComboBox.SelectedIndex = 0;
        }

        private void SaveControlsToSettings()
        {
            // Debug settings
            _currentSettings.Debug.ShowDebugButtons = _showDebugButtonsCheckBox.Checked;
            _currentSettings.Debug.ExportPdf = _exportPdfCheckBox.Checked;
            _currentSettings.Debug.DebugOutputPath = _debugOutputPathTextBox.Text.Trim();

            // Code block settings
            _currentSettings.CodeBlock.BackgroundColor = ColorToHex(_codeBlockBgColorPanel.BackColor);
            _currentSettings.CodeBlock.TextColor = ColorToHex(_codeBlockTextColorPanel.BackColor);
            _currentSettings.CodeBlock.FontFamily = _codeBlockFontComboBox.Text;
            _currentSettings.CodeBlock.FontSize = (double)_codeBlockFontSizeNumeric.Value;
            _currentSettings.CodeBlock.SpaceBetween = (double)_codeBlockSpaceBetweenNumeric.Value;
            _currentSettings.CodeBlock.EnableSyntaxHighlight = _enableSyntaxHighlightCheckBox.Checked;

            // Inline code settings
            _currentSettings.InlineCode.BackgroundColor = ColorToHex(_inlineCodeBgColorPanel.BackColor);
            _currentSettings.InlineCode.FontFamily = _inlineCodeFontComboBox.Text;

            // Quote block settings
            _currentSettings.QuoteBlock.BackgroundColor = ColorToHex(_quoteBlockBgColorPanel.BackColor);

            // Heading settings
            for (int i = 0; i < 6; i++)
            {
                _currentSettings.Headings.SetFontSize(i + 1, (double)_headingFontSizeNumerics[i].Value);
            }

            // Mermaid settings
            if (_currentSettings.Mermaid == null)
                _currentSettings.Mermaid = new MermaidSettings();
            _currentSettings.Mermaid.Theme = _mermaidThemeComboBox.Text;
            _currentSettings.Mermaid.MaxWidth = (int)_mermaidMaxWidthNumeric.Value;
            _currentSettings.Mermaid.MaxHeight = (int)_mermaidMaxHeightNumeric.Value;

            // Language settings
            var selectedLanguage = _languageComboBox.SelectedItem as LanguageItem;
            _currentSettings.Language = selectedLanguage?.Code == "auto" ? string.Empty : selectedLanguage?.Code;
        }

        private static AppSettings CloneSettings(AppSettings source)
        {
            return new AppSettings
            {
                Debug = new DebugSettings
                {
                    ShowDebugButtons = source.Debug.ShowDebugButtons,
                    ExportPdf = source.Debug.ExportPdf,
                    DebugOutputPath = source.Debug.DebugOutputPath
                },
                CodeBlock = new CodeBlockStyleSettings
                {
                    BackgroundColor = source.CodeBlock.BackgroundColor,
                    TextColor = source.CodeBlock.TextColor,
                    FontFamily = source.CodeBlock.FontFamily,
                    FontSize = source.CodeBlock.FontSize,
                    SpaceBetween = source.CodeBlock.SpaceBetween,
                    EnableSyntaxHighlight = source.CodeBlock.EnableSyntaxHighlight
                },
                InlineCode = new InlineCodeStyleSettings
                {
                    BackgroundColor = source.InlineCode.BackgroundColor,
                    FontFamily = source.InlineCode.FontFamily
                },
                QuoteBlock = new QuoteBlockStyleSettings
                {
                    BackgroundColor = source.QuoteBlock.BackgroundColor
                },
                Headings = new HeadingStyleSettings
                {
                    H1FontSize = source.Headings.H1FontSize,
                    H2FontSize = source.Headings.H2FontSize,
                    H3FontSize = source.Headings.H3FontSize,
                    H4FontSize = source.Headings.H4FontSize,
                    H5FontSize = source.Headings.H5FontSize,
                    H6FontSize = source.Headings.H6FontSize
                },
                Mermaid = new MermaidSettings
                {
                    Theme = source.Mermaid?.Theme ?? "default",
                    MaxWidth = source.Mermaid?.MaxWidth ?? 1920,
                    MaxHeight = source.Mermaid?.MaxHeight ?? 1080
                },
                Layout = new Core.Configuration.LayoutSettings
                {
                    ListIndent = source.Layout.ListIndent,
                    TableColumnWidth = source.Layout.TableColumnWidth,
                    ParagraphSpaceBefore = source.Layout.ParagraphSpaceBefore,
                    ParagraphSpaceAfter = source.Layout.ParagraphSpaceAfter
                },
                Image = new ImageSettings
                {
                    DownloadTimeoutSeconds = source.Image.DownloadTimeoutSeconds,
                    MaxFileSizeBytes = source.Image.MaxFileSizeBytes
                },
                Language = source.Language
            };
        }

        #region Helper Methods

        /// <summary>
        /// Safely sets a NumericUpDown value, clamping to valid range if needed.
        /// Handles legacy configs with missing or out-of-range values.
        /// </summary>
        private static void SafeSetNumericValue(NumericUpDown numeric, decimal value)
        {
            value = Math.Max(value, numeric.Minimum);
            value = Math.Min(value, numeric.Maximum);
            numeric.Value = value;
        }

        private void SelectOrAddItem(ComboBox comboBox, string value)
        {
            int index = comboBox.Items.IndexOf(value);
            if (index >= 0)
            {
                comboBox.SelectedIndex = index;
            }
            else
            {
                comboBox.Items.Add(value);
                comboBox.SelectedIndex = comboBox.Items.Count - 1;
            }
        }

        private static Color ColorFromHex(string hex)
        {
            try
            {
                if (hex.StartsWith("#"))
                    hex = hex.Substring(1);
                return ColorTranslator.FromHtml("#" + hex);
            }
            catch
            {
                return Color.White;
            }
        }

        private static string ColorToHex(Color color)
        {
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        #endregion

        #region Event Handlers

        private void OkButton_Click(object sender, EventArgs e)
        {
            SaveControlsToSettings();
        }

        private void ResetButton_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show(
                UIResources.GetString("Dialog_Confirm_ResetSettings"),
                UIResources.GetString("Dialog_Confirm_ResetTitle"),
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                _currentSettings = AppSettings.CreateDefault();
                LoadSettingsToControls();
            }
        }

        private void BrowseDebugPath_Click(object sender, EventArgs e)
        {
            string selectedPath = null;
            string initialPath = _debugOutputPathTextBox.Text;

            // FolderBrowserDialog must be run on an STA thread
            var thread = new Thread(() =>
            {
                using (var dialog = new FolderBrowserDialog())
                {
                    dialog.Description = UIResources.GetString("Dialog_Title_BrowseFolder");
                    if (!string.IsNullOrEmpty(initialPath))
                    {
                        dialog.SelectedPath = initialPath;
                    }

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        selectedPath = dialog.SelectedPath;
                    }
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (!string.IsNullOrEmpty(selectedPath))
            {
                _debugOutputPathTextBox.Text = selectedPath;
            }
        }

        #endregion

        #region Helper Classes

        /// <summary>
        /// Helper class for language combo box items.
        /// </summary>
        private class LanguageItem
        {
            public string Code { get; }
            public string DisplayName { get; }

            public LanguageItem(string code, string displayName)
            {
                Code = code;
                DisplayName = displayName;
            }

            public override string ToString() => DisplayName;
        }

        #endregion
    }
}
