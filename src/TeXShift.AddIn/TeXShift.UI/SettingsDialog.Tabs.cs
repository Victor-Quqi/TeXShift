using System.Drawing;
using System.Windows.Forms;
using TeXShift.AddIn.Localization;
using TeXShift.Core.Configuration;
using TeXShift.Core.Localization;

namespace TeXShift.AddIn.UI
{
    /// <summary>
    /// Tab creation methods for SettingsDialog.
    /// </summary>
    public partial class SettingsDialog
    {
        private void CreateStyleTab()
        {
            var tab = new TabPage(UIResources.GetString("Settings_Tab_Style"));

            int y = 20;

            // Quote block section
            var quoteLabel = new Label
            {
                Text = $"── {UIResources.GetString("Settings_Section_QuoteBlock")} ──",
                Location = new Point(20, y),
                AutoSize = true,
                ForeColor = Color.Gray
            };
            tab.Controls.Add(quoteLabel);
            y += 30;

            var quoteBgLabel = new Label { Text = UIResources.GetString("Settings_Label_BackgroundColor") + ":", Location = new Point(20, y + 4), AutoSize = true };
            _quoteBlockBgColorPanel = CreateColorPanel(new Point(120, y));
            _quoteBlockBgColorButton = new Button { Text = UIResources.GetString("Settings_Button_Select"), Location = new Point(165, y), Size = new Size(60, 24) };
            _quoteBlockBgColorButton.Click += (s, e) => PickColor(_quoteBlockBgColorPanel);
            tab.Controls.Add(quoteBgLabel);
            tab.Controls.Add(_quoteBlockBgColorPanel);
            tab.Controls.Add(_quoteBlockBgColorButton);
            y += 45;

            // Heading section
            var headingLabel = new Label
            {
                Text = $"── {UIResources.GetString("Settings_Section_HeadingSize")} ──",
                Location = new Point(20, y),
                AutoSize = true,
                ForeColor = Color.Gray
            };
            tab.Controls.Add(headingLabel);
            y += 30;

            _headingFontSizeNumerics = new NumericUpDown[6];
            for (int i = 0; i < 6; i++)
            {
                var label = new Label
                {
                    Text = $"H{i + 1}:",
                    Location = new Point(20 + (i % 3) * 140, y + (i / 3) * 35 + 4),
                    AutoSize = true
                };

                _headingFontSizeNumerics[i] = new NumericUpDown
                {
                    Location = new Point(55 + (i % 3) * 140, y + (i / 3) * 35),
                    Size = new Size(70, 23),
                    Minimum = 8,
                    Maximum = 36,
                    DecimalPlaces = 1,
                    Increment = 0.5m
                };

                tab.Controls.Add(label);
                tab.Controls.Add(_headingFontSizeNumerics[i]);
            }

            _tabControl.TabPages.Add(tab);
        }

        private void CreateCodeBlockTab()
        {
            var tab = new TabPage(UIResources.GetString("Settings_Tab_CodeBlock"));

            int y = 20;

            // Background color
            var bgLabel = new Label { Text = UIResources.GetString("Settings_Label_BackgroundColor") + ":", Location = new Point(20, y + 4), AutoSize = true };
            _codeBlockBgColorPanel = CreateColorPanel(new Point(120, y));
            _codeBlockBgColorButton = new Button { Text = UIResources.GetString("Settings_Button_Select"), Location = new Point(165, y), Size = new Size(60, 24) };
            _codeBlockBgColorButton.Click += (s, e) => PickColor(_codeBlockBgColorPanel);
            tab.Controls.Add(bgLabel);
            tab.Controls.Add(_codeBlockBgColorPanel);
            tab.Controls.Add(_codeBlockBgColorButton);
            y += 35;

            // Text color
            var textLabel = new Label { Text = UIResources.GetString("Settings_Label_TextColor") + ":", Location = new Point(20, y + 4), AutoSize = true };
            _codeBlockTextColorPanel = CreateColorPanel(new Point(120, y));
            _codeBlockTextColorButton = new Button { Text = UIResources.GetString("Settings_Button_Select"), Location = new Point(165, y), Size = new Size(60, 24) };
            _codeBlockTextColorButton.Click += (s, e) => PickColor(_codeBlockTextColorPanel);
            tab.Controls.Add(textLabel);
            tab.Controls.Add(_codeBlockTextColorPanel);
            tab.Controls.Add(_codeBlockTextColorButton);
            y += 35;

            // Font family
            var fontLabel = new Label { Text = UIResources.GetString("Settings_Label_Font") + ":", Location = new Point(20, y + 4), AutoSize = true };
            _codeBlockFontComboBox = new ComboBox
            {
                Location = new Point(120, y),
                Size = new Size(150, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _codeBlockFontComboBox.Items.AddRange(FontPresets.MonospaceFonts);
            tab.Controls.Add(fontLabel);
            tab.Controls.Add(_codeBlockFontComboBox);
            y += 35;

            // Font size
            var sizeLabel = new Label { Text = $"{UIResources.GetString("Settings_Label_FontSize")} (pt):", Location = new Point(20, y + 4), AutoSize = true };
            _codeBlockFontSizeNumeric = new NumericUpDown
            {
                Location = new Point(120, y),
                Size = new Size(80, 23),
                Minimum = 8,
                Maximum = 24,
                DecimalPlaces = 1,
                Increment = 0.5m
            };
            tab.Controls.Add(sizeLabel);
            tab.Controls.Add(_codeBlockFontSizeNumeric);
            y += 35;

            // Space between lines
            var spaceBetweenLabel = new Label { Text = $"{UIResources.GetString("Settings_Label_LineSpacing")} (pt):", Location = new Point(20, y + 4), AutoSize = true };
            _codeBlockSpaceBetweenNumeric = new NumericUpDown
            {
                Location = new Point(120, y),
                Size = new Size(80, 23),
                Minimum = 12,
                Maximum = 36,
                DecimalPlaces = 1,
                Increment = 0.5m
            };
            tab.Controls.Add(spaceBetweenLabel);
            tab.Controls.Add(_codeBlockSpaceBetweenNumeric);
            y += 35;

            // Syntax highlight
            _enableSyntaxHighlightCheckBox = new CheckBox
            {
                Text = UIResources.GetString("Settings_Checkbox_EnableSyntaxHighlight"),
                Location = new Point(20, y),
                AutoSize = true
            };
            tab.Controls.Add(_enableSyntaxHighlightCheckBox);
            y += 40;

            // Inline code section
            var inlineLabel = new Label
            {
                Text = $"── {UIResources.GetString("Settings_Section_InlineCode")} ──",
                Location = new Point(20, y),
                AutoSize = true,
                ForeColor = Color.Gray
            };
            tab.Controls.Add(inlineLabel);
            y += 30;

            // Inline code background
            var inlineBgLabel = new Label { Text = UIResources.GetString("Settings_Label_BackgroundColor") + ":", Location = new Point(20, y + 4), AutoSize = true };
            _inlineCodeBgColorPanel = CreateColorPanel(new Point(120, y));
            _inlineCodeBgColorButton = new Button { Text = UIResources.GetString("Settings_Button_Select"), Location = new Point(165, y), Size = new Size(60, 24) };
            _inlineCodeBgColorButton.Click += (s, e) => PickColor(_inlineCodeBgColorPanel);
            tab.Controls.Add(inlineBgLabel);
            tab.Controls.Add(_inlineCodeBgColorPanel);
            tab.Controls.Add(_inlineCodeBgColorButton);
            y += 35;

            // Inline code font
            var inlineFontLabel = new Label { Text = UIResources.GetString("Settings_Label_Font") + ":", Location = new Point(20, y + 4), AutoSize = true };
            _inlineCodeFontComboBox = new ComboBox
            {
                Location = new Point(120, y),
                Size = new Size(150, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _inlineCodeFontComboBox.Items.AddRange(FontPresets.MonospaceFonts);
            tab.Controls.Add(inlineFontLabel);
            tab.Controls.Add(_inlineCodeFontComboBox);

            _tabControl.TabPages.Add(tab);
        }

        private void CreateMermaidTab()
        {
            var tab = new TabPage(UIResources.GetString("Settings_Tab_Mermaid"));

            int y = 20;

            // Mermaid section
            var mermaidLabel = new Label
            {
                Text = $"── {UIResources.GetString("Settings_Section_Mermaid")} ──",
                Location = new Point(20, y),
                AutoSize = true,
                ForeColor = Color.Gray
            };
            tab.Controls.Add(mermaidLabel);
            y += 30;

            // Theme
            var themeLabel = new Label { Text = UIResources.GetString("Settings_Label_Theme") + ":", Location = new Point(20, y + 4), AutoSize = true };
            _mermaidThemeComboBox = new ComboBox
            {
                Location = new Point(120, y),
                Size = new Size(150, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _mermaidThemeComboBox.Items.AddRange(new[] { "default", "dark", "forest", "neutral" });
            tab.Controls.Add(themeLabel);
            tab.Controls.Add(_mermaidThemeComboBox);
            y += 35;

            // Max Width
            var maxWidthLabel = new Label { Text = UIResources.GetString("Settings_Label_MaxWidth") + ":", Location = new Point(20, y + 4), AutoSize = true };
            _mermaidMaxWidthNumeric = new NumericUpDown
            {
                Location = new Point(120, y),
                Size = new Size(100, 23),
                Minimum = 640,
                Maximum = 3840,
                Increment = 64
            };
            tab.Controls.Add(maxWidthLabel);
            tab.Controls.Add(_mermaidMaxWidthNumeric);
            y += 35;

            // Max Height
            var maxHeightLabel = new Label { Text = UIResources.GetString("Settings_Label_MaxHeight") + ":", Location = new Point(20, y + 4), AutoSize = true };
            _mermaidMaxHeightNumeric = new NumericUpDown
            {
                Location = new Point(120, y),
                Size = new Size(100, 23),
                Minimum = 480,
                Maximum = 2160,
                Increment = 48
            };
            tab.Controls.Add(maxHeightLabel);
            tab.Controls.Add(_mermaidMaxHeightNumeric);

            _tabControl.TabPages.Add(tab);
        }

        private void CreateDebugTab()
        {
            var tab = new TabPage(UIResources.GetString("Settings_Tab_Debug"));

            int y = 20;

            // Language section
            var languageLabel = new Label
            {
                Text = $"── {Resources.GetString("UI_LanguageTitle")} ──",
                Location = new Point(20, y),
                AutoSize = true,
                ForeColor = Color.Gray
            };
            tab.Controls.Add(languageLabel);
            y += 30;

            var langLabel = new Label { Text = Resources.GetString("UI_LanguageLabel") + ":", Location = new Point(20, y + 4), AutoSize = true };
            _languageComboBox = new ComboBox
            {
                Location = new Point(120, y),
                Size = new Size(150, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _languageComboBox.Items.Add(new LanguageItem("auto", Resources.GetString("UI_Language_System")));
            _languageComboBox.Items.Add(new LanguageItem("zh-CN", Resources.GetString("UI_Language_Chinese")));
            _languageComboBox.Items.Add(new LanguageItem("en-US", Resources.GetString("UI_Language_English")));
            tab.Controls.Add(langLabel);
            tab.Controls.Add(_languageComboBox);
            y += 45;

            // Debug section
            var debugLabel = new Label
            {
                Text = $"── {UIResources.GetString("Settings_Tab_Debug")} ──",
                Location = new Point(20, y),
                AutoSize = true,
                ForeColor = Color.Gray
            };
            tab.Controls.Add(debugLabel);
            y += 30;

            // Show debug buttons
            _showDebugButtonsCheckBox = new CheckBox
            {
                Text = UIResources.GetString("Settings_Checkbox_ShowDebugButtons"),
                Location = new Point(20, y),
                Size = new Size(400, 24),
                AutoSize = true
            };
            tab.Controls.Add(_showDebugButtonsCheckBox);
            y += 30;

            // Export PDF
            _exportPdfCheckBox = new CheckBox
            {
                Text = UIResources.GetString("Settings_Checkbox_ExportPdf"),
                Location = new Point(20, y),
                Size = new Size(400, 24),
                AutoSize = true
            };
            tab.Controls.Add(_exportPdfCheckBox);
            y += 35;

            // Debug output path
            var debugPathLabel = new Label
            {
                Text = UIResources.GetString("Settings_Description_DebugOutputPath"),
                Location = new Point(20, y),
                AutoSize = true
            };
            tab.Controls.Add(debugPathLabel);
            y += 25;

            _debugOutputPathTextBox = new TextBox
            {
                Location = new Point(20, y),
                Size = new Size(340, 23)
            };

            _browseDebugPathButton = new Button
            {
                Text = UIResources.GetString("Settings_Button_Browse"),
                Location = new Point(365, y - 1),
                Size = new Size(70, 25)
            };
            _browseDebugPathButton.Click += BrowseDebugPath_Click;

            tab.Controls.Add(_debugOutputPathTextBox);
            tab.Controls.Add(_browseDebugPathButton);

            _tabControl.TabPages.Add(tab);
        }

        private Panel CreateColorPanel(Point location)
        {
            return new Panel
            {
                Location = location,
                Size = new Size(40, 24),
                BorderStyle = BorderStyle.FixedSingle
            };
        }

        private void PickColor(Panel colorPanel)
        {
            using (var dialog = new ColorDialog())
            {
                dialog.Color = colorPanel.BackColor;
                dialog.FullOpen = true;
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    colorPanel.BackColor = dialog.Color;
                }
            }
        }
    }
}
