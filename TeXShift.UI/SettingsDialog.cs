using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using TeXShift.Core.Configuration;

namespace TeXShift.UI
{
    /// <summary>
    /// Settings dialog for TeXShift configuration.
    /// Uses TabControl to organize settings into categories.
    /// </summary>
    public class SettingsDialog : Form
    {
        private readonly AppSettings _originalSettings;
        private AppSettings _currentSettings;

        // Tab control
        private TabControl _tabControl;

        // Debug settings
        private CheckBox _showDebugButtonsCheckBox;
        private TextBox _debugOutputPathTextBox;
        private Button _browseDebugPathButton;

        // Code block settings
        private Panel _codeBlockBgColorPanel;
        private Button _codeBlockBgColorButton;
        private Panel _codeBlockTextColorPanel;
        private Button _codeBlockTextColorButton;
        private ComboBox _codeBlockFontComboBox;
        private NumericUpDown _codeBlockFontSizeNumeric;
        private CheckBox _enableSyntaxHighlightCheckBox;

        // Inline code settings
        private Panel _inlineCodeBgColorPanel;
        private Button _inlineCodeBgColorButton;
        private ComboBox _inlineCodeFontComboBox;

        // Quote block settings
        private Panel _quoteBlockBgColorPanel;
        private Button _quoteBlockBgColorButton;

        // Heading settings
        private NumericUpDown[] _headingFontSizeNumerics;

        // Buttons
        private Button _okButton;
        private Button _cancelButton;
        private Button _resetButton;

        public SettingsDialog(AppSettings settings)
        {
            _originalSettings = settings ?? AppSettings.CreateDefault();
            _currentSettings = CloneSettings(_originalSettings);

            InitializeComponent();
            LoadSettingsToControls();
        }

        /// <summary>
        /// Gets the updated settings after the dialog is closed with OK.
        /// </summary>
        public AppSettings GetUpdatedSettings()
        {
            return _currentSettings;
        }

        private void InitializeComponent()
        {
            this.Text = "TeXShift 设置";
            this.Size = new Size(500, 520);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Font = new Font("Microsoft YaHei UI", 9F);

            // Create tab control
            _tabControl = new TabControl
            {
                Location = new Point(12, 12),
                Size = new Size(460, 420)
            };

            // Create tabs
            CreateStyleTab();
            CreateCodeBlockTab();
            CreateDebugTab();

            this.Controls.Add(_tabControl);

            // Create buttons
            _okButton = new Button
            {
                Text = "确定",
                Location = new Point(216, 445),
                Size = new Size(80, 28),
                DialogResult = DialogResult.OK
            };
            _okButton.Click += OkButton_Click;

            _cancelButton = new Button
            {
                Text = "取消",
                Location = new Point(302, 445),
                Size = new Size(80, 28),
                DialogResult = DialogResult.Cancel
            };

            _resetButton = new Button
            {
                Text = "恢复默认",
                Location = new Point(388, 445),
                Size = new Size(80, 28)
            };
            _resetButton.Click += ResetButton_Click;

            this.Controls.Add(_okButton);
            this.Controls.Add(_cancelButton);
            this.Controls.Add(_resetButton);

            this.AcceptButton = _okButton;
            this.CancelButton = _cancelButton;
        }

        private void CreateDebugTab()
        {
            var tab = new TabPage("调试");

            // Show debug buttons
            _showDebugButtonsCheckBox = new CheckBox
            {
                Text = "显示调试按钮 (调试转换、查看XML)",
                Location = new Point(20, 25),
                Size = new Size(400, 24),
                AutoSize = true
            };

            // Debug output path
            var debugPathLabel = new Label
            {
                Text = "调试输出目录 (留空使用默认):",
                Location = new Point(20, 65),
                AutoSize = true
            };

            _debugOutputPathTextBox = new TextBox
            {
                Location = new Point(20, 90),
                Size = new Size(340, 23)
            };

            _browseDebugPathButton = new Button
            {
                Text = "浏览...",
                Location = new Point(365, 89),
                Size = new Size(70, 25)
            };
            _browseDebugPathButton.Click += BrowseDebugPath_Click;

            tab.Controls.Add(_showDebugButtonsCheckBox);
            tab.Controls.Add(debugPathLabel);
            tab.Controls.Add(_debugOutputPathTextBox);
            tab.Controls.Add(_browseDebugPathButton);

            _tabControl.TabPages.Add(tab);
        }

        private void CreateCodeBlockTab()
        {
            var tab = new TabPage("代码块");

            int y = 20;

            // Background color
            var bgLabel = new Label { Text = "背景颜色:", Location = new Point(20, y + 4), AutoSize = true };
            _codeBlockBgColorPanel = CreateColorPanel(new Point(120, y));
            _codeBlockBgColorButton = new Button { Text = "选择...", Location = new Point(165, y), Size = new Size(60, 24) };
            _codeBlockBgColorButton.Click += (s, e) => PickColor(_codeBlockBgColorPanel);
            tab.Controls.Add(bgLabel);
            tab.Controls.Add(_codeBlockBgColorPanel);
            tab.Controls.Add(_codeBlockBgColorButton);
            y += 35;

            // Text color
            var textLabel = new Label { Text = "文字颜色:", Location = new Point(20, y + 4), AutoSize = true };
            _codeBlockTextColorPanel = CreateColorPanel(new Point(120, y));
            _codeBlockTextColorButton = new Button { Text = "选择...", Location = new Point(165, y), Size = new Size(60, 24) };
            _codeBlockTextColorButton.Click += (s, e) => PickColor(_codeBlockTextColorPanel);
            tab.Controls.Add(textLabel);
            tab.Controls.Add(_codeBlockTextColorPanel);
            tab.Controls.Add(_codeBlockTextColorButton);
            y += 35;

            // Font family
            var fontLabel = new Label { Text = "字体:", Location = new Point(20, y + 4), AutoSize = true };
            _codeBlockFontComboBox = new ComboBox
            {
                Location = new Point(120, y),
                Size = new Size(150, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _codeBlockFontComboBox.Items.AddRange(new[] { "Consolas", "Courier New", "Source Code Pro", "Fira Code", "JetBrains Mono" });
            tab.Controls.Add(fontLabel);
            tab.Controls.Add(_codeBlockFontComboBox);
            y += 35;

            // Font size
            var sizeLabel = new Label { Text = "字号 (pt):", Location = new Point(20, y + 4), AutoSize = true };
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

            // Syntax highlight
            _enableSyntaxHighlightCheckBox = new CheckBox
            {
                Text = "启用语法高亮",
                Location = new Point(20, y),
                AutoSize = true
            };
            tab.Controls.Add(_enableSyntaxHighlightCheckBox);
            y += 40;

            // Inline code section
            var inlineLabel = new Label
            {
                Text = "── 内联代码 ──",
                Location = new Point(20, y),
                AutoSize = true,
                ForeColor = Color.Gray
            };
            tab.Controls.Add(inlineLabel);
            y += 30;

            // Inline code background
            var inlineBgLabel = new Label { Text = "背景颜色:", Location = new Point(20, y + 4), AutoSize = true };
            _inlineCodeBgColorPanel = CreateColorPanel(new Point(120, y));
            _inlineCodeBgColorButton = new Button { Text = "选择...", Location = new Point(165, y), Size = new Size(60, 24) };
            _inlineCodeBgColorButton.Click += (s, e) => PickColor(_inlineCodeBgColorPanel);
            tab.Controls.Add(inlineBgLabel);
            tab.Controls.Add(_inlineCodeBgColorPanel);
            tab.Controls.Add(_inlineCodeBgColorButton);
            y += 35;

            // Inline code font
            var inlineFontLabel = new Label { Text = "字体:", Location = new Point(20, y + 4), AutoSize = true };
            _inlineCodeFontComboBox = new ComboBox
            {
                Location = new Point(120, y),
                Size = new Size(150, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _inlineCodeFontComboBox.Items.AddRange(new[] { "Consolas", "Courier New", "Source Code Pro", "Fira Code", "JetBrains Mono" });
            tab.Controls.Add(inlineFontLabel);
            tab.Controls.Add(_inlineCodeFontComboBox);

            _tabControl.TabPages.Add(tab);
        }

        private void CreateStyleTab()
        {
            var tab = new TabPage("样式");

            int y = 20;

            // Quote block section
            var quoteLabel = new Label
            {
                Text = "── 引用块 ──",
                Location = new Point(20, y),
                AutoSize = true,
                ForeColor = Color.Gray
            };
            tab.Controls.Add(quoteLabel);
            y += 30;

            var quoteBgLabel = new Label { Text = "背景颜色:", Location = new Point(20, y + 4), AutoSize = true };
            _quoteBlockBgColorPanel = CreateColorPanel(new Point(120, y));
            _quoteBlockBgColorButton = new Button { Text = "选择...", Location = new Point(165, y), Size = new Size(60, 24) };
            _quoteBlockBgColorButton.Click += (s, e) => PickColor(_quoteBlockBgColorPanel);
            tab.Controls.Add(quoteBgLabel);
            tab.Controls.Add(_quoteBlockBgColorPanel);
            tab.Controls.Add(_quoteBlockBgColorButton);
            y += 45;

            // Heading section
            var headingLabel = new Label
            {
                Text = "── 标题字号 (pt) ──",
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

        private void LoadSettingsToControls()
        {
            // Debug settings
            _showDebugButtonsCheckBox.Checked = _currentSettings.Debug.ShowDebugButtons;
            _debugOutputPathTextBox.Text = _currentSettings.Debug.DebugOutputPath;

            // Code block settings
            _codeBlockBgColorPanel.BackColor = ColorFromHex(_currentSettings.CodeBlock.BackgroundColor);
            _codeBlockTextColorPanel.BackColor = ColorFromHex(_currentSettings.CodeBlock.TextColor);
            SelectOrAddItem(_codeBlockFontComboBox, _currentSettings.CodeBlock.FontFamily);
            _codeBlockFontSizeNumeric.Value = (decimal)_currentSettings.CodeBlock.FontSize;
            _enableSyntaxHighlightCheckBox.Checked = _currentSettings.CodeBlock.EnableSyntaxHighlight;

            // Inline code settings
            _inlineCodeBgColorPanel.BackColor = ColorFromHex(_currentSettings.InlineCode.BackgroundColor);
            SelectOrAddItem(_inlineCodeFontComboBox, _currentSettings.InlineCode.FontFamily);

            // Quote block settings
            _quoteBlockBgColorPanel.BackColor = ColorFromHex(_currentSettings.QuoteBlock.BackgroundColor);

            // Heading settings
            for (int i = 0; i < 6; i++)
            {
                _headingFontSizeNumerics[i].Value = (decimal)_currentSettings.Headings.GetFontSize(i + 1);
            }
        }

        private void SaveControlsToSettings()
        {
            // Debug settings
            _currentSettings.Debug.ShowDebugButtons = _showDebugButtonsCheckBox.Checked;
            _currentSettings.Debug.DebugOutputPath = _debugOutputPathTextBox.Text.Trim();

            // Code block settings
            _currentSettings.CodeBlock.BackgroundColor = ColorToHex(_codeBlockBgColorPanel.BackColor);
            _currentSettings.CodeBlock.TextColor = ColorToHex(_codeBlockTextColorPanel.BackColor);
            _currentSettings.CodeBlock.FontFamily = _codeBlockFontComboBox.Text;
            _currentSettings.CodeBlock.FontSize = (double)_codeBlockFontSizeNumeric.Value;
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
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            SaveControlsToSettings();
        }

        private void ResetButton_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show(
                "确定要将所有设置恢复为默认值吗？",
                "确认恢复默认",
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
                    dialog.Description = "选择调试输出目录";
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

        private static AppSettings CloneSettings(AppSettings source)
        {
            return new AppSettings
            {
                Debug = new DebugSettings
                {
                    ShowDebugButtons = source.Debug.ShowDebugButtons,
                    DebugOutputPath = source.Debug.DebugOutputPath
                },
                CodeBlock = new CodeBlockStyleSettings
                {
                    BackgroundColor = source.CodeBlock.BackgroundColor,
                    TextColor = source.CodeBlock.TextColor,
                    FontFamily = source.CodeBlock.FontFamily,
                    FontSize = source.CodeBlock.FontSize,
                    LineHeight = source.CodeBlock.LineHeight,
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
                }
            };
        }
    }
}
