using System.Drawing;
using System.Windows.Forms;
using TeXShift.AddIn.Localization;
using TeXShift.Core.Configuration;

namespace TeXShift.AddIn.UI
{
    /// <summary>
    /// Settings dialog for TeXShift configuration (WinForms fallback).
    /// Uses TabControl to organize settings into categories.
    /// </summary>
    /// <remarks>
    /// This is a partial class split into:
    /// - SettingsDialog.cs: Fields, constructor, InitializeComponent
    /// - SettingsDialog.Tabs.cs: Tab creation methods
    /// - SettingsDialog.Settings.cs: Settings load/save, helpers, event handlers
    /// </remarks>
    public partial class SettingsDialog : Form
    {
        #region Fields

        private readonly AppSettings _originalSettings;
        private AppSettings _currentSettings;

        // Tab control
        private TabControl _tabControl;

        // Debug settings
        private CheckBox _showDebugButtonsCheckBox;
        private CheckBox _exportPdfCheckBox;
        private TextBox _debugOutputPathTextBox;
        private Button _browseDebugPathButton;

        // Code block settings
        private Panel _codeBlockBgColorPanel;
        private Button _codeBlockBgColorButton;
        private Panel _codeBlockTextColorPanel;
        private Button _codeBlockTextColorButton;
        private ComboBox _codeBlockFontComboBox;
        private NumericUpDown _codeBlockFontSizeNumeric;
        private NumericUpDown _codeBlockSpaceBetweenNumeric;
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

        // Mermaid settings
        private ComboBox _mermaidThemeComboBox;
        private NumericUpDown _mermaidMaxWidthNumeric;
        private NumericUpDown _mermaidMaxHeightNumeric;

        // Language settings
        private ComboBox _languageComboBox;

        // Buttons
        private Button _okButton;
        private Button _cancelButton;
        private Button _resetButton;

        #endregion

        #region Constructor

        public SettingsDialog(AppSettings settings)
        {
            _originalSettings = settings ?? AppSettings.CreateDefault();
            _currentSettings = CloneSettings(_originalSettings);

            InitializeComponent();
            LoadSettingsToControls();
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Gets the updated settings after the dialog is closed with OK.
        /// </summary>
        public AppSettings GetUpdatedSettings()
        {
            return _currentSettings;
        }

        #endregion

        #region Initialization

        private void InitializeComponent()
        {
            this.Text = UIResources.GetString("Settings_Title");
            this.Size = new Size(500, 560);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Font = new Font("Microsoft YaHei UI", 9F);

            // Create tab control
            _tabControl = new TabControl
            {
                Location = new Point(12, 12),
                Size = new Size(460, 460)
            };

            // Create tabs (defined in SettingsDialog.Tabs.cs)
            CreateStyleTab();
            CreateCodeBlockTab();
            CreateMermaidTab();
            CreateDebugTab();

            this.Controls.Add(_tabControl);

            // Create buttons
            _okButton = new Button
            {
                Text = UIResources.GetString("Settings_Button_Ok"),
                Location = new Point(216, 485),
                Size = new Size(80, 28),
                DialogResult = DialogResult.OK
            };
            _okButton.Click += OkButton_Click;

            _cancelButton = new Button
            {
                Text = UIResources.GetString("Settings_Button_Cancel"),
                Location = new Point(302, 485),
                Size = new Size(80, 28),
                DialogResult = DialogResult.Cancel
            };

            _resetButton = new Button
            {
                Text = UIResources.GetString("Settings_Button_ResetDefaults"),
                Location = new Point(388, 485),
                Size = new Size(80, 28)
            };
            _resetButton.Click += ResetButton_Click;

            this.Controls.Add(_okButton);
            this.Controls.Add(_cancelButton);
            this.Controls.Add(_resetButton);

            this.AcceptButton = _okButton;
            this.CancelButton = _cancelButton;
        }

        #endregion
    }
}
