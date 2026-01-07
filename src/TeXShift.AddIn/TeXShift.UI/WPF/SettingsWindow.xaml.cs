using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using TeXShift.AddIn.Localization;
using TeXShift.AddIn.UI.WPF.ViewModels;
using TeXShift.Core.Configuration;

namespace TeXShift.AddIn.UI.WPF
{
    /// <summary>
    /// Helper class to wrap a window handle for use with WinForms dialogs.
    /// </summary>
    internal class WpfWin32Window : System.Windows.Forms.IWin32Window
    {
        public IntPtr Handle { get; }
        public WpfWin32Window(Window window)
        {
            var helper = new WindowInteropHelper(window);
            if (helper.Handle == IntPtr.Zero)
            {
                helper.EnsureHandle();
            }
            Handle = helper.Handle;
        }
    }

    /// <summary>
    /// Settings window using WPF and Material Design.
    /// </summary>
    public partial class SettingsWindow : Window
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private IntPtr _ownerHandle;

        public SettingsViewModel ViewModel { get; }

        public SettingsWindow()
        {
            ViewModel = new SettingsViewModel();
            DataContext = ViewModel;
            InitializeComponent();

            this.Loaded += (s, e) =>
            {
                _ownerHandle = new WindowInteropHelper(this).Owner;
            };
            this.Closing += (s, e) =>
            {
                // Set focus back to Owner before closing to avoid focus flash
                if (_ownerHandle != IntPtr.Zero)
                {
                    SetForegroundWindow(_ownerHandle);
                }
            };
        }

        public SettingsWindow(AppSettings settings) : this()
        {
            ViewModel.LoadFromSettings(settings);
        }

        public AppSettings GetUpdatedSettings()
        {
            return ViewModel.ToSettings();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            DataContext = null;
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                UIResources.GetString("Dialog_Confirm_ResetSettings"),
                UIResources.GetString("Dialog_Confirm_ResetTitle"),
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                ViewModel.LoadFromSettings(AppSettings.CreateDefault());
            }
        }

        #region Dialog Handlers

        private void BrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            var owner = new WpfWin32Window(this);
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = UIResources.GetString("Dialog_Title_BrowseFolder");
                dialog.ShowNewFolderButton = true;
                if (!string.IsNullOrEmpty(ViewModel.DebugOutputPath))
                    dialog.SelectedPath = ViewModel.DebugOutputPath;

                if (dialog.ShowDialog(owner) == System.Windows.Forms.DialogResult.OK)
                {
                    ViewModel.DebugOutputPath = dialog.SelectedPath;
                }
            }
        }

        private void PickQuoteBlockBgColor_Click(object sender, RoutedEventArgs e)
        {
            var newColor = ShowColorDialog(ViewModel.QuoteBlockBackgroundColor);
            if (newColor != null)
            {
                ViewModel.QuoteBlockBackgroundColor = newColor;
            }
        }

        private void PickCodeBlockBgColor_Click(object sender, RoutedEventArgs e)
        {
            var newColor = ShowColorDialog(ViewModel.CodeBlockBackgroundColor);
            if (newColor != null)
            {
                ViewModel.CodeBlockBackgroundColor = newColor;
            }
        }

        private void PickCodeBlockTextColor_Click(object sender, RoutedEventArgs e)
        {
            var newColor = ShowColorDialog(ViewModel.CodeBlockTextColor);
            if (newColor != null)
            {
                ViewModel.CodeBlockTextColor = newColor;
            }
        }

        private void PickInlineCodeBgColor_Click(object sender, RoutedEventArgs e)
        {
            var newColor = ShowColorDialog(ViewModel.InlineCodeBackgroundColor);
            if (newColor != null)
            {
                ViewModel.InlineCodeBackgroundColor = newColor;
            }
        }

        private string ShowColorDialog(string currentHexColor)
        {
            var owner = new WpfWin32Window(this);
            using (var dialog = new System.Windows.Forms.ColorDialog())
            {
                dialog.FullOpen = true;
                dialog.Color = HexToDrawingColor(currentHexColor);

                if (dialog.ShowDialog(owner) == System.Windows.Forms.DialogResult.OK)
                {
                    return DrawingColorToHex(dialog.Color);
                }
            }
            return null;
        }

        private static System.Drawing.Color HexToDrawingColor(string hex)
        {
            try
            {
                hex = hex?.TrimStart('#') ?? "FFFFFF";
                if (hex.Length == 6)
                {
                    return System.Drawing.Color.FromArgb(
                        Convert.ToByte(hex.Substring(0, 2), 16),
                        Convert.ToByte(hex.Substring(2, 2), 16),
                        Convert.ToByte(hex.Substring(4, 2), 16));
                }
            }
            catch { }
            return System.Drawing.Color.White;
        }

        private static string DrawingColorToHex(System.Drawing.Color color)
        {
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        #endregion
    }
}
