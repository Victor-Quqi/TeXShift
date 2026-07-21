using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Animation;
using TeXShift.AddIn.Interop;
using TeXShift.AddIn.Localization;
using TeXShift.AddIn.UI.WPF.Converters;
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
                    NativeMethods.SetForegroundWindow(_ownerHandle);
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
            ShowDialogWithAnimation();
        }

        private void ShowDialogWithAnimation()
        {
            DialogScale.ScaleX = 0.85;
            DialogScale.ScaleY = 0.85;
            DialogOverlay.Visibility = Visibility.Visible;

            var duration = TimeSpan.FromMilliseconds(120);
            var ease = new QuadraticEase { EasingMode = EasingMode.EaseOut };

            var animX = new DoubleAnimation(0.85, 1.0, duration) { EasingFunction = ease };
            var animY = new DoubleAnimation(0.85, 1.0, duration) { EasingFunction = ease };

            DialogScale.BeginAnimation(System.Windows.Media.ScaleTransform.ScaleXProperty, animX);
            DialogScale.BeginAnimation(System.Windows.Media.ScaleTransform.ScaleYProperty, animY);
        }

        private void HideDialog()
        {
            var duration = TimeSpan.FromMilliseconds(80);
            var ease = new QuadraticEase { EasingMode = EasingMode.EaseIn };

            var animX = new DoubleAnimation(1.0, 0.85, duration) { EasingFunction = ease };
            var animY = new DoubleAnimation(1.0, 0.85, duration) { EasingFunction = ease };

            animX.Completed += (s, e) => DialogOverlay.Visibility = Visibility.Collapsed;

            DialogScale.BeginAnimation(System.Windows.Media.ScaleTransform.ScaleXProperty, animX);
            DialogScale.BeginAnimation(System.Windows.Media.ScaleTransform.ScaleYProperty, animY);
        }

        private void DialogConfirm_Click(object sender, RoutedEventArgs e)
        {
            HideDialog();
            ViewModel.LoadFromSettings(AppSettings.CreateDefault());
        }

        private void DialogCancel_Click(object sender, RoutedEventArgs e)
        {
            HideDialog();
        }

        private void DialogOverlay_MouseDown(object sender, MouseButtonEventArgs e)
        {
            HideDialog();
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
            var color = HexColorParser.ParseOrDefault(hex, System.Windows.Media.Colors.White);
            return System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B);
        }

        private static string DrawingColorToHex(System.Drawing.Color color)
        {
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        #endregion
    }
}
