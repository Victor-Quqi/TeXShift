using System;
using System.Threading;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;
using TeXShift.AddIn.Interop;
using TeXShift.AddIn.UI.WPF.ViewModels;
using TeXShift.Core.Logging;

namespace TeXShift.AddIn.UI.WPF
{
    public partial class ErrorDialog : Window
    {
        private IntPtr _ownerHandle;

        internal static void WarmUp()
        {
            var thread = new Thread(() =>
            {
                try
                {
                    // Load WPF, the theme and control templates before the first error without showing a window.
                    var dialog = new ErrorDialog(new ErrorDialogViewModel(string.Empty, string.Empty, null));
                    try
                    {
                        var content = (FrameworkElement)dialog.Content;
                        var size = new Size(dialog.Width, dialog.Height);
                        content.Measure(size);
                        content.Arrange(new Rect(size));
                        content.UpdateLayout();
                    }
                    finally { dialog.Close(); }
                }
                catch (Exception ex)
                {
                    RuntimeLog.Write($"Error dialog warm-up failed: {ex}");
                }
                finally { Dispatcher.CurrentDispatcher.InvokeShutdown(); }
            }) { IsBackground = true, Name = "TeXShift_ErrorDialog_WarmUp_STA" };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        public ErrorDialog(ErrorDialogViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel ?? throw new ArgumentNullException(nameof(viewModel));

            Loaded += (s, e) =>
            {
                var helper = new WindowInteropHelper(this);
                if (helper.Handle == IntPtr.Zero)
                {
                    helper.EnsureHandle();
                }
                _ownerHandle = helper.Owner;
            };

            Closing += (s, e) =>
            {
                if (_ownerHandle != IntPtr.Zero)
                {
                    NativeMethods.SetForegroundWindow(_ownerHandle);
                }
            };
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            DataContext = null;
        }
    }
}
