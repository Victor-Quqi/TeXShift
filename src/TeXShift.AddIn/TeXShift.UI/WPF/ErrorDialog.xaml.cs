using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using TeXShift.AddIn.UI.WPF.ViewModels;

namespace TeXShift.AddIn.UI.WPF
{
    public partial class ErrorDialog : Window
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private IntPtr _ownerHandle;

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
                    SetForegroundWindow(_ownerHandle);
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
