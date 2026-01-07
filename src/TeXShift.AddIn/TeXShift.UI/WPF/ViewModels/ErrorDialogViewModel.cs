using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Input;
using TeXShift.Core.Localization;

namespace TeXShift.AddIn.UI.WPF.ViewModels
{
    public class ErrorDialogViewModel : ViewModelBase
    {
        public ErrorDialogViewModel(string userMessage, string technicalDetails, string debugFolderPath)
        {
            UserMessage = userMessage ?? Resources.GetString("Error_GenericUnexpected");
            TechnicalDetails = technicalDetails ?? string.Empty;
            DebugFolderPath = debugFolderPath ?? string.Empty;

            Title = Resources.GetString("Dialog_ErrorTitle");
            CopyButtonText = Resources.GetString("Dialog_CopyDetails");
            OpenFolderButtonText = Resources.GetString("Dialog_OpenDebugFolder");
            OkButtonText = Resources.GetString("Dialog_Ok");
            TechnicalDetailsHeader = Resources.GetString("Dialog_TechnicalDetails");

            CopyCommand = new RelayCommand(CopyDetails, CanCopyDetails);
            OpenFolderCommand = new RelayCommand(OpenDebugFolder, CanOpenDebugFolder);
        }

        public string Title { get; }
        public string UserMessage { get; }
        public string TechnicalDetails { get; }
        public string DebugFolderPath { get; }
        public string CopyButtonText { get; }
        public string OpenFolderButtonText { get; }
        public string OkButtonText { get; }
        public string TechnicalDetailsHeader { get; }

        public ICommand CopyCommand { get; }
        public ICommand OpenFolderCommand { get; }

        private bool CanCopyDetails(object param)
        {
            return !string.IsNullOrWhiteSpace(TechnicalDetails) || !string.IsNullOrWhiteSpace(UserMessage);
        }

        private void CopyDetails(object param)
        {
            var builder = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(UserMessage))
            {
                builder.AppendLine(UserMessage);
            }

            if (!string.IsNullOrWhiteSpace(TechnicalDetails))
            {
                if (builder.Length > 0)
                {
                    builder.AppendLine();
                }
                builder.AppendLine(TechnicalDetails);
            }

            if (!string.IsNullOrWhiteSpace(DebugFolderPath))
            {
                if (builder.Length > 0)
                {
                    builder.AppendLine();
                }
                var label = Resources.GetString("Dialog_DebugFolderLabel");
                builder.AppendLine(label + ": " + DebugFolderPath);
            }

            try
            {
                Clipboard.SetText(builder.ToString());
            }
            catch
            {
                // Ignore clipboard errors to avoid breaking the dialog flow.
            }
        }

        private bool CanOpenDebugFolder(object param)
        {
            return !string.IsNullOrWhiteSpace(DebugFolderPath);
        }

        private void OpenDebugFolder(object param)
        {
            if (string.IsNullOrWhiteSpace(DebugFolderPath))
                return;

            if (!Directory.Exists(DebugFolderPath))
            {
                Directory.CreateDirectory(DebugFolderPath);
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = DebugFolderPath,
                    UseShellExecute = true
                });
            }
            catch
            {
                // Ignore failures to keep the dialog responsive.
            }
        }
    }
}
