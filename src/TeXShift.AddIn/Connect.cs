using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using TeXShift.AddIn.Localization;
using TeXShift.Core.Configuration;
using TeXShift.Core.Localization;
using TeXShift.Core.Services;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace TeXShift.AddIn
{
    /// <summary>
    /// Helper class to wrap a window handle for use with WinForms dialogs.
    /// </summary>
    internal class Win32Window : IWin32Window
    {
        public IntPtr Handle { get; }
        public Win32Window(IntPtr handle) => Handle = handle;
    }

    /// <summary>
    /// OneNote COM Add-in entry point.
    /// Partial class split into:
    /// - Connect.cs: Core add-in lifecycle and utilities
    /// - Connect.Ribbon.cs: Ribbon UI callbacks
    /// - Connect.Settings.cs: Settings dialog handling
    /// - Connect.Conversion.cs: Markdown conversion logic
    /// - Connect.Debug.cs: Debug tools
    /// </summary>
    [ComVisible(true)]
    [Guid(ComIdentity.Clsid)]
    [ProgId(ComIdentity.ProgId)]
    public partial class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        #region Native Methods

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool EnableWindow(IntPtr hWnd, bool bEnable);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        #endregion

        #region Fields

        private static readonly string _addInDirectory;
        private static readonly Mutex _processMutex =
            new Mutex(false, @"Local\TeXShift.AddIn." + ComIdentity.Clsid);

        private OneNote.Application _oneNoteApp;
        private IRibbonUI _ribbon;
        private ServiceContainer _serviceContainer;
        private AppSettings _appSettings;
        private SettingsManager _settingsManager;
        private Exception _initError;

        #endregion

        #region Static Constructor

        /// <summary>
        /// Static constructor to set up assembly resolution for COM Add-in.
        /// </summary>
        static Connect()
        {
            _addInDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;
        }

        /// <summary>
        /// Handles assembly resolution for dependencies that can't be found in the default probe paths.
        /// Also resolves satellite resource assemblies from culture subdirectories (e.g. zh-CN\).
        /// </summary>
        private static Assembly OnAssemblyResolve(object sender, ResolveEventArgs args)
        {
            var assemblyName = new AssemblyName(args.Name);

            // Satellite resource assemblies have a non-empty CultureName.
            // They reside in culture subdirectories, e.g. zh-CN\TeXShift.Core.resources.dll
            if (!string.IsNullOrEmpty(assemblyName.CultureName))
            {
                var satellitePath = Path.Combine(
                    _addInDirectory,
                    assemblyName.CultureName,
                    assemblyName.Name + ".dll");

                if (File.Exists(satellitePath))
                    return Assembly.LoadFrom(satellitePath);

                return null;
            }

            var assemblyPath = Path.Combine(_addInDirectory, assemblyName.Name + ".dll");

            if (File.Exists(assemblyPath))
                return Assembly.LoadFrom(assemblyPath);

            return null;
        }

        #endregion

        #region IDTExtensibility2 Implementation

        /// <summary>
        /// Called when the add-in is connected to OneNote.
        /// </summary>
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _oneNoteApp = (OneNote.Application)Application;

                // Load settings from JSON file
                _settingsManager = new SettingsManager();
                _appSettings = _settingsManager.Load();

                // Initialize localization based on settings or system culture
                LocalizationManager.Initialize(_appSettings?.Language);

                // Initialize dependency injection container
                _serviceContainer = new ServiceContainer();

                // Apply loaded settings to style configuration
                ApplySettingsToStyleConfig();
            }
            catch (Exception ex)
            {
                _initError = ex;
                System.Diagnostics.Debug.WriteLine($"[TeXShift] OnConnection failed: {ex}");
            }
        }

        /// <summary>
        /// Called when the add-in is disconnected from OneNote.
        /// </summary>
        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            // Dispose ServiceContainer to clean up WebView2 and other resources
            _serviceContainer?.Dispose();
            _serviceContainer = null;

            // Explicitly release the COM object to ensure OneNote can shut down cleanly.
            SafeReleaseComObject(_oneNoteApp);

            // While explicit release is key, garbage collection can help clean up any other managed wrappers.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Called when the add-in is loaded on startup.
        /// </summary>
        public void OnStartupComplete(ref Array custom)
        {
        }

        /// <summary>
        /// Called when OneNote is shutting down.
        /// </summary>
        public void OnBeginShutdown(ref Array custom)
        {
        }

        /// <summary>
        /// Called when add-ins are updated.
        /// </summary>
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        #endregion

        #region IRibbonExtensibility Implementation

        /// <summary>
        /// Returns the XML for the custom Ribbon UI.
        /// </summary>
        public string GetCustomUI(string RibbonID)
        {
            return GetResourceText("TeXShift.AddIn.Ribbon.xml");
        }

        public void OnRibbonLoad(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        #endregion

        #region Utility Methods

        /// <summary>
        /// Safely releases a COM object and sets its reference to null.
        /// </summary>
        /// <param name="obj">The COM object to release.</param>
        private void SafeReleaseComObject(object obj)
        {
            if (obj != null)
            {
                try
                {
                    Marshal.ReleaseComObject(obj);
                }
                catch (Exception ex)
                {
                    // Log the exception but don't throw - object might already be released
                    System.Diagnostics.Debug.WriteLine($"Warning: Failed to release COM object: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Helper function: Shows a MessageBox that appears on top of all windows.
        /// </summary>
        private DialogResult ShowTopMostMessageBox(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            var owner = new Win32Window(GetForegroundWindow());
            return MessageBox.Show(owner, text, caption, buttons, icon);
        }

        /// <summary>
        /// Shows a localized add-in initialization failure message if initialization has failed.
        /// </summary>
        private bool TryShowInitializationError()
        {
            if (_initError == null)
            {
                return false;
            }

            ShowTopMostMessageBox(
                string.Format(UIResources.GetString("Message_AddInInitializationFailed"), _initError),
                Resources.GetString("Dialog_ErrorTitle"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return true;
        }

        /// <summary>
        /// Helper function: Creates a form with a scrollbar to display a large amount of text.
        /// </summary>
        private void ShowTextInScrollableMessageBox(string text, string caption)
        {
            var owner = new Win32Window(GetForegroundWindow());
            using (Form form = new Form
            {
                Text = caption,
                Size = new System.Drawing.Size(600, 400),
                StartPosition = FormStartPosition.CenterParent
            })
            {
                TextBox textBox = new TextBox
                {
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    Dock = DockStyle.Fill,
                    ReadOnly = true,
                    Font = new System.Drawing.Font("Consolas", 10),
                    Text = text
                };
                form.Controls.Add(textBox);
                form.ShowDialog(owner);
            }
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            foreach (string name in resourceNames)
            {
                if (string.Compare(resourceName, name, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(name)))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
