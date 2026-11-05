using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Win32;
using OneNoteInterop = Microsoft.Office.Interop.OneNote;
using TeXShift.Core.OneNote;
using TeXShift.Core.Utils;
using TeXShift.Tests.E2E.Models;

namespace TeXShift.Tests.E2E
{
    internal sealed class TestPageManager : IDisposable
    {
        private const string TestNotebookName = "TeXShift_E2E_Tests";
        private const string TestSectionName = "E2E";
        private const string OneNoteAppPathRegistryKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\onenote.exe";
        private const string DefaultOneNoteExecutablePath = @"C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE";
        private const int SwShowNoActivate = 4;
        private const int OneNoteStartupWindowTimeoutMs = 10000;
        private const int OneNoteStartupWindowPollIntervalMs = 200;
        private const int OneNoteComReadyTimeoutMs = 10000;
        private const int OneNoteComReadyPollIntervalMs = 250;
        private const int OneNoteCloseTimeoutMs = 8000;
        private const uint WmSysCommand = 0x0112;
        private const int ScClose = 0xF060;
        private const int VkEnd = 0x23;
        private const int VkDown = 0x28;
        private const int MaxSelectionKeySteps = 200;
        private const int SelectionKeyTimeoutMs = 1000;
        private const int SelectionPollIntervalMs = 50;
        private const int HrNotYetSynchronized = unchecked((int)0x8004201D);
        private static readonly XNamespace OneNoteNamespace = OneNoteXml.Namespace;

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder className, int maxCount);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        private static bool IsBoolAttributeTrue(XElement element, string attributeName)
        {
            string value = (string)element?.Attribute(attributeName);
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            value = value.Trim();
            if (bool.TryParse(value, out bool result))
            {
                return result;
            }

            // xsd:boolean may also be serialized as "1"/"0".
            return string.Equals(value, "1", StringComparison.Ordinal);
        }

        private readonly OneNoteComDispatcher _oneNoteDispatcher;
        private bool _disposed;
        private string _createdNotebookId;
        private string _createdSectionId;
        private string _testPageId;
        private string _importedSectionPath;
        private bool _createdNotebook;  // Track if we created the notebook
        private string _notebookPath;   // Path to delete on cleanup
        private string _originalPageId; // Page ID to navigate back to after test
        private string _importedPageId; // Page reproduced from a section file
        private string _importedSectionId; // Section file copy holding that page
        private bool _launchedOneNoteProcess; // Track if we launched OneNote
        private Process _ownedOneNoteProcess;
        private bool _shouldCloseOneNote; // Flag to close OneNote on dispose

        private TestPageManager()
        {
            try
            {
                // OneNote does not expose its restore target without a real window. Navigating to
                // and deleting E2E pages in that state can leave its saved page pointer dangling.
                LaunchOneNoteWithWindowIfNeeded();

                // Existing processes may still be opening their window. Wait before COM activation
                // so it attaches to a notebook window rather than a headless -Embedding instance.
                WaitForOneNoteMainWindow();

                var window = GetOneNoteMainWindowHandleOrThrow();
                if (IsIconic(window)) ShowWindow(window, SwShowNoActivate);
                _oneNoteDispatcher = OneNoteComDispatcher.Connect();
            }
            catch
            {
                _ownedOneNoteProcess?.Dispose();
                throw;
            }
        }

        public static Task<TestPageManager> CreateAsync()
        {
            return Task.Run(() => new TestPageManager());
        }

        public IntPtr MainWindowHandle => GetOneNoteMainWindowHandleOrThrow();

        public Task<string> GetPageSelectionXmlAsync(string pageId) => _oneNoteDispatcher.InvokeAsync(app =>
        {
            app.GetPageContent(pageId, out string xml, OneNoteInterop.PageInfo.piSelection, OneNoteInterop.XMLSchema.xs2013);
            return xml;
        });

        public Task<bool> ExportPageAsync(string pageId, string outputPath, OneNoteInterop.PublishFormat format) => _oneNoteDispatcher.InvokeAsync(app =>
        {
            app.Publish(pageId, outputPath, format, string.Empty);
            return File.Exists(outputPath);
        });

        public OneNoteComDispatcher OneNoteDispatcher => _oneNoteDispatcher;

        public bool LaunchedOneNoteProcess => _launchedOneNoteProcess;
        public string OriginalPageId => _originalPageId;
        public string TestPageId => _testPageId;
        public int? OneNoteProcessId
        {
            get
            {
                var window = GetOneNoteMainWindowHandle();
                if (window == IntPtr.Zero) return _ownedOneNoteProcess?.Id;
                GetWindowThreadProcessId(window, out uint pid);
                return (int)pid;
            }
        }

        public Task<string> GetHierarchyPagesXmlAsync()
        {
            return _oneNoteDispatcher.InvokeAsync(app =>
            {
                EnsureNotDisposed();
                app.GetHierarchy(null, OneNoteInterop.HierarchyScope.hsPages, out string hierarchyXml, OneNoteInterop.XMLSchema.xs2013);
                return hierarchyXml;
            });
        }

        public Task<string> GetPageContentBasicXmlAsync(string pageId)
        {
            if (string.IsNullOrWhiteSpace(pageId))
            {
                throw new ArgumentException("Page ID is required.", nameof(pageId));
            }

            return _oneNoteDispatcher.InvokeAsync(app =>
            {
                EnsureNotDisposed();
                app.GetPageContent(pageId, out string pageXml, OneNoteInterop.PageInfo.piBasic, OneNoteInterop.XMLSchema.xs2013);
                return pageXml;
            });
        }

        public Task NavigateToAsync(string pageId)
        {
            if (string.IsNullOrWhiteSpace(pageId))
            {
                throw new ArgumentException("Page ID is required.", nameof(pageId));
            }

            return _oneNoteDispatcher.InvokeAsync(app =>
            {
                EnsureNotDisposed();
                app.NavigateTo(pageId, null, false);
            });
        }

        public Task NavigateToWhenReadyAsync(string pageId)
        {
            if (string.IsNullOrWhiteSpace(pageId))
            {
                throw new ArgumentException("Page ID is required.", nameof(pageId));
            }

            return _oneNoteDispatcher.InvokeAsync(app =>
            {
                EnsureNotDisposed();
                var deadline = DateTime.UtcNow.AddMilliseconds(OneNoteComReadyTimeoutMs);

                while (true)
                {
                    try
                    {
                        app.NavigateTo(pageId, null, false);
                        return;
                    }
                    catch (COMException)
                    {
                        if (DateTime.UtcNow >= deadline)
                        {
                            throw;
                        }
                    }

                    Thread.Sleep(OneNoteComReadyPollIntervalMs);
                }
            });
        }

        public Task WaitForOneNoteReadyAsync()
        {
            return _oneNoteDispatcher.InvokeAsync(app =>
            {
                EnsureNotDisposed();
                WaitForCurrentPage(app);
                if (string.IsNullOrWhiteSpace(_originalPageId))
                {
                    throw new TimeoutException("OneNote did not expose a current page before the COM readiness timeout.");
                }
            });
        }

        public void CloseLaunchedOneNoteOnDispose()
        {
            EnsureNotDisposed();
            _shouldCloseOneNote = _launchedOneNoteProcess;
        }

        public static void CloseAllOneNoteProcesses()
        {
            var processes = Process.GetProcessesByName("ONENOTE");
            try
            {
                foreach (var process in processes) CloseOneNoteProcess(process);
            }
            finally { foreach (var process in processes) process.Dispose(); }
        }

        private static void CloseOneNoteProcess(Process process)
        {
            var deadline = DateTime.UtcNow.AddMilliseconds(OneNoteCloseTimeoutMs);
            while (!process.HasExited && DateTime.UtcNow < deadline)
            {
                process.Refresh();
                var window = process.MainWindowHandle;
                if (window != IntPtr.Zero)
                    PostMessage(window, WmSysCommand, new IntPtr(ScClose), IntPtr.Zero);
                if (process.WaitForExit(250)) return;
            }
            if (!process.HasExited)
                throw new TimeoutException($"OneNote process {process.Id} did not exit cleanly.");
        }

        public Task<string> CreateTestPageAsync(string testName, string markdownContent, PageSelection selection)
        {
            if (string.IsNullOrWhiteSpace(testName))
            {
                throw new ArgumentException("Test name is required.", nameof(testName));
            }

            return _oneNoteDispatcher.InvokeAsync(app => CreateTestPage(app, testName, markdownContent ?? string.Empty, selection ?? PageSelection.Default));
        }

        /// <summary>
        /// Reproduces a captured page in the test notebook from page XML (dump --page) or a section
        /// file exported from OneNote (.one, the faithful copy) and sets up <paramref name="selection"/>.
        /// </summary>
        public Task<string> CreateTestPageFromFileAsync(string pageFilePath, PageSelection selection)
        {
            if (string.IsNullOrWhiteSpace(pageFilePath))
            {
                throw new ArgumentException("Page file is required.", nameof(pageFilePath));
            }

            return _oneNoteDispatcher.InvokeAsync(app => CreateTestPageFromFile(app, pageFilePath, selection ?? PageSelection.Default));
        }

        /// <summary>
        /// Deletes only resources created by this test session, including incomplete pages.
        /// </summary>
        public Task CleanupTestResourcesAsync()
        {
            return _oneNoteDispatcher.InvokeAsync(CleanupTestResources);
        }

        /// <summary>
        /// Saves the current active page ID so we can navigate back after test.
        /// Call this before creating test pages.
        /// </summary>
        public Task SaveCurrentPageAsync()
        {
            return _oneNoteDispatcher.InvokeAsync(SaveCurrentPage);
        }

        /// <summary>
        /// Navigates back to the original page that was active before the test.
        /// Call this before cleanup to keep the restore target valid.
        /// </summary>
        public Task RestoreOriginalPageAsync()
        {
            return _oneNoteDispatcher.InvokeAsync(RestoreOriginalPage);
        }

        private void SaveCurrentPage(OneNoteInterop.Application app)
        {
            EnsureNotDisposed();

            if (_launchedOneNoteProcess)
            {
                WaitForCurrentPage(app);
                return;
            }

            SaveCurrentPageOnce(app);
        }

        private void SaveCurrentPageOnce(OneNoteInterop.Application app)
        {
            OneNoteInterop.Window window = null;
            OneNoteInterop.Windows windows = null;

            try
            {
                windows = app.Windows;
                if (windows == null || windows.Count == 0)
                {
                    return;
                }

                window = windows.CurrentWindow;
                if (window != null)
                {
                    _originalPageId = window.CurrentPageId;
                }
            }
            finally
            {
                if (window != null) Marshal.ReleaseComObject(window);
                if (windows != null) Marshal.ReleaseComObject(windows);
            }
        }

        private void WaitForCurrentPage(OneNoteInterop.Application app)
        {
            var deadline = DateTime.UtcNow.AddMilliseconds(OneNoteComReadyTimeoutMs);

            while (DateTime.UtcNow < deadline)
            {
                OneNoteInterop.Window window = null;
                OneNoteInterop.Windows windows = null;

                try
                {
                    windows = app.Windows;
                    if (windows != null && windows.Count > 0)
                    {
                        window = windows.CurrentWindow;
                        string currentPageId = window?.CurrentPageId;
                        if (!string.IsNullOrWhiteSpace(currentPageId))
                        {
                            _originalPageId = currentPageId;
                            return;
                        }
                    }
                }
                catch
                {
                    // OneNote may reject COM calls while its UI is still initializing
                }
                finally
                {
                    if (window != null)
                    {
                        try
                        {
                            Marshal.ReleaseComObject(window);
                        }
                        catch
                        {
                            // Ignore release failures during startup polling
                        }
                    }

                    if (windows != null)
                    {
                        try
                        {
                            Marshal.ReleaseComObject(windows);
                        }
                        catch
                        {
                            // Ignore release failures during startup polling
                        }
                    }
                }

                Thread.Sleep(OneNoteComReadyPollIntervalMs);
            }
        }

        private void RestoreOriginalPage(OneNoteInterop.Application app)
        {
            EnsureNotDisposed();

            _shouldCloseOneNote = _launchedOneNoteProcess;

            if (!string.IsNullOrWhiteSpace(_originalPageId))
            {
                app.NavigateTo(_originalPageId, null, false);
                return;
            }
            if (_testPageId != null)
                throw new InvalidOperationException("The original OneNote page was unavailable; test resources were retained.");
        }

        private void LaunchOneNoteWithWindowIfNeeded()
        {
            if (IsOneNoteProcessRunning())
            {
                return;
            }

            try
            {
                string executablePath = GetOneNoteExecutablePath();
                string commandLine = $"\"{executablePath.Trim().Trim('\"')}\"";

                // WMI launches through WmiPrvSE so OneNote cannot inherit foreground activation rights.
                using (var processClass = new ManagementClass("Win32_Process"))
                using (var inputParameters = processClass.GetMethodParameters("Create"))
                {
                    inputParameters["CommandLine"] = commandLine;

                    using (var outputParameters = processClass.InvokeMethod("Create", inputParameters, null))
                    {
                        if (outputParameters != null)
                        {
                            object returnValue = outputParameters["ReturnValue"];
                            _launchedOneNoteProcess = returnValue != null && Convert.ToUInt32(returnValue) == 0;
                            if (_launchedOneNoteProcess)
                            {
                                _ownedOneNoteProcess = Process.GetProcessById(Convert.ToInt32(outputParameters["ProcessId"]));
                                // Retain a process handle so PID reuse cannot target another instance.
                                _ = _ownedOneNoteProcess.Handle;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Could not launch a OneNote window for the test.", ex);
            }
            if (!_launchedOneNoteProcess)
                throw new InvalidOperationException("WMI could not launch OneNote for the test.");
        }

        private static bool IsOneNoteProcessRunning()
        {
            Process[] processes = null;

            try
            {
                processes = Process.GetProcessesByName("ONENOTE");
                return processes.Length > 0;
            }
            catch
            {
                // Avoid starting a duplicate process when detection is unavailable
                return true;
            }
            finally
            {
                if (processes != null)
                {
                    foreach (var process in processes)
                    {
                        process.Dispose();
                    }
                }
            }
        }

        private static string GetOneNoteExecutablePath()
        {
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(OneNoteAppPathRegistryKey))
                {
                    string registeredPath = key?.GetValue(null) as string;
                    if (!string.IsNullOrWhiteSpace(registeredPath))
                    {
                        return registeredPath;
                    }
                }
            }
            catch
            {
                // Use the standard Office x64 installation path below
            }

            return DefaultOneNoteExecutablePath;
        }

        private static void WaitForOneNoteMainWindow()
        {
            var deadline = DateTime.UtcNow.AddMilliseconds(OneNoteStartupWindowTimeoutMs);

            while (DateTime.UtcNow < deadline)
            {
                IntPtr mainWindowHandle = GetOneNoteMainWindowHandle();
                var className = new StringBuilder(256);
                if (mainWindowHandle != IntPtr.Zero)
                    GetClassName(mainWindowHandle, className, className.Capacity);
                if (className.ToString() == "Framework::CFrame")
                {
                    return;
                }

                Thread.Sleep(OneNoteStartupWindowPollIntervalMs);
            }
            throw new TimeoutException("OneNote did not open a notebook window. Check for a startup or recovery dialog.");
        }

        private static IntPtr GetOneNoteMainWindowHandle()
        {
            Process[] processes = null;

            try
            {
                processes = Process.GetProcessesByName("ONENOTE");
                foreach (var process in processes)
                {
                    try
                    {
                        process.Refresh();
                        IntPtr mainWindowHandle = process.MainWindowHandle;
                        if (mainWindowHandle != IntPtr.Zero)
                        {
                            return mainWindowHandle;
                        }
                    }
                    catch
                    {
                        // The process may exit or still be initializing
                    }
                }
            }
            catch
            {
                // Ignore process lookup failures while OneNote is starting
            }
            finally
            {
                if (processes != null)
                {
                    foreach (var process in processes)
                    {
                        process.Dispose();
                    }
                }
            }

            return IntPtr.Zero;
        }

        private void CloseOneNoteProcess()
        {
            if (_ownedOneNoteProcess != null) CloseOneNoteProcess(_ownedOneNoteProcess);
        }

        private void CleanupTestResources(OneNoteInterop.Application app)
        {
            EnsureNotDisposed();
            var failures = new List<Exception>();
            void Attempt(Action action)
            {
                try { action(); }
                catch (Exception ex) { failures.Add(ex); }
            }

            if (_testPageId != null) Attempt(() => DeletePage(app, _testPageId));
            if (_importedSectionId != null)
                Attempt(() =>
                {
                    app.DeleteHierarchy(_importedSectionId, DateTime.MinValue, true);
                    _importedSectionId = null;
                });
            if (_importedSectionPath != null && File.Exists(_importedSectionPath))
                Attempt(() => File.Delete(_importedSectionPath));
            if (_createdSectionId != null)
                Attempt(() => app.DeleteHierarchy(_createdSectionId, DateTime.MinValue, true));
            if (_createdNotebook)
                Attempt(() =>
                {
                    app.CloseNotebook(_createdNotebookId, false);
                    if (Directory.Exists(_notebookPath)) Directory.Delete(_notebookPath, recursive: true);
                });
            if (failures.Count > 0) throw new AggregateException("Test resource cleanup failed.", failures);
        }

        private string CreateTestPage(OneNoteInterop.Application app, string testName, string markdownContent, PageSelection selection)
        {
            EnsureNotDisposed();

            string notebookId = GetOrCreateNotebookId(app);
            string sectionId = GetOrCreateSectionId(app, notebookId);

            app.CreateNewPage(sectionId, out string pageId, OneNoteInterop.NewPageStyle.npsDefault);
            if (string.IsNullOrWhiteSpace(pageId))
            {
                throw new InvalidOperationException("Failed to create a new OneNote page.");
            }

            _testPageId = pageId;
            UpdatePageContent(app, pageId, testName, markdownContent, addCaretParagraph: selection.IsDefault);
            PlaceSelection(app, pageId, selection);
            return pageId;
        }

        private void UpdatePageContent(OneNoteInterop.Application app, string pageId, string title, string content, bool addCaretParagraph)
        {
            app.GetPageContent(pageId, out string pageXml, OneNoteInterop.PageInfo.piAll, OneNoteInterop.XMLSchema.xs2013);
            var doc = XDocument.Parse(pageXml);
            var ns = doc.Root?.Name.Namespace ?? OneNoteNamespace;

            // Update title
            var titleElement = doc.Root?.Element(ns + "Title");
            if (titleElement != null)
            {
                var titleOE = titleElement.Element(ns + "OE");
                if (titleOE != null)
                {
                    var titleText = titleOE.Element(ns + "T");
                    if (titleText != null)
                    {
                        titleText.Value = title;
                    }
                }
            }

            // Update or create content outline
            var outline = doc.Root.Element(ns + "Outline");
            if (outline == null)
            {
                outline = new XElement(ns + "Outline");
                doc.Root.Add(outline);
            }

            // Clear existing content
            var existingChildren = outline.Element(ns + "OEChildren");
            existingChildren?.Remove();

            // Create OEChildren with one OE per line (mimics normal OneNote structure).
            // An empty first OE takes the default caret, so the whole Markdown is converted.
            var oeChildren = new XElement(ns + "OEChildren");
            if (addCaretParagraph)
            {
                oeChildren.Add(new XElement(ns + "OE", new XElement(ns + "T", new XCData(string.Empty))));
            }

            var lines = (content ?? string.Empty).Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            foreach (var line in lines)
            {
                var oe = new XElement(ns + "OE");
                var text = new XElement(ns + "T", new XCData(OneNoteHtmlTextEscaper.Escape(line)));
                oe.Add(text);
                oeChildren.Add(oe);
            }
            outline.Add(oeChildren);

            app.UpdatePageContent(doc.ToString(), DateTime.MinValue, OneNoteInterop.XMLSchema.xs2013, true);
        }

        private string CreateTestPageFromFile(OneNoteInterop.Application app, string pageFilePath, PageSelection selection)
        {
            EnsureNotDisposed();

            string notebookId = GetOrCreateNotebookId(app);
            string pageId = string.Equals(Path.GetExtension(pageFilePath), ".one", StringComparison.OrdinalIgnoreCase)
                ? ImportSectionPage(app, notebookId, pageFilePath)
                : CreatePageFromXml(app, GetOrCreateSectionId(app, notebookId), File.ReadAllText(pageFilePath));

            PlaceSelection(app, pageId, selection);
            return pageId;
        }

        private string CreatePageFromXml(OneNoteInterop.Application app, string sectionId, string pageXml)
        {
            app.CreateNewPage(sectionId, out string pageId, OneNoteInterop.NewPageStyle.npsDefault);
            if (string.IsNullOrWhiteSpace(pageId))
            {
                throw new InvalidOperationException("Failed to create a new OneNote page.");
            }

            _testPageId = pageId;
            var doc = XDocument.Parse(pageXml);

            // Object identities, view state and selection belong to the source page.
            doc.Root.SetAttributeValue("ID", pageId);
            foreach (var element in doc.Root.DescendantsAndSelf())
            {
                element.Attribute("objectID")?.Remove();
                element.Attribute("selected")?.Remove();
                element.Attribute("isCurrentlyViewed")?.Remove();
            }

            app.UpdatePageContent(doc.ToString(), DateTime.MinValue, OneNoteInterop.XMLSchema.xs2013, true);
            return pageId;
        }

        /// <summary>
        /// Opens a copy of the section file in the test notebook and returns its first page.
        /// </summary>
        private string ImportSectionPage(OneNoteInterop.Application app, string notebookId, string sectionFilePath)
        {
            app.GetHierarchy(notebookId, OneNoteInterop.HierarchyScope.hsSelf, out string notebookXml, OneNoteInterop.XMLSchema.xs2013);
            string notebookPath = (string)XDocument.Parse(notebookXml).Root?.Attribute("path");
            if (string.IsNullOrWhiteSpace(notebookPath))
            {
                throw new InvalidOperationException("Cannot get notebook path to import the section.");
            }

            string sectionPath = Path.Combine(
                notebookPath,
                $"{Path.GetFileNameWithoutExtension(sectionFilePath)}_{DateTime.Now:yyyyMMdd_HHmmss}.one");
            File.Copy(sectionFilePath, sectionPath);
            _importedSectionPath = sectionPath;

            app.OpenHierarchy(sectionPath, null, out string sectionId, OneNoteInterop.CreateFileType.cftNone);
            if (string.IsNullOrWhiteSpace(sectionId))
            {
                throw new InvalidOperationException($"Failed to open the section file '{sectionPath}'.");
            }

            _importedSectionId = sectionId;
            _importedPageId = WaitForSectionPageIds(app, sectionId).FirstOrDefault()
                ?? throw new InvalidOperationException("The section file has no page.");
            _testPageId = _importedPageId;
            return _importedPageId;
        }

        /// <summary>
        /// Lists the pages of a just-opened section, waiting while OneNote is still loading it.
        /// </summary>
        private List<string> WaitForSectionPageIds(OneNoteInterop.Application app, string sectionId)
        {
            var deadline = DateTime.UtcNow.AddMilliseconds(OneNoteComReadyTimeoutMs);

            while (true)
            {
                try
                {
                    var pageIds = GetPageIds(app, sectionId);
                    if (pageIds.Count > 0 || DateTime.UtcNow >= deadline)
                    {
                        return pageIds;
                    }
                }
                catch (COMException ex) when (ex.HResult == HrNotYetSynchronized && DateTime.UtcNow < deadline)
                {
                }

                Thread.Sleep(OneNoteComReadyPollIntervalMs);
            }
        }

        /// <summary>
        /// Sets up <paramref name="selection"/> with OneNote's own navigation and keyboard handling, so
        /// the reader sees the selection markup a user's caret or drag would produce.
        /// </summary>
        private void PlaceSelection(OneNoteInterop.Application app, string pageId, PageSelection selection)
        {
            if (selection.IsEmptyOutline)
            {
                app.NavigateTo(pageId, null, false);
                var changes = new XElement(OneNoteNamespace + "Page", new XAttribute("ID", pageId),
                    new XElement(OneNoteNamespace + "Outline",
                        new XElement(OneNoteNamespace + "Position", new XAttribute("x", 700), new XAttribute("y", 100), new XAttribute("z", 0)),
                        new XElement(OneNoteNamespace + "OEChildren", new XElement(OneNoteNamespace + "OE",
                            new XElement(OneNoteNamespace + "T", new XAttribute("selected", "all"), new XCData(string.Empty))))));
                app.UpdatePageContent(changes.ToString(), DateTime.MinValue, OneNoteInterop.XMLSchema.xs2013, true);
                // A live empty editor survives; an empty outline imported without selection does not.
                WaitForSelection(app, pageId, page => page.Descendants(OneNoteNamespace + "Outline").Any(outline =>
                    outline.Descendants(OneNoteNamespace + "T").Any(t => (string)t.Attribute("selected") == "all") &&
                    outline.Descendants(OneNoteNamespace + "T").All(t => t.Value.Length == 0)));
                return;
            }

            app.GetPageContent(pageId, out string pageXml, OneNoteInterop.PageInfo.piBasic, OneNoteInterop.XMLSchema.xs2013);
            var paragraphs = XDocument.Parse(pageXml).Root.Elements(OneNoteNamespace + "Outline")
                .SelectMany(outline => outline.Descendants(OneNoteNamespace + "OE"))
                .ToList();

            var first = selection.IsDefault
                ? paragraphs.FirstOrDefault() ?? throw new InvalidOperationException("The page has no Outline to convert.")
                : FindParagraph(paragraphs, selection.FromText);
            string firstId = (string)first.Attribute("objectID");

            if (selection.WholeOutline)
            {
                string outlineId = (string)first.Ancestors(OneNoteNamespace + "Outline").First().Attribute("objectID");
                app.NavigateTo(pageId, outlineId, false);
                WaitForSelection(app, pageId, page => page.Descendants(OneNoteNamespace + "Outline")
                    .Any(outline => (string)outline.Attribute("objectID") == outlineId && outline.Attribute("selected") != null));
                return;
            }

            // NavigateTo selects the whole paragraph.
            app.NavigateTo(pageId, firstId, false);
            WaitForSelection(app, pageId, page => FindById(page, firstId)?.DescendantsAndSelf()
                .Any(node => node.Attribute("selected") != null) == true);

            if (!selection.IsRange)
            {
                // End collapses it to a caret, which OneNote reports as an empty selected T.
                OneNoteEditorInput.PressKey(MainWindowHandle, VkEnd);
                WaitForSelection(app, pageId, page => FindById(page, firstId)?.Elements(OneNoteNamespace + "T")
                    .Any(text => (string)text.Attribute("selected") == "all" && text.Value.Length == 0) == true);
                return;
            }

            var last = FindParagraph(paragraphs, selection.ToText);
            if (paragraphs.IndexOf(last) < paragraphs.IndexOf(first)
                || last.Ancestors(OneNoteNamespace + "Outline").First() != first.Ancestors(OneNoteNamespace + "Outline").First())
            {
                throw new InvalidOperationException(
                    $"The paragraph containing \"{selection.ToText}\" must follow \"{selection.FromText}\" in the same Outline.");
            }

            ExtendSelectionTo(app, pageId, (string)last.Attribute("objectID"));
        }

        /// <summary>
        /// Extends the selection line by line (Shift+Down) until it reaches the last paragraph, then to the
        /// end of each of its lines (Shift+End) until the paragraph is fully selected.
        /// </summary>
        private void ExtendSelectionTo(OneNoteInterop.Application app, string pageId, string lastOeId)
        {
            OneNoteEditorInput.HoldShift(MainWindowHandle, () =>
            {
                string signature = ReadSelection(app, pageId, lastOeId, out string lastState);
                int lastKey = 0;
                for (int step = 0; lastState != "all"; step++)
                {
                    if (step >= MaxSelectionKeySteps)
                    {
                        throw new InvalidOperationException("Could not extend the selection to the last paragraph.");
                    }

                    lastKey = lastState == null || lastKey == VkEnd ? VkDown : VkEnd;
                    OneNoteEditorInput.PressKey(MainWindowHandle, lastKey);

                    // End at a line end changes nothing, so an unchanged selection after the wait is expected.
                    var deadline = DateTime.UtcNow.AddMilliseconds(SelectionKeyTimeoutMs);
                    string previous = signature;
                    do
                    {
                        Thread.Sleep(SelectionPollIntervalMs);
                        signature = ReadSelection(app, pageId, lastOeId, out lastState);
                    }
                    while (signature == previous && DateTime.UtcNow < deadline);
                }
            });
        }

        private static XElement FindParagraph(IEnumerable<XElement> paragraphs, string containedText)
        {
            return paragraphs.FirstOrDefault(oe => oe.Elements(OneNoteNamespace + "T")
                    .Any(text => text.Value.IndexOf(containedText, StringComparison.Ordinal) >= 0))
                ?? throw new InvalidOperationException($"No paragraph on the page contains \"{containedText}\".");
        }

        private static XElement FindById(XDocument page, string objectId)
        {
            return page.Descendants(OneNoteNamespace + "OE")
                .FirstOrDefault(node => (string)node.Attribute("objectID") == objectId);
        }

        /// <summary>
        /// Returns a fingerprint of the page selection and how much of the given OE it covers.
        /// </summary>
        private string ReadSelection(OneNoteInterop.Application app, string pageId, string oeId, out string oeState)
        {
            app.GetPageContent(pageId, out string xml, OneNoteInterop.PageInfo.piSelection, OneNoteInterop.XMLSchema.xs2013);
            var page = XDocument.Parse(xml);
            oeState = (string)FindById(page, oeId)?.Attribute("selected");

            return string.Join("|", page.Descendants()
                .Where(node => node.Attribute("selected") != null)
                .Select(node => $"{(string)node.Attribute("objectID") ?? node.Name.LocalName}:{(string)node.Attribute("selected")}:{(node.Name == OneNoteNamespace + "T" ? node.Value.Length : 0)}"));
        }

        /// <summary>
        /// Polls GetPageContent until the page selection satisfies <paramref name="isReady"/>, or timeout.
        /// </summary>
        private void WaitForSelection(
            OneNoteInterop.Application app,
            string pageId,
            Func<XDocument, bool> isReady,
            int timeoutMs = 10000,
            int pollIntervalMs = 100)
        {
            var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);

            while (DateTime.UtcNow < deadline)
            {
                app.GetPageContent(pageId, out string xml, OneNoteInterop.PageInfo.piSelection, OneNoteInterop.XMLSchema.xs2013);
                if (isReady(XDocument.Parse(xml)))
                {
                    return;
                }

                Thread.Sleep(pollIntervalMs);
            }

            // Timeout reached - fail fast to avoid flaky follow-up steps
            throw new TimeoutException("WaitForSelection: Timeout waiting for the page selection.");
        }

        private static IntPtr GetOneNoteMainWindowHandleOrThrow()
        {
            IntPtr mainWindow = GetOneNoteMainWindowHandle();
            if (mainWindow == IntPtr.Zero)
            {
                throw new InvalidOperationException("OneNote main window not found.");
            }

            return mainWindow;
        }

        private void DeletePage(OneNoteInterop.Application app, string pageId)
        {
            EnsureNotDisposed();

            // An imported section copy exists only for its page, so it is removed with it.
            bool imported = pageId == _importedPageId;
            app.DeleteHierarchy(imported ? _importedSectionId : pageId, DateTime.MinValue, true);
            if (_testPageId == pageId) _testPageId = null;
            if (imported)
            {
                _importedPageId = null;
                _importedSectionId = null;
            }
        }

        private string GetOrCreateNotebookId(OneNoteInterop.Application app)
        {
            app.GetHierarchy(null, OneNoteInterop.HierarchyScope.hsNotebooks, out string hierarchyXml, OneNoteInterop.XMLSchema.xs2013);
            var doc = XDocument.Parse(hierarchyXml);

            var notebook = doc.Descendants(OneNoteNamespace + "Notebook")
                .FirstOrDefault(node =>
                    string.Equals((string)node.Attribute("name"), TestNotebookName, StringComparison.OrdinalIgnoreCase)
                    && !IsBoolAttributeTrue(node, "isInRecycleBin"));
            if (notebook != null)
            {
                // Notebook already exists - NOT created by us, don't delete on cleanup
                _createdNotebook = false;
                return (string)notebook.Attribute("ID");
            }

            app.GetSpecialLocation(OneNoteInterop.SpecialLocation.slDefaultNotebookFolder, out string defaultFolder);
            if (string.IsNullOrWhiteSpace(defaultFolder))
            {
                throw new InvalidOperationException("Cannot resolve OneNote default notebook folder.");
            }

            string notebookPath = Path.Combine(defaultFolder, TestNotebookName);
            bool directoryExisted = Directory.Exists(notebookPath);
            app.OpenHierarchy(notebookPath, null, out string notebookId, OneNoteInterop.CreateFileType.cftNotebook);
            if (string.IsNullOrWhiteSpace(notebookId))
            {
                throw new InvalidOperationException("Failed to create or open the E2E notebook.");
            }

            // An existing directory may contain a closed notebook; only delete newly created directories.
            _createdNotebook = !directoryExisted;
            _createdNotebookId = notebookId;
            _notebookPath = Path.GetFullPath(notebookPath);

            return notebookId;
        }

        private string GetOrCreateSectionId(OneNoteInterop.Application app, string notebookId)
        {
            if (string.IsNullOrWhiteSpace(notebookId))
            {
                throw new InvalidOperationException("Notebook ID is required.");
            }

            app.GetHierarchy(notebookId, OneNoteInterop.HierarchyScope.hsSections, out string hierarchyXml, OneNoteInterop.XMLSchema.xs2013);
            var doc = XDocument.Parse(hierarchyXml);

            // Try to find existing E2E section (exclude recycle bin)
            var section = doc.Descendants(OneNoteNamespace + "Section")
                .FirstOrDefault(node =>
                    string.Equals((string)node.Attribute("name"), TestSectionName, StringComparison.OrdinalIgnoreCase)
                    && !IsBoolAttributeTrue(node, "isInRecycleBin"));
            if (section != null)
            {
                return (string)section.Attribute("ID");
            }

            // Create new section
            app.GetHierarchy(notebookId, OneNoteInterop.HierarchyScope.hsSelf, out string notebookXml, OneNoteInterop.XMLSchema.xs2013);
            var notebookDoc = XDocument.Parse(notebookXml);
            var notebookPath = (string)notebookDoc.Root?.Attribute("path");

            if (string.IsNullOrWhiteSpace(notebookPath))
            {
                throw new InvalidOperationException("Cannot get notebook path to create section.");
            }

            string sectionPath = Path.Combine(notebookPath, TestSectionName + ".one");
            bool sectionExisted = File.Exists(sectionPath);
            app.OpenHierarchy(sectionPath, null, out string sectionId, OneNoteInterop.CreateFileType.cftSection);

            if (string.IsNullOrWhiteSpace(sectionId))
            {
                throw new InvalidOperationException($"Failed to create section '{TestSectionName}' in notebook.");
            }

            if (sectionExisted) return sectionId;

            _createdSectionId = sectionId;
            // OneNote automatically creates an untitled page when creating a new section.
            // Capture the auto-generated page IDs immediately and delete only those.
            var autoGeneratedPageIds = WaitForSectionPageIds(app, sectionId);
            foreach (var autoPageId in autoGeneratedPageIds)
            {
                app.DeleteHierarchy(autoPageId, DateTime.MinValue, true);
            }

            return sectionId;
        }

        private void EnsureNotDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(TestPageManager));
            }
        }

        /// <summary>
        /// Gets all page IDs in the specified section.
        /// </summary>
        private List<string> GetPageIds(OneNoteInterop.Application app, string sectionId)
        {
            app.GetHierarchy(sectionId, OneNoteInterop.HierarchyScope.hsPages, out string pagesXml, OneNoteInterop.XMLSchema.xs2013);
            var doc = XDocument.Parse(pagesXml);

            return doc.Descendants(OneNoteNamespace + "Page")
                .Select(page => (string)page.Attribute("ID"))
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .ToList();
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try
            {
                // Release COM before closing a process owned by this session.
                _oneNoteDispatcher.Dispose();
                if (_shouldCloseOneNote) CloseOneNoteProcess();
            }
            finally { _ownedOneNoteProcess?.Dispose(); }
        }

    }
}
