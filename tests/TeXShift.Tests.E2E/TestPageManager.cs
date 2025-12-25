using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using OneNoteInterop = Microsoft.Office.Interop.OneNote;
using TeXShift.Core.Utils;

namespace TeXShift.Tests.E2E
{
    internal sealed class TestPageManager : IDisposable
    {
        private const string TestNotebookName = "TeXShift_E2E_Tests";
        private const string TestSectionName = "E2E";
        private static readonly XNamespace OneNoteNamespace = "http://schemas.microsoft.com/office/onenote/2013/onenote";

        private readonly StaTaskRunner _staRunner;
        private OneNoteInterop.Application _oneNoteApp;
        private bool _disposed;
        private bool _createdNotebook;  // Track if we created the notebook
        private string _notebookPath;   // Path to delete on cleanup

        private TestPageManager()
        {
            _staRunner = new StaTaskRunner("TeXShift_E2E_OneNote_STA");
            _oneNoteApp = _staRunner.Run(() =>
            {
                var appType = Type.GetTypeFromProgID("OneNote.Application");
                return (OneNoteInterop.Application)Activator.CreateInstance(appType);
            });
        }

        public static Task<TestPageManager> CreateAsync()
        {
            return Task.Run(() => new TestPageManager());
        }

        public OneNoteInterop.Application OneNoteApp => _oneNoteApp;

        public Task<string> CreateTestPageAsync(string testName, string markdownContent)
        {
            if (string.IsNullOrWhiteSpace(testName))
            {
                throw new ArgumentException("Test name is required.", nameof(testName));
            }

            return Task.Run(() => _staRunner.Run(() => CreateTestPage(testName, markdownContent ?? string.Empty)));
        }

        public Task DeletePageAsync(string pageId)
        {
            if (string.IsNullOrWhiteSpace(pageId))
            {
                return Task.CompletedTask;
            }

            return Task.Run(() => _staRunner.Run(() => DeletePage(pageId)));
        }

        /// <summary>
        /// Deletes the test notebook and section used by E2E tests.
        /// Call this after all tests are complete.
        /// </summary>
        public Task CleanupTestResourcesAsync()
        {
            return Task.Run(() => _staRunner.Run(CleanupTestResources));
        }

        private void CleanupTestResources()
        {
            EnsureNotDisposed();

            // Only cleanup if we created the notebook in this session
            if (!_createdNotebook || string.IsNullOrWhiteSpace(_notebookPath))
            {
                return;
            }

            _oneNoteApp.GetHierarchy(null, OneNoteInterop.HierarchyScope.hsNotebooks, out string hierarchyXml, OneNoteInterop.XMLSchema.xs2013);
            var doc = XDocument.Parse(hierarchyXml);

            var notebook = doc.Descendants(OneNoteNamespace + "Notebook")
                .FirstOrDefault(node => string.Equals((string)node.Attribute("name"), TestNotebookName, StringComparison.OrdinalIgnoreCase));

            if (notebook != null)
            {
                string notebookId = (string)notebook.Attribute("ID");

                if (!string.IsNullOrWhiteSpace(notebookId))
                {
                    // Close the notebook first (required before deletion)
                    _oneNoteApp.CloseNotebook(notebookId);

                    // Delete the notebook folder from disk
                    if (Directory.Exists(_notebookPath))
                    {
                        try
                        {
                            Directory.Delete(_notebookPath, recursive: true);
                        }
                        catch (IOException)
                        {
                            // Files may be locked, ignore
                        }
                        catch (UnauthorizedAccessException)
                        {
                            // No permission, ignore
                        }
                    }
                }
            }
        }

        private string CreateTestPage(string testName, string markdownContent)
        {
            EnsureNotDisposed();

            string notebookId = GetOrCreateNotebookId();
            string sectionId = GetOrCreateSectionId(notebookId);

            _oneNoteApp.CreateNewPage(sectionId, out string pageId, OneNoteInterop.NewPageStyle.npsDefault);
            if (string.IsNullOrWhiteSpace(pageId))
            {
                throw new InvalidOperationException("Failed to create a new OneNote page.");
            }

            // Update page with title and content in one operation
            // (UpdatePageContent also handles navigating to the content OE for selection)
            UpdatePageContent(pageId, testName, markdownContent);

            return pageId;
        }

        private void UpdatePageContent(string pageId, string title, string content)
        {
            _oneNoteApp.GetPageContent(pageId, out string pageXml, OneNoteInterop.PageInfo.piAll, OneNoteInterop.XMLSchema.xs2013);
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

            // Create OEChildren with one OE per line (mimics normal OneNote structure)
            // First OE is empty - used as cursor position to trigger cursor mode in reader
            var oeChildren = new XElement(ns + "OEChildren");
            var cursorOE = new XElement(ns + "OE", new XElement(ns + "T", new XCData(string.Empty)));
            oeChildren.Add(cursorOE);

            var lines = (content ?? string.Empty).Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            foreach (var line in lines)
            {
                var oe = new XElement(ns + "OE");
                var text = new XElement(ns + "T", new XCData(HtmlEscaper.Escape(line)));
                oe.Add(text);
                oeChildren.Add(oe);
            }
            outline.Add(oeChildren);

            _oneNoteApp.UpdatePageContent(doc.ToString(), DateTime.MinValue, OneNoteInterop.XMLSchema.xs2013, true);

            // Get the updated page to find the first content OE's objectID
            _oneNoteApp.GetPageContent(pageId, out string verifyXml, OneNoteInterop.PageInfo.piAll, OneNoteInterop.XMLSchema.xs2013);
            var verifyDoc = XDocument.Parse(verifyXml);

            // Find the first OE in the content outline (not the title)
            var outlineOE = verifyDoc.Descendants(OneNoteNamespace + "OE")
                .FirstOrDefault(node => !node.Ancestors(OneNoteNamespace + "Title").Any());

            string oeId = outlineOE?.Attribute("objectID")?.Value;

            // Navigate to the content OE to select it, then poll until selection is reflected
            if (!string.IsNullOrWhiteSpace(oeId))
            {
                _oneNoteApp.NavigateTo(pageId, oeId, false);
                WaitForSelection(pageId, oeId);
            }
        }

        /// <summary>
        /// Polls GetPageContent until the specified OE has a 'selected' attribute, or timeout.
        /// </summary>
        private void WaitForSelection(string pageId, string targetOeId, int timeoutMs = 10000, int pollIntervalMs = 100)
        {
            var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);

            while (DateTime.UtcNow < deadline)
            {
                _oneNoteApp.GetPageContent(pageId, out string xml, OneNoteInterop.PageInfo.piAll, OneNoteInterop.XMLSchema.xs2013);
                var doc = XDocument.Parse(xml);

                // Check if any descendant of the target OE (or the OE itself) has 'selected' attribute
                var targetOE = doc.Descendants(OneNoteNamespace + "OE")
                    .FirstOrDefault(node => (string)node.Attribute("objectID") == targetOeId);

                if (targetOE != null)
                {
                    // Check if target OE or any of its descendants have 'selected' attribute
                    bool hasSelection = targetOE.DescendantsAndSelf()
                        .Any(node => node.Attribute("selected") != null);

                    if (hasSelection)
                    {
                        return; // Selection is reflected, we're good
                    }
                }

                Thread.Sleep(pollIntervalMs);
            }

            // Timeout reached - fail fast to avoid flaky follow-up steps
            throw new TimeoutException($"WaitForSelection: Timeout waiting for selection on OE {targetOeId}");
        }

        private void DeletePage(string pageId)
        {
            EnsureNotDisposed();
            _oneNoteApp.DeleteHierarchy(pageId, DateTime.MinValue, true);
        }

        private string GetOrCreateNotebookId()
        {
            _oneNoteApp.GetHierarchy(null, OneNoteInterop.HierarchyScope.hsNotebooks, out string hierarchyXml, OneNoteInterop.XMLSchema.xs2013);
            var doc = XDocument.Parse(hierarchyXml);

            var notebook = doc.Descendants(OneNoteNamespace + "Notebook")
                .FirstOrDefault(node => string.Equals((string)node.Attribute("name"), TestNotebookName, StringComparison.OrdinalIgnoreCase));
            if (notebook != null)
            {
                // Notebook already exists - NOT created by us, don't delete on cleanup
                _createdNotebook = false;
                return (string)notebook.Attribute("ID");
            }

            _oneNoteApp.GetSpecialLocation(OneNoteInterop.SpecialLocation.slDefaultNotebookFolder, out string defaultFolder);
            if (string.IsNullOrWhiteSpace(defaultFolder))
            {
                throw new InvalidOperationException("Cannot resolve OneNote default notebook folder.");
            }

            string notebookPath = Path.Combine(defaultFolder, TestNotebookName);
            _oneNoteApp.OpenHierarchy(notebookPath, null, out string notebookId, OneNoteInterop.CreateFileType.cftNotebook);
            if (string.IsNullOrWhiteSpace(notebookId))
            {
                throw new InvalidOperationException("Failed to create or open the E2E notebook.");
            }

            // We created this notebook - safe to delete on cleanup
            _createdNotebook = true;
            _notebookPath = notebookPath;

            return notebookId;
        }

        private string GetOrCreateSectionId(string notebookId)
        {
            if (string.IsNullOrWhiteSpace(notebookId))
            {
                throw new InvalidOperationException("Notebook ID is required.");
            }

            _oneNoteApp.GetHierarchy(notebookId, OneNoteInterop.HierarchyScope.hsSections, out string hierarchyXml, OneNoteInterop.XMLSchema.xs2013);
            var doc = XDocument.Parse(hierarchyXml);

            // Try to find existing E2E section
            var section = doc.Descendants(OneNoteNamespace + "Section")
                .FirstOrDefault(node => string.Equals((string)node.Attribute("name"), TestSectionName, StringComparison.OrdinalIgnoreCase));
            if (section != null)
            {
                return (string)section.Attribute("ID");
            }

            // Create new section
            _oneNoteApp.GetHierarchy(notebookId, OneNoteInterop.HierarchyScope.hsSelf, out string notebookXml, OneNoteInterop.XMLSchema.xs2013);
            var notebookDoc = XDocument.Parse(notebookXml);
            var notebookPath = (string)notebookDoc.Root?.Attribute("path");

            if (string.IsNullOrWhiteSpace(notebookPath))
            {
                throw new InvalidOperationException("Cannot get notebook path to create section.");
            }

            string sectionPath = Path.Combine(notebookPath, TestSectionName + ".one");
            _oneNoteApp.OpenHierarchy(sectionPath, null, out string sectionId, OneNoteInterop.CreateFileType.cftSection);

            if (string.IsNullOrWhiteSpace(sectionId))
            {
                throw new InvalidOperationException($"Failed to create section '{TestSectionName}' in notebook.");
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

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            if (_oneNoteApp != null)
            {
                Task.Run(() => _staRunner.Run(() =>
                {
                    try
                    {
                        Marshal.ReleaseComObject(_oneNoteApp);
                    }
                    catch
                    {
                    }
                })).GetAwaiter().GetResult();
                _oneNoteApp = null;
            }

            _staRunner.Dispose();
        }

        private sealed class StaTaskRunner : IDisposable
        {
            private readonly Thread _staThread;
            private readonly TaskCompletionSource<bool> _ready = new TaskCompletionSource<bool>();
            private SynchronizationContext _syncContext;
            private bool _disposed;

            public StaTaskRunner(string threadName)
            {
                _staThread = new Thread(RunMessageLoop)
                {
                    IsBackground = true,
                    Name = threadName
                };
                _staThread.SetApartmentState(ApartmentState.STA);
                _staThread.Start();
                _ready.Task.GetAwaiter().GetResult();
            }

            public T Run<T>(Func<T> func)
            {
                if (_disposed)
                {
                    throw new ObjectDisposedException(nameof(StaTaskRunner));
                }

                var tcs = new TaskCompletionSource<T>();
                _syncContext.Post(_ =>
                {
                    try
                    {
                        tcs.SetResult(func());
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                    }
                }, null);

                return tcs.Task.GetAwaiter().GetResult();
            }

            public void Run(Action action)
            {
                Run(() =>
                {
                    action();
                    return true;
                });
            }

            private void RunMessageLoop()
            {
                _syncContext = new WindowsFormsSynchronizationContext();
                SynchronizationContext.SetSynchronizationContext(_syncContext);
                _ready.SetResult(true);
                Application.Run();
            }

            public void Dispose()
            {
                if (_disposed) return;
                _disposed = true;

                if (_syncContext != null)
                {
                    _syncContext.Post(_ => Application.ExitThread(), null);
                }

                if (_staThread != null && _staThread.IsAlive)
                {
                    _staThread.Join(TimeSpan.FromSeconds(2));
                }
            }
        }
    }
}
