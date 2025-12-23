using System;
using System.Collections.Concurrent;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using Xunit;

namespace TeXShift.Tests.E2E.Fixtures
{
    [CollectionDefinition("OneNoteE2E", DisableParallelization = true)]
    public class OneNoteCollection : ICollectionFixture<OneNoteFixture>
    {
    }

    public sealed class OneNoteFixture : IDisposable
    {
        private readonly StaDispatcher _dispatcher;
        private readonly ConcurrentBag<string> _createdPageIds = new ConcurrentBag<string>();
        private Application _oneNoteApp;
        private string _sectionId;

        public OneNoteFixture()
        {
            EnsureNotRunningAsAdministrator();

            _dispatcher = new StaDispatcher();
            _dispatcher.Invoke(() =>
            {
                _oneNoteApp = new Application();
                _sectionId = ResolveSectionId();
            });
        }

        public string CreateTestPage(string title)
        {
            if (string.IsNullOrWhiteSpace(title))
            {
                throw new ArgumentException("Title is required.", nameof(title));
            }

            return _dispatcher.Invoke(() =>
            {
                _oneNoteApp.CreateNewPage(_sectionId, out string pageId, NewPageStyle.npsDefault);
                _createdPageIds.Add(pageId);
                SetPageTitle(pageId, title);
                return pageId;
            });
        }

        public void UpdatePageOutline(string pageId, XElement outline)
        {
            if (string.IsNullOrWhiteSpace(pageId))
            {
                throw new ArgumentException("Page ID is required.", nameof(pageId));
            }
            if (outline == null)
            {
                throw new ArgumentNullException(nameof(outline));
            }

            _dispatcher.Invoke(() =>
            {
                var pageDoc = GetPageDocument(pageId);
                var ns = pageDoc.Root?.Name.Namespace ?? XNamespace.None;
                var existingOutline = pageDoc.Descendants(ns + "Outline").FirstOrDefault();
                var newOutline = new XElement(outline);

                if (existingOutline != null)
                {
                    existingOutline.ReplaceWith(newOutline);
                }
                else if (pageDoc.Root != null)
                {
                    pageDoc.Root.Add(newOutline);
                }

                _oneNoteApp.UpdatePageContent(pageDoc.ToString(), DateTime.MinValue, XMLSchema.xs2013, true);
            });
        }

        public XDocument GetPageContent(string pageId)
        {
            if (string.IsNullOrWhiteSpace(pageId))
            {
                throw new ArgumentException("Page ID is required.", nameof(pageId));
            }

            return _dispatcher.Invoke(() => GetPageDocument(pageId));
        }

        public void Dispose()
        {
            _dispatcher.Invoke(() =>
            {
                foreach (var pageId in _createdPageIds)
                {
                    try
                    {
                        _oneNoteApp.DeleteHierarchy(pageId, DateTime.MinValue);
                    }
                    catch
                    {
                        // Ignore cleanup failures to avoid masking test results.
                    }
                }

                if (_oneNoteApp != null)
                {
                    SafeReleaseComObject(_oneNoteApp);
                    _oneNoteApp = null;
                }
            });

            _dispatcher.Dispose();
        }

        private string ResolveSectionId()
        {
            string sectionId = null;
            Windows windows = null;
            Window window = null;
            try
            {
                windows = _oneNoteApp.Windows;
                window = windows.CurrentWindow;
                sectionId = window?.CurrentSectionId;
            }
            finally
            {
                SafeReleaseComObject(window);
                SafeReleaseComObject(windows);
            }

            if (!string.IsNullOrWhiteSpace(sectionId))
            {
                return sectionId;
            }

            _oneNoteApp.GetHierarchy(null, HierarchyScope.hsSections, out string hierarchyXml);
            var doc = XDocument.Parse(hierarchyXml);
            var ns = doc.Root?.Name.Namespace ?? "http://schemas.microsoft.com/office/onenote/2013/onenote";
            var section = doc.Descendants(ns + "Section").FirstOrDefault();
            sectionId = section?.Attribute("ID")?.Value;

            if (string.IsNullOrWhiteSpace(sectionId))
            {
                throw new InvalidOperationException("Unable to locate a OneNote section for E2E tests.");
            }

            return sectionId;
        }

        private XDocument GetPageDocument(string pageId)
        {
            _oneNoteApp.GetPageContent(pageId, out string pageXml, PageInfo.piAll, XMLSchema.xs2013);
            return XDocument.Parse(pageXml);
        }

        private void SetPageTitle(string pageId, string title)
        {
            var doc = GetPageDocument(pageId);
            var ns = doc.Root?.Name.Namespace ?? XNamespace.None;
            var titleElement = doc.Root?.Element(ns + "Title");

            if (titleElement == null)
            {
                titleElement = new XElement(ns + "Title",
                    new XElement(ns + "OE",
                        new XElement(ns + "T", new XCData(title))));
                doc.Root?.AddFirst(titleElement);
            }
            else
            {
                var oe = titleElement.Element(ns + "OE") ?? new XElement(ns + "OE");
                var t = oe.Element(ns + "T") ?? new XElement(ns + "T");
                t.ReplaceNodes(new XCData(title));
                if (t.Parent == null)
                {
                    oe.Add(t);
                }
                if (oe.Parent == null)
                {
                    titleElement.Add(oe);
                }
            }

            _oneNoteApp.UpdatePageContent(doc.ToString(), DateTime.MinValue, XMLSchema.xs2013, true);
        }

        private void EnsureNotRunningAsAdministrator()
        {
            using (var identity = WindowsIdentity.GetCurrent())
            {
                var principal = new WindowsPrincipal(identity);
                if (principal.IsInRole(WindowsBuiltInRole.Administrator))
                {
                    throw new InvalidOperationException("Do not run E2E tests as administrator.");
                }
            }
        }

        private void SafeReleaseComObject(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch
            {
                // Ignore COM release failures.
            }
        }

        private sealed class StaDispatcher : IDisposable
        {
            private readonly BlockingCollection<Action> _queue = new BlockingCollection<Action>();
            private readonly Thread _thread;

            public StaDispatcher()
            {
                _thread = new Thread(Run)
                {
                    IsBackground = true,
                    Name = "TeXShift.Tests.E2E.STA"
                };
                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }

            public void Invoke(Action action)
            {
                if (action == null)
                {
                    throw new ArgumentNullException(nameof(action));
                }

                var tcs = new TaskCompletionSource<bool>();
                _queue.Add(() =>
                {
                    try
                    {
                        action();
                        tcs.SetResult(true);
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                    }
                });

                tcs.Task.GetAwaiter().GetResult();
            }

            public T Invoke<T>(Func<T> func)
            {
                if (func == null)
                {
                    throw new ArgumentNullException(nameof(func));
                }

                var tcs = new TaskCompletionSource<T>();
                _queue.Add(() =>
                {
                    try
                    {
                        tcs.SetResult(func());
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                    }
                });

                return tcs.Task.GetAwaiter().GetResult();
            }

            private void Run()
            {
                foreach (var action in _queue.GetConsumingEnumerable())
                {
                    action();
                }
            }

            public void Dispose()
            {
                _queue.CompleteAdding();
                _thread.Join(TimeSpan.FromSeconds(5));
                _queue.Dispose();
            }
        }
    }
}
