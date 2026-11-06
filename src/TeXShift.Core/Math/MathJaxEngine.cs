using System;
using System.IO;
using System.Reflection;
using System.Xml.Linq;
using System.Web.Script.Serialization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using TeXShift.Core.Errors;
using TeXShift.Core.Localization;
using TeXShift.Core.Logging;
using TeXShift.Core.Utils;

namespace TeXShift.Core.Math
{
    /// <summary>
    /// Converts LaTeX math expressions to MathML using WebView2 and MathJax.
    /// Uses a dedicated STA thread to ensure WebView2 compatibility.
    /// </summary>
    internal class MathJaxEngine : IMathEngine
    {
        private readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        private WebView2 _webView;
        private bool _isInitialized;
        private volatile bool _isDisposed;
        private readonly object _lifecycle = new object();
        private Exception _initFailure;
        private readonly SemaphoreSlim _initLock = new SemaphoreSlim(1, 1);

        private StaWorker _worker;
        private string _userDataFolder;

        private async Task InitializeAsync()
        {
            if (_isDisposed) throw new ObjectDisposedException(GetType().Name);
            if (_isInitialized) return;
            if (_initFailure != null) throw _initFailure;

            await _initLock.WaitAsync().ConfigureAwait(false);
            try
            {
                if (_isInitialized) return;
                if (_initFailure != null) throw _initFailure;

                try
                {
                    StaWorker worker;
                    lock (_lifecycle)
                    {
                        if (_isDisposed) throw new ObjectDisposedException(GetType().Name);
                        worker = _worker = new StaWorker("TeXShift_Math_WebView2_STA", DisposeWebView);
                    }
                    await worker.InvokeAsync(async () =>
                    {
                        await InitializeWebView2Async();
                        return true;
                    }).ConfigureAwait(false);
                    lock (_lifecycle)
                    {
                        if (_isDisposed) throw new ObjectDisposedException(GetType().Name);
                        _isInitialized = true;
                    }
                }
                catch (Exception ex)
                {
                    StopWorker();
                    if (ex is TeXShiftException)
                    {
                        _initFailure = ex;
                        throw;
                    }

                    _initFailure = new MathConversionException(
                        Resources.GetString("Error_MathInitFailed"),
                        ex.Message,
                        ex);
                    throw _initFailure;
                }
            }
            finally
            {
                _initLock.Release();
            }
        }

        private async Task InitializeWebView2Async()
        {
            _webView = new WebView2();
            _webView.Visible = false;

            var userDataFolder = CreateUserDataFolder();

            CoreWebView2Environment env;
            try
            {
                using (PerformanceTraceContext.Measure("Math.WebView2.CreateEnvironment"))
                    env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                using (PerformanceTraceContext.Measure("Math.WebView2.EnsureControl"))
                    await _webView.EnsureCoreWebView2Async(env);
            }
            catch (WebView2RuntimeNotFoundException ex)
            {
                throw new MathConversionException(
                    Resources.GetString("Error_WebView2NotInstalled"),
                    "WebView2 Runtime not found on this system.",
                    ex);
            }

            // Find MathJax folder
            var mathjaxPath = FindMathJaxPath();
            if (string.IsNullOrEmpty(mathjaxPath))
            {
                throw new InvalidOperationException(
                    "MathJax not found. Expected at Lib/mathjax relative to assembly or project root.");
            }

            // Generate loader HTML with file:// URL to MathJax
            var mathjaxFileUrl = "file:///" + Path.Combine(mathjaxPath, "es5", "tex-mml-svg.js").Replace('\\', '/');
            var html = GetMathJaxLoaderHtml().Replace(
                "https://mathjax.local/es5/tex-mml-svg.js",
                mathjaxFileUrl);

            // Write loader HTML to temp folder and navigate via file://
            var loaderPath = Path.Combine(userDataFolder, "loader.html");
            File.WriteAllText(loaderPath, html);

            // Wait for navigation to complete
            var navTcs = new TaskCompletionSource<bool>();
            void OnNavigationCompleted(object s, CoreWebView2NavigationCompletedEventArgs e)
            {
                _webView.CoreWebView2.NavigationCompleted -= OnNavigationCompleted;
                if (e.IsSuccess)
                    navTcs.SetResult(true);
                else
                    navTcs.SetException(new Exception($"Navigation failed: {e.WebErrorStatus}"));
            }
            _webView.CoreWebView2.NavigationCompleted += OnNavigationCompleted;
            using (PerformanceTraceContext.Measure("Math.WebView2.Navigate"))
            {
                _webView.CoreWebView2.Navigate("file:///" + loaderPath.Replace('\\', '/'));
                await navTcs.Task;
            }

            // Wait for MathJax to be ready
            using (PerformanceTraceContext.Measure("Math.WaitForReady"))
                await WaitForMathJaxReady();
        }

        public async Task<MathExpression> ParseAsync(string latex, bool displayMode)
        {
            var xml = await InvokeAsync<string>("parse", new { latex, displayMode }).ConfigureAwait(false);
            var mathml = XElement.Parse(xml);
            var error = System.Linq.Enumerable.FirstOrDefault(mathml.Descendants(mathml.Name.Namespace + "merror"));
            if (error != null)
                throw new MathConversionException(Resources.GetString("Error_MathConversionFailed"), error.Value);
            return new MathExpression(latex, displayMode, mathml);
        }

        public Task<MathImage> RenderImageAsync(MathExpression expression, double fontSizePt) =>
            InvokeAsync<MathImage>("image", new { mathml = expression.CopyMathML().ToString(SaveOptions.DisableFormatting), fontSizePt });

        private async Task<T> InvokeAsync<T>(string operation, object arguments)
        {
            using (PerformanceTraceContext.Measure("Math.Initialize"))
                await InitializeAsync().ConfigureAwait(false);
            try
            {
                using (PerformanceTraceContext.Measure("Math.Request", operation))
                return await _worker.InvokeAsync(async () =>
                {
                    var completion = new TaskCompletionSource<T>();
                    var id = Guid.NewGuid().ToString("N");
                    void Receive(object sender, CoreWebView2WebMessageReceivedEventArgs args)
                    {
                        var message = _json.Deserialize<MathResponse>(args.WebMessageAsJson);
                        if (message.Id != id) return;
                        if (message.Error != null) completion.TrySetException(new MathConversionException(
                            Resources.GetString("Error_MathConversionFailed"), message.Error));
                        else completion.TrySetResult(_json.ConvertToType<T>(message.Result));
                    }
                    _webView.CoreWebView2.WebMessageReceived += Receive;
                    try
                    {
                        var request = _json.Serialize(new { id, operation, arguments });
                        await _webView.CoreWebView2.ExecuteScriptAsync($"executeMathRequest({request})");
                        if (await Task.WhenAny(completion.Task, Task.Delay(30000)) != completion.Task)
                            throw new TimeoutException("MathJax request timed out.");
                        return await completion.Task;
                    }
                    finally { _webView.CoreWebView2.WebMessageReceived -= Receive; }
                }).ConfigureAwait(false);
            }
            catch (Exception ex) when (!(ex is TeXShiftException))
            {
                throw new MathConversionException(Resources.GetString("Error_MathConversionFailed"), ex.Message, ex);
            }
        }

        private sealed class MathResponse
        {
            public string Id { get; set; }
            public string Error { get; set; }
            public object Result { get; set; }
        }

        /// <summary>
        /// Finds the MathJax folder path, checking multiple locations for dev/production environments.
        /// </summary>
        private string FindMathJaxPath()
        {
            var assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            // Production: Lib folder next to DLL
            var prodPath = Path.Combine(assemblyDir, "Lib", "mathjax");
            if (Directory.Exists(prodPath))
            {
                return prodPath;
            }

            // Development: Walk up to find project root, then check known locations
            var dir = new DirectoryInfo(assemblyDir);
            while (dir != null)
            {
                // Check direct Lib folder (Add-in project structure)
                var devPath = Path.Combine(dir.FullName, "Lib", "mathjax");
                if (Directory.Exists(devPath))
                {
                    return devPath;
                }

                // Check src/TeXShift.AddIn/Lib (for E2E tests and other projects)
                var addInLibPath = Path.Combine(dir.FullName, "src", "TeXShift.AddIn", "Lib", "mathjax");
                if (Directory.Exists(addInLibPath))
                {
                    return addInLibPath;
                }

                dir = dir.Parent;
            }

            return null;
        }

        private string GetMathJaxLoaderHtml()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "TeXShift.Core.Resources.Math.mathjax-loader.html";

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream != null)
                {
                    using (var reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }

            throw new InvalidOperationException(
                "MathJax loader HTML not found. Expected embedded resource 'TeXShift.Core.Resources.Math.mathjax-loader.html'.");
        }

        private async Task WaitForMathJaxReady()
        {
            var maxWaitMs = 30000; // 30 seconds timeout
            var intervalMs = 100;
            var elapsed = 0;

            while (elapsed < maxWaitMs)
            {
                var result = await _webView.CoreWebView2.ExecuteScriptAsync("isMathJaxReady()");
                if (result == "true")
                {
                    return;
                }

                await Task.Delay(intervalMs);
                elapsed += intervalMs;
            }

            throw new TimeoutException("MathJax failed to initialize within timeout period.");
        }


        public void Dispose()
        {
            lock (_lifecycle)
            {
                _isDisposed = true;
                _isInitialized = false;
            }
            StopWorker();
            // The semaphore can still have initialization waiters; let GC reclaim it.
        }

        private void StopWorker()
        {
            StaWorker worker;
            lock (_lifecycle)
            {
                worker = _worker;
                _worker = null;
            }

            worker?.Dispose();
            DeleteUserDataFolder();
        }

        private string CreateUserDataFolder()
        {
            var root = Path.Combine(Path.GetTempPath(), "TeXShift_WebView2");
            var folder = Path.Combine(root, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            _userDataFolder = folder;
            return folder;
        }

        private void DeleteUserDataFolder()
        {
            var folder = Interlocked.Exchange(ref _userDataFolder, null);
            if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder)) return;

            try
            {
                Directory.Delete(folder, true);
            }
            catch (IOException)
            {
                // WebView2 may release profile files shortly after its process exits.
            }
            catch (UnauthorizedAccessException)
            {
                // A locked profile can be removed by the operating system later.
            }
        }

        private void DisposeWebView()
        {
            _webView?.Dispose();
            _webView = null;
        }

    }
}
