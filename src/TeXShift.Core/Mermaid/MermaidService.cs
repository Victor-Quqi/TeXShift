using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using TeXShift.Core.Errors;
using TeXShift.Core.Localization;

namespace TeXShift.Core.Mermaid
{
    /// <summary>
    /// Renders Mermaid diagram code to PNG images using WebView2 and Mermaid.js.
    /// Uses a dedicated STA thread to ensure WebView2 compatibility.
    /// </summary>
    internal class MermaidService : IMermaidService
    {
        private WebView2 _webView;
        private bool _isInitialized;
        private bool _isDisposed;
        private Exception _initFailure;
        private readonly SemaphoreSlim _initLock = new SemaphoreSlim(1, 1);

        // STA thread for WebView2 operations
        private Thread _staThread;
        private TaskCompletionSource<bool> _staReady;
        private SynchronizationContext _staSyncContext;

        private readonly object _pendingLock = new object();
        private readonly Dictionary<string, TaskCompletionSource<MermaidRenderResult>> _pendingRequests =
            new Dictionary<string, TaskCompletionSource<MermaidRenderResult>>(StringComparer.Ordinal);

        public bool IsInitialized => _isInitialized;

        public async Task InitializeAsync()
        {
            if (_isInitialized) return;
            if (_initFailure != null) throw _initFailure;

            await _initLock.WaitAsync().ConfigureAwait(false);
            try
            {
                if (_isInitialized) return;
                if (_initFailure != null) throw _initFailure;

                try
                {
                    _staReady = new TaskCompletionSource<bool>();
                    _staThread = new Thread(StaThreadStart);
                    _staThread.SetApartmentState(ApartmentState.STA);
                    _staThread.IsBackground = true;
                    _staThread.Name = "TeXShift_Mermaid_WebView2_STA";
                    _staThread.Start();

                    await _staReady.Task.ConfigureAwait(false);

                    var initTcs = new TaskCompletionSource<bool>();
                    _staSyncContext.Post(async _ =>
                    {
                        try
                        {
                            await InitializeWebView2Async().ConfigureAwait(false);
                            initTcs.SetResult(true);
                        }
                        catch (Exception ex)
                        {
                            initTcs.SetException(ex);
                        }
                    }, null);

                    await initTcs.Task.ConfigureAwait(false);
                    _isInitialized = true;
                }
                catch (Exception ex)
                {
                    // Cache the failure so subsequent calls fail fast without retrying
                    // expensive STA thread + WebView2 initialization.
                    if (ex is TeXShiftException)
                    {
                        _initFailure = ex;
                        throw;
                    }

                    _initFailure = new MermaidConversionException(
                        Resources.GetString("Error_MermaidInitFailed"),
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

        private void StaThreadStart()
        {
            // Create and install a synchronization context for this STA thread
            var form = new Form { Visible = false };
            _staSyncContext = SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();
            SynchronizationContext.SetSynchronizationContext(_staSyncContext);

            _staReady.SetResult(true);

            // Run message loop to keep thread alive and process messages
            Application.Run();
        }

        private async Task InitializeWebView2Async()
        {
            _webView = new WebView2();
            _webView.Visible = false;

            var userDataFolder = Path.Combine(Path.GetTempPath(), "TeXShift_Mermaid_WebView2");
            Directory.CreateDirectory(userDataFolder);

            CoreWebView2Environment env;
            try
            {
                env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                await _webView.EnsureCoreWebView2Async(env);
            }
            catch (WebView2RuntimeNotFoundException ex)
            {
                throw new MermaidConversionException(
                    Resources.GetString("Error_WebView2NotInstalled"),
                    "WebView2 Runtime not found on this system.",
                    ex);
            }

            _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

            var mermaidScriptPath = FindMermaidScriptPath();
            if (string.IsNullOrEmpty(mermaidScriptPath))
            {
                throw new InvalidOperationException(
                    "Mermaid not found. Expected at Lib/mermaid/mermaid.min.js relative to assembly or project root.");
            }

            var mermaidFileUrl = "file:///" + mermaidScriptPath.Replace('\\', '/');
            var html = GetMermaidLoaderHtml().Replace(
                "https://mermaid.local/mermaid.min.js",
                mermaidFileUrl);

            var loaderPath = Path.Combine(userDataFolder, "mermaid-loader.html");
            File.WriteAllText(loaderPath, html, Encoding.UTF8);

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
            _webView.CoreWebView2.Navigate("file:///" + loaderPath.Replace('\\', '/'));
            await navTcs.Task;

            await WaitForMermaidReady();
        }

        public async Task<MermaidRenderResult> RenderToImageAsync(string mermaidCode, MermaidRenderOptions options = null)
        {
            if (!_isInitialized)
            {
                await InitializeAsync().ConfigureAwait(false);
            }

            if (string.IsNullOrWhiteSpace(mermaidCode))
            {
                return new MermaidRenderResult
                {
                    Success = false,
                    ErrorMessage = Resources.GetString("Error_MermaidConversionFailed")
                };
            }

            options = options ?? new MermaidRenderOptions();

            try
            {
                var tcs = new TaskCompletionSource<MermaidRenderResult>();
                _staSyncContext.Post(async _ =>
                {
                    try
                    {
                        var result = await RenderToImageOnStaAsync(mermaidCode, options).ConfigureAwait(false);
                        tcs.SetResult(result);
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                    }
                }, null);

                return await tcs.Task.ConfigureAwait(false);
            }
            catch (Exception ex) when (!(ex is TeXShiftException))
            {
                return new MermaidRenderResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        private async Task<MermaidRenderResult> RenderToImageOnStaAsync(string mermaidCode, MermaidRenderOptions options)
        {
            var requestId = Guid.NewGuid().ToString("N");
            var responseTcs = new TaskCompletionSource<MermaidRenderResult>();

            lock (_pendingLock)
            {
                _pendingRequests[requestId] = responseTcs;
            }

            var messageJson = BuildRenderRequestJson(requestId, mermaidCode, options);
            _webView.CoreWebView2.PostWebMessageAsJson(messageJson);

            var timeoutMs = 10000;
            var completed = await Task.WhenAny(responseTcs.Task, Task.Delay(timeoutMs)).ConfigureAwait(false);
            if (completed != responseTcs.Task)
            {
                lock (_pendingLock)
                {
                    _pendingRequests.Remove(requestId);
                }

                return new MermaidRenderResult
                {
                    Success = false,
                    ErrorMessage = "Mermaid render timed out."
                };
            }

            return await responseTcs.Task.ConfigureAwait(false);
        }

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                var json = e.WebMessageAsJson;
                var message = DeserializeJson<RenderMermaidResultMessage>(json);
                if (message == null || !string.Equals(message.Type, "renderMermaidResult", StringComparison.Ordinal))
                {
                    return;
                }

                if (string.IsNullOrEmpty(message.Id) || message.Result == null)
                {
                    return;
                }

                TaskCompletionSource<MermaidRenderResult> tcs = null;
                lock (_pendingLock)
                {
                    if (_pendingRequests.TryGetValue(message.Id, out tcs))
                    {
                        _pendingRequests.Remove(message.Id);
                    }
                }

                if (tcs == null)
                {
                    return;
                }

                var result = new MermaidRenderResult
                {
                    Success = message.Result.Success,
                    Base64PngData = message.Result.Base64PngData,
                    Width = message.Result.Width,
                    Height = message.Result.Height,
                    ErrorMessage = message.Result.ErrorMessage
                };

                tcs.SetResult(result);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TeXShift] Malformed WebMessage ignored: {ex.Message}");
            }
        }

        private string FindMermaidScriptPath()
        {
            var assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (string.IsNullOrEmpty(assemblyDir))
            {
                return null;
            }

            // Production: Lib folder next to DLL
            var prodPath = Path.Combine(assemblyDir, "Lib", "mermaid", "mermaid.min.js");
            if (File.Exists(prodPath))
            {
                return prodPath;
            }

            // Development: Walk up to find project root, then check known locations
            var dir = new DirectoryInfo(assemblyDir);
            while (dir != null)
            {
                var devPath = Path.Combine(dir.FullName, "Lib", "mermaid", "mermaid.min.js");
                if (File.Exists(devPath))
                {
                    return devPath;
                }

                var addInLibPath = Path.Combine(dir.FullName, "src", "TeXShift.AddIn", "Lib", "mermaid", "mermaid.min.js");
                if (File.Exists(addInLibPath))
                {
                    return addInLibPath;
                }

                dir = dir.Parent;
            }

            return null;
        }

        private string GetMermaidLoaderHtml()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "TeXShift.Core.Resources.Mermaid.mermaid-loader.html";

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream != null)
                {
                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }

            throw new InvalidOperationException(
                "Mermaid loader HTML not found. Expected embedded resource 'TeXShift.Core.Resources.Mermaid.mermaid-loader.html'.");
        }

        private async Task WaitForMermaidReady()
        {
            var maxWaitMs = 30000;
            var intervalMs = 100;
            var elapsed = 0;

            while (elapsed < maxWaitMs)
            {
                var result = await _webView.CoreWebView2.ExecuteScriptAsync("isMermaidReady()");
                if (result == "true")
                {
                    return;
                }

                await Task.Delay(intervalMs);
                elapsed += intervalMs;
            }

            throw new TimeoutException("Mermaid failed to initialize within timeout period.");
        }

        private static T DeserializeJson<T>(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return default(T);

            var serializer = new DataContractJsonSerializer(typeof(T));
            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                return (T)serializer.ReadObject(ms);
            }
        }

        private static string BuildRenderRequestJson(string id, string code, MermaidRenderOptions options)
        {
            var maxWidth = options?.MaxWidth ?? 1920;
            var maxHeight = options?.MaxHeight ?? 1080;
            var theme = options?.Theme ?? "default";

            return "{"
                   + "\"type\":\"renderMermaid\","
                   + "\"id\":\"" + EscapeForJsonString(id) + "\","
                   + "\"code\":\"" + EscapeForJsonString(code) + "\","
                   + "\"options\":{"
                   + "\"MaxWidth\":" + maxWidth.ToString(CultureInfo.InvariantCulture) + ","
                   + "\"MaxHeight\":" + maxHeight.ToString(CultureInfo.InvariantCulture) + ","
                   + "\"Theme\":\"" + EscapeForJsonString(theme) + "\""
                   + "}"
                   + "}";
        }

        private static string EscapeForJsonString(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return string.Empty;
            }

            var sb = new StringBuilder(input.Length + 16);
            foreach (var c in input)
            {
                switch (c)
                {
                    case '\\':
                        sb.Append("\\\\");
                        break;
                    case '"':
                        sb.Append("\\\"");
                        break;
                    case '\b':
                        sb.Append("\\b");
                        break;
                    case '\f':
                        sb.Append("\\f");
                        break;
                    case '\n':
                        sb.Append("\\n");
                        break;
                    case '\r':
                        sb.Append("\\r");
                        break;
                    case '\t':
                        sb.Append("\\t");
                        break;
                    default:
                        if (c < ' ')
                        {
                            sb.Append("\\u");
                            sb.Append(((int)c).ToString("X4", CultureInfo.InvariantCulture));
                        }
                        else
                        {
                            sb.Append(c);
                        }
                        break;
                }
            }

            return sb.ToString();
        }

        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;

            if (_staSyncContext != null)
            {
                _staSyncContext.Post(_ =>
                {
                    try
                    {
                        if (_webView?.CoreWebView2 != null)
                        {
                            _webView.CoreWebView2.WebMessageReceived -= OnWebMessageReceived;
                        }
                        _webView?.Dispose();
                    }
                    finally
                    {
                        Application.ExitThread();
                    }
                }, null);
            }

            _initLock?.Dispose();
        }

        [DataContract]
        private class RenderMermaidResultMessage
        {
            [DataMember(Name = "type")]
            public string Type { get; set; }

            [DataMember(Name = "id")]
            public string Id { get; set; }

            [DataMember(Name = "result")]
            public RenderMermaidResultPayload Result { get; set; }
        }

        [DataContract]
        private class RenderMermaidResultPayload
        {
            [DataMember(Name = "success")]
            public bool Success { get; set; }

            [DataMember(Name = "base64PngData")]
            public string Base64PngData { get; set; }

            [DataMember(Name = "width")]
            public int Width { get; set; }

            [DataMember(Name = "height")]
            public int Height { get; set; }

            [DataMember(Name = "errorMessage")]
            public string ErrorMessage { get; set; }
        }
    }
}
