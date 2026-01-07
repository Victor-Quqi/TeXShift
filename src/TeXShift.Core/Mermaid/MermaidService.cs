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

            await _initLock.WaitAsync().ConfigureAwait(false);
            try
            {
                if (_isInitialized) return;

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
                catch (Exception ex) when (!(ex is TeXShiftException))
                {
                    throw new MermaidConversionException(
                        Resources.GetString("Error_MermaidInitFailed"),
                        ex.Message,
                        ex);
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
            var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder).ConfigureAwait(false);
            await _webView.EnsureCoreWebView2Async(env).ConfigureAwait(false);

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
            await navTcs.Task.ConfigureAwait(false);

            await WaitForMermaidReady().ConfigureAwait(false);
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
            catch
            {
                // Ignore malformed messages
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

            // Fallback: load from disk in dev scenarios (when the embedded resource wasn't added in VS)
            var diskPath = FindMermaidLoaderHtmlPath();
            if (!string.IsNullOrEmpty(diskPath) && File.Exists(diskPath))
            {
                return File.ReadAllText(diskPath, Encoding.UTF8);
            }

            // Last resort: inline loader (placeholder URL will be replaced with file://)
            return @"<!DOCTYPE html>
<html lang=""en"">
<head>
    <meta charset=""UTF-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1"">
    <title>TeXShift Mermaid Loader</title>
    <script src=""https://mermaid.local/mermaid.min.js""></script>
</head>
<body>
<div id=""texshift-mermaid-container"" style=""position: fixed; left: -100000px; top: -100000px; width: 1px; height: 1px; overflow: hidden;""></div>
<script>
    var mermaidReady = false;

    function safeMermaidInitialize(theme) {
        if (typeof mermaid === 'undefined' || !mermaid || typeof mermaid.initialize !== 'function') {
            return;
        }
        try {
            mermaid.initialize({
                startOnLoad: false,
                theme: theme || 'default',
                securityLevel: 'loose',
                flowchart: {
                    htmlLabels: false
                },
                sequence: {
                    useMaxWidth: false
                }
            });
        } catch (e) {
        }
    }

    function isMermaidReady() {
        return mermaidReady;
    }

    function getFirstNumber(value) {
        if (!value) return 0;
        var m = String(value).match(/-?\d+(\.\d+)?/);
        return m ? parseFloat(m[0]) : 0;
    }

    function getSvgDimensions(svgEl) {
        var width = getFirstNumber(svgEl.getAttribute('width'));
        var height = getFirstNumber(svgEl.getAttribute('height'));
        if (width > 0 && height > 0) {
            return { width: width, height: height };
        }

        var viewBox = svgEl.getAttribute('viewBox');
        if (viewBox) {
            var parts = viewBox.trim().split(/\s+/);
            if (parts.length === 4) {
                var vbW = parseFloat(parts[2]);
                var vbH = parseFloat(parts[3]);
                if (vbW > 0 && vbH > 0) {
                    return { width: vbW, height: vbH };
                }
            }
        }

        try {
            var container = document.getElementById('texshift-mermaid-container');
            container.innerHTML = '';
            container.appendChild(svgEl);

            var rect = svgEl.getBoundingClientRect();
            if (rect && rect.width > 0 && rect.height > 0) {
                container.innerHTML = '';
                return { width: rect.width, height: rect.height };
            }

            var bbox = svgEl.getBBox();
            container.innerHTML = '';
            if (bbox && bbox.width > 0 && bbox.height > 0) {
                return { width: bbox.width, height: bbox.height };
            }
        } catch (e) {
        }

        return { width: 800, height: 600 };
    }

    function computeTargetSize(width, height, maxWidth, maxHeight) {
        var w = Math.max(1, width || 1);
        var h = Math.max(1, height || 1);
        var scale = 1;
        if (maxWidth && maxWidth > 0) scale = Math.min(scale, maxWidth / w);
        if (maxHeight && maxHeight > 0) scale = Math.min(scale, maxHeight / h);
        if (!isFinite(scale) || scale <= 0) scale = 1;
        if (scale > 1) scale = 1;
        return {
            width: Math.max(1, Math.floor(w * scale)),
            height: Math.max(1, Math.floor(h * scale))
        };
    }

    async function svgToPngBase64(svgText, outWidth, outHeight) {
        var parser = new DOMParser();
        var svgDoc = parser.parseFromString(svgText, 'image/svg+xml');
        var svgEl = svgDoc.querySelector('svg');

        if (!svgEl) {
            throw new Error('Invalid SVG');
        }

        var foreignObjects = svgEl.querySelectorAll('foreignObject');
        foreignObjects.forEach(function(fo) {
            var text = fo.textContent || '';
            var parent = fo.parentNode;
            if (parent && text.trim()) {
                var textEl = document.createElementNS('http://www.w3.org/2000/svg', 'text');
                textEl.textContent = text.trim();
                var x = fo.getAttribute('x') || '0';
                var y = fo.getAttribute('y') || '0';
                textEl.setAttribute('x', x);
                textEl.setAttribute('y', y);
                textEl.setAttribute('font-size', '14');
                textEl.setAttribute('fill', '#333');
                parent.replaceChild(textEl, fo);
            } else {
                fo.remove();
            }
        });

        var serializer = new XMLSerializer();
        var processedSvgText = serializer.serializeToString(svgEl);

        var encodedSvg = encodeURIComponent(processedSvgText)
            .replace(/'/g, '%27')
            .replace(/""/g, '%22');
        var dataUrl = 'data:image/svg+xml;charset=utf-8,' + encodedSvg;

        return await new Promise(function (resolve, reject) {
            var img = new Image();
            img.onload = function () {
                try {
                    var canvas = document.createElement('canvas');
                    canvas.width = outWidth;
                    canvas.height = outHeight;
                    var ctx = canvas.getContext('2d');
                    ctx.imageSmoothingEnabled = true;
                    ctx.imageSmoothingQuality = 'high';
                    ctx.drawImage(img, 0, 0, outWidth, outHeight);
                    var pngDataUrl = canvas.toDataURL('image/png');
                    var parts = String(pngDataUrl).split(',');
                    resolve(parts.length > 1 ? parts[1] : '');
                } catch (e) {
                    reject(e);
                }
            };
            img.onerror = function (e) {
                reject(new Error('Failed to load SVG image'));
            };
            img.src = dataUrl;
        });
    }

    async function renderMermaid(code, options) {
        try {
            if (!code || !String(code).trim()) {
                return { success: false, errorMessage: 'Empty Mermaid code.' };
            }

            options = options || {};
            var theme = options.Theme || options.theme || 'default';
            var maxWidth = options.MaxWidth || options.maxWidth || 1920;
            var maxHeight = options.MaxHeight || options.maxHeight || 1080;

            safeMermaidInitialize(theme);

            var id = 'texshift-mermaid-' + Date.now().toString(36) + '-' + Math.random().toString(16).slice(2);

            var svg = null;
            try {
                var r = mermaid.render(id, code);
                if (r && typeof r.then === 'function') {
                    r = await r;
                }
                if (typeof r === 'string') {
                    svg = r;
                } else if (r && r.svg) {
                    svg = r.svg;
                }
            } catch (e) {
                return { success: false, errorMessage: (e && e.message) ? e.message : String(e) };
            }

            if (!svg) {
                return { success: false, errorMessage: 'Mermaid did not produce SVG.' };
            }

            var tmp = document.createElement('div');
            tmp.innerHTML = svg;
            var svgEl = tmp.querySelector('svg');
            if (!svgEl) {
                return { success: false, errorMessage: 'No SVG element in Mermaid output.' };
            }

            if (!svgEl.getAttribute('xmlns')) {
                svgEl.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
            }

            var dims = getSvgDimensions(svgEl);
            svgEl.setAttribute('width', dims.width + 'px');
            svgEl.setAttribute('height', dims.height + 'px');

            var serializer = new XMLSerializer();
            var svgText = serializer.serializeToString(svgEl);

            var target = computeTargetSize(dims.width, dims.height, maxWidth, maxHeight);
            var base64 = await svgToPngBase64(svgText, target.width, target.height);
            if (!base64) {
                return { success: false, errorMessage: 'PNG conversion produced empty data.' };
            }

            return {
                success: true,
                base64PngData: base64,
                width: target.width,
                height: target.height,
                errorMessage: ''
            };
        } catch (e) {
            var msg = (e && e.message) ? e.message : String(e);
            return { success: false, errorMessage: msg };
        }
    }

    (function () {
        if (!window.chrome || !window.chrome.webview) return;

        window.chrome.webview.addEventListener('message', function (event) {
            var msg = event.data;
            if (!msg || msg.type !== 'renderMermaid' || !msg.id) return;

            (async function () {
                var result = await renderMermaid(msg.code, msg.options);
                window.chrome.webview.postMessage({
                    type: 'renderMermaidResult',
                    id: msg.id,
                    result: result
                });
            })().catch(function (err) {
                window.chrome.webview.postMessage({
                    type: 'renderMermaidResult',
                    id: msg.id,
                    result: {
                        success: false,
                        errorMessage: (err && err.message) ? err.message : String(err)
                    }
                });
            });
        });
    })();

    try {
        if (typeof mermaid !== 'undefined') {
            safeMermaidInitialize('default');
            mermaidReady = true;
        }
    } catch (e) {
        mermaidReady = false;
    }
</script>
</body>
</html>";
        }

        private string FindMermaidLoaderHtmlPath()
        {
            var assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (string.IsNullOrEmpty(assemblyDir))
            {
                return null;
            }

            // If someone copied Resources folder next to DLL
            var prodPath = Path.Combine(assemblyDir, "Resources", "Mermaid", "mermaid-loader.html");
            if (File.Exists(prodPath))
            {
                return prodPath;
            }

            var dir = new DirectoryInfo(assemblyDir);
            while (dir != null)
            {
                var devPath = Path.Combine(dir.FullName, "src", "TeXShift.Core", "Resources", "Mermaid", "mermaid-loader.html");
                if (File.Exists(devPath))
                {
                    return devPath;
                }

                var resourcesPath = Path.Combine(dir.FullName, "Resources", "Mermaid", "mermaid-loader.html");
                if (File.Exists(resourcesPath))
                {
                    return resourcesPath;
                }

                dir = dir.Parent;
            }

            return null;
        }

        private async Task WaitForMermaidReady()
        {
            var maxWaitMs = 30000;
            var intervalMs = 100;
            var elapsed = 0;

            while (elapsed < maxWaitMs)
            {
                var result = await _webView.CoreWebView2.ExecuteScriptAsync("isMermaidReady()").ConfigureAwait(false);
                if (result == "true")
                {
                    return;
                }

                await Task.Delay(intervalMs).ConfigureAwait(false);
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
