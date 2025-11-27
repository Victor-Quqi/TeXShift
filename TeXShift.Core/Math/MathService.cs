using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace TeXShift.Core.Math
{
    /// <summary>
    /// Converts LaTeX math expressions to MathML using WebView2 and MathJax.
    /// Uses a dedicated STA thread to ensure WebView2 compatibility.
    /// </summary>
    internal class MathService : IMathService
    {
        private WebView2 _webView;
        private bool _isInitialized;
        private bool _isDisposed;
        private readonly SemaphoreSlim _initLock = new SemaphoreSlim(1, 1);

        // STA thread for WebView2 operations
        private Thread _staThread;
        private TaskCompletionSource<bool> _staReady;
        private SynchronizationContext _staSyncContext;

        // Regex to add mml: namespace prefix to MathML elements
        private static readonly Regex MathMLElementRegex = new Regex(
            @"<(/?)(math|mi|mo|mn|ms|mtext|mspace|mglyph|maligngroup|malignmark|" +
            @"mrow|mfrac|msqrt|mroot|mstyle|merror|mpadded|mphantom|mfenced|menclose|" +
            @"msub|msup|msubsup|munder|mover|munderover|mmultiscripts|mtable|mlabeledtr|" +
            @"mtr|mtd|maction|semantics|annotation|annotation-xml)(\s|>|/>)",
            RegexOptions.Compiled);

        public bool IsInitialized => _isInitialized;

        public async Task InitializeAsync()
        {
            if (_isInitialized) return;

            await _initLock.WaitAsync().ConfigureAwait(false);
            try
            {
                if (_isInitialized) return;

                // Start dedicated STA thread for WebView2
                _staReady = new TaskCompletionSource<bool>();
                _staThread = new Thread(StaThreadStart);
                _staThread.SetApartmentState(ApartmentState.STA);
                _staThread.IsBackground = true;
                _staThread.Name = "TeXShift_WebView2_STA";
                _staThread.Start();

                // Wait for STA thread to be ready
                await _staReady.Task.ConfigureAwait(false);

                // Initialize WebView2 on STA thread
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

            // Initialize WebView2 environment
            var userDataFolder = Path.Combine(Path.GetTempPath(), "TeXShift_WebView2");
            Directory.CreateDirectory(userDataFolder);
            var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder).ConfigureAwait(false);
            await _webView.EnsureCoreWebView2Async(env).ConfigureAwait(false);

            // Find MathJax folder
            var mathjaxPath = FindMathJaxPath();
            if (string.IsNullOrEmpty(mathjaxPath))
            {
                throw new InvalidOperationException(
                    "MathJax not found. Expected at Lib/mathjax relative to assembly or project root.");
            }

            // Generate loader HTML with file:// URL to MathJax
            var mathjaxFileUrl = "file:///" + Path.Combine(mathjaxPath, "es5", "tex-mml-chtml.js").Replace('\\', '/');
            var html = GetMathJaxLoaderHtml().Replace(
                "https://mathjax.local/es5/tex-mml-chtml.js",
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
            _webView.CoreWebView2.Navigate("file:///" + loaderPath.Replace('\\', '/'));
            await navTcs.Task.ConfigureAwait(false);

            // Wait for MathJax to be ready
            await WaitForMathJaxReady().ConfigureAwait(false);
        }

        public async Task<string> LatexToMathMLAsync(string latex, bool displayMode)
        {
            if (!_isInitialized)
            {
                throw new InvalidOperationException("MathService not initialized. Call InitializeAsync first.");
            }

            if (string.IsNullOrWhiteSpace(latex))
            {
                return string.Empty;
            }

            // Execute on STA thread
            var tcs = new TaskCompletionSource<string>();
            _staSyncContext.Post(async _ =>
            {
                try
                {
                    var result = await ConvertLatexAsync(latex, displayMode).ConfigureAwait(false);
                    tcs.SetResult(result);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            }, null);

            return await tcs.Task.ConfigureAwait(false);
        }

        private async Task<string> ConvertLatexAsync(string latex, bool displayMode)
        {
            // Escape LaTeX for JavaScript string
            var escapedLatex = EscapeForJavaScript(latex);
            var displayArg = displayMode ? "true" : "false";

            // Call MathJax.tex2mml() via WebView2
            var script = $"texToMml('{escapedLatex}', {displayArg})";
            var result = await _webView.CoreWebView2.ExecuteScriptAsync(script).ConfigureAwait(false);

            // Result is a JSON-encoded string, remove quotes
            var mathml = UnescapeJsonString(result);

            // Add mml: namespace prefix for OneNote compatibility
            mathml = AddMmlNamespacePrefix(mathml);

            return mathml;
        }

        public string WrapMathMLForOneNote(string mathml)
        {
            if (string.IsNullOrWhiteSpace(mathml))
            {
                return string.Empty;
            }

            // Remove MathJax-specific data attributes
            mathml = RemoveMathJaxAttributes(mathml);

            // Compact MathML: remove newlines and extra whitespace
            mathml = CompactMathML(mathml);

            // Add fence="false" to brackets (verified fix for bracket/comma issues)
            mathml = AddFenceAttributeToBrackets(mathml);

            // Split multi-character identifiers into single chars (stability fix)
            mathml = SplitMultiCharIdentifiers(mathml);

            // Wrap with zero-width spaces and conditional comment
            const string zeroWidthSpan = "<span style='font-family:Arial'>\u200B</span>";
            return $"{zeroWidthSpan}<!--[if mathML]>{mathml}<![endif]-->{zeroWidthSpan}";
        }

        /// <summary>
        /// Removes MathJax-specific attributes that cause issues in OneNote.
        /// </summary>
        private string RemoveMathJaxAttributes(string mathml)
        {
            // Remove data-mjx-* attributes (MathJax internal, not standard MathML)
            var result = Regex.Replace(mathml, @"\s*data-mjx-[a-z]+=""[^""]*""", "", RegexOptions.IgnoreCase);

            // Remove stretchy="false" - without this, formulas like \int_{0}^{\pi} fail to render
            result = result.Replace(" stretchy=\"false\"", "");

            // Remove function application operator before parenthesis (verified fix for sin(x) etc.)
            result = Regex.Replace(result, @"<mml:mo>&#x2061;</mml:mo>\s*<mml:mo>\(", "<mml:mo>(");

            return result;
        }

        /// <summary>
        /// Compacts MathML by removing newlines and extra whitespace between tags.
        /// </summary>
        private string CompactMathML(string mathml)
        {
            // Remove newlines and multiple spaces
            var result = Regex.Replace(mathml, @"\s*\n\s*", "");
            result = Regex.Replace(result, @">\s+<", "><");
            return result.Trim();
        }

        /// <summary>
        /// Adds fence="false" to bracket operators to prevent OneNote from converting them to mfenced.
        /// Without this, OneNote converts (a, b, c) to mfenced and deletes our commas.
        /// </summary>
        private string AddFenceAttributeToBrackets(string mathml)
        {
            // Add fence="false" to parentheses
            var result = Regex.Replace(mathml, @"<mml:mo>\(</mml:mo>", "<mml:mo fence=\"false\">(</mml:mo>");
            result = Regex.Replace(result, @"<mml:mo>\)</mml:mo>", "<mml:mo fence=\"false\">)</mml:mo>");

            // Add fence="false" to square brackets
            result = Regex.Replace(result, @"<mml:mo>\[</mml:mo>", "<mml:mo fence=\"false\">[</mml:mo>");
            result = Regex.Replace(result, @"<mml:mo>\]</mml:mo>", "<mml:mo fence=\"false\">]</mml:mo>");

            // Add fence="false" to curly braces (if used as brackets, not grouping)
            result = Regex.Replace(result, @"<mml:mo>\{</mml:mo>", "<mml:mo fence=\"false\">{</mml:mo>");
            result = Regex.Replace(result, @"<mml:mo>\}</mml:mo>", "<mml:mo fence=\"false\">}</mml:mo>");

            return result;
        }

        /// <summary>
        /// Splits multi-character function names (like sin, cos, lim) into individual characters.
        /// OneNote re-parses MathML during page updates and splits them anyway,
        /// so we do it upfront to ensure formula stability across page edits.
        /// </summary>
        private string SplitMultiCharIdentifiers(string mathml)
        {
            // Match <mml:mi>abc</mml:mi> or <mml:mo>abc</mml:mo> where content is 2+ letters
            var result = Regex.Replace(mathml, @"<mml:(mi|mo)>([a-zA-Z]{2,})</mml:\1>", match =>
            {
                var content = match.Groups[2].Value;
                var chars = string.Concat(content.Select(c => $"<mml:mi>{c}</mml:mi>"));
                return $"<mml:mrow>{chars}</mml:mrow>";
            });

            return result;
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

            // Development: Walk up to find project root with Lib folder
            var dir = new DirectoryInfo(assemblyDir);
            while (dir != null)
            {
                var devPath = Path.Combine(dir.FullName, "Lib", "mathjax");
                if (Directory.Exists(devPath))
                {
                    return devPath;
                }
                dir = dir.Parent;
            }

            return null;
        }

        private string GetMathJaxLoaderHtml()
        {
            // Try to load from embedded resource first
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

            // Fallback: return inline HTML with MathJax CDN (requires internet)
            return @"<!DOCTYPE html>
<html>
<head>
    <script>
        MathJax = {
            startup: { typeset: false },
            tex: { packages: {'[+]': ['ams', 'newcommand', 'configmacros']} }
        };
    </script>
    <script src=""https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js""></script>
</head>
<body>
<script>
    var mathJaxReady = false;
    MathJax.startup.promise.then(() => { mathJaxReady = true; });

    function isMathJaxReady() {
        return mathJaxReady;
    }

    function texToMml(latex, display) {
        try {
            return MathJax.tex2mml(latex, { display: display });
        } catch (e) {
            return '<math xmlns=""http://www.w3.org/1998/Math/MathML""><merror><mtext>' + e.message + '</mtext></merror></math>';
        }
    }
</script>
</body>
</html>";
        }

        private async Task WaitForMathJaxReady()
        {
            var maxWaitMs = 30000; // 30 seconds timeout
            var intervalMs = 100;
            var elapsed = 0;

            while (elapsed < maxWaitMs)
            {
                var result = await _webView.CoreWebView2.ExecuteScriptAsync("isMathJaxReady()").ConfigureAwait(false);
                if (result == "true")
                {
                    return;
                }

                await Task.Delay(intervalMs).ConfigureAwait(false);
                elapsed += intervalMs;
            }

            throw new TimeoutException("MathJax failed to initialize within timeout period.");
        }

        private string AddMmlNamespacePrefix(string mathml)
        {
            // Add mml: prefix to all MathML elements
            var result = MathMLElementRegex.Replace(mathml, "<$1mml:$2$3");

            // Update xmlns to use mml: prefix
            result = result.Replace(
                "xmlns=\"http://www.w3.org/1998/Math/MathML\"",
                "xmlns:mml=\"http://www.w3.org/1998/Math/MathML\"");

            return result;
        }

        private string EscapeForJavaScript(string input)
        {
            return input
                .Replace("\\", "\\\\")
                .Replace("'", "\\'")
                .Replace("\n", "\\n")
                .Replace("\r", "\\r")
                .Replace("\t", "\\t");
        }

        private string UnescapeJsonString(string json)
        {
            if (string.IsNullOrEmpty(json) || json == "null")
            {
                return string.Empty;
            }

            // Remove surrounding quotes
            if (json.StartsWith("\"") && json.EndsWith("\""))
            {
                json = json.Substring(1, json.Length - 2);
            }

            // Unescape JSON string
            return Regex.Unescape(json);
        }

        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;

            // Stop the STA thread message loop
            if (_staSyncContext != null)
            {
                _staSyncContext.Post(_ =>
                {
                    _webView?.Dispose();
                    Application.ExitThread();
                }, null);
            }

            _initLock?.Dispose();
        }
    }
}
