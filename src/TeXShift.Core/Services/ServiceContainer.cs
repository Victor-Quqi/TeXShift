using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Markdig;
using TeXShift.Core.Abstractions;
using TeXShift.Core.Configuration;
using TeXShift.Core.Logging;
using TeXShift.Core.Markdown;
using TeXShift.Core.Math;
using TeXShift.Core.Mermaid;
using TeXShift.Core.OneNote;
using TeXShift.Core.OneNoteToMarkdown;
using TeXShift.Core.OneNoteToMarkdown.Abstractions;
using OneNoteApp = Microsoft.Office.Interop.OneNote;

namespace TeXShift.Core.Services
{
    /// <summary>
    /// Simple dependency injection container for managing service lifetimes.
    /// Implements the Service Locator pattern for COM Add-in scenarios.
    /// </summary>
    public class ServiceContainer : IDisposable
    {
        /// <summary>Reads the dialog owner on the shared OneNote STA; no COM objects leave it.</summary>
        public Task<IntPtr> GetOneNoteWindowHandleAsync(OneNoteApp.Application application)
        {
            return GetOneNoteDispatcher(application).InvokeAsync(app =>
            {
                OneNoteApp.Windows windows = null;
                OneNoteApp.Window window = null;
                try
                {
                    windows = app.Windows;
                    window = windows?.CurrentWindow;
                    return window == null ? IntPtr.Zero : new IntPtr(unchecked((long)window.WindowHandle));
                }
                finally
                {
                    if (window != null) Marshal.ReleaseComObject(window);
                    if (windows != null) Marshal.ReleaseComObject(windows);
                }
            });
        }

        // Singleton instances (shared for entire add-in lifetime)
        private readonly Lazy<OneNoteStyleConfig> _styleConfig;
        private readonly Lazy<MarkdownPipeline> _markdownPipeline;
        private readonly Lazy<IMathEngine> _mathEngine;
        private readonly Lazy<IMermaidService> _mermaidService;
        private readonly object _oneNoteDispatcherLock = new object();
        private OneNoteComDispatcher _oneNoteDispatcher;
        private bool _disposed;

        /// <summary>
        /// Gets or sets the Mermaid render options.
        /// Set this before conversion to apply theme, maxWidth, maxHeight settings.
        /// </summary>
        public MermaidRenderOptions MermaidOptions { get; set; }
        public MathSettings MathSettings { get; set; } = new MathSettings();

        public void ApplySettings(AppSettings settings)
        {
            var code = settings.CodeBlock ?? new CodeBlockStyleSettings();
            StyleConfig.SetCodeBlockStyle(code.BackgroundColor, code.TextColor, code.FontFamily,
                code.FontSize, code.SpaceBetween, code.EnableSyntaxHighlight);
            var inline = settings.InlineCode ?? new InlineCodeStyleSettings();
            StyleConfig.SetInlineCodeStyle(inline.FontFamily, inline.BackgroundColor);
            StyleConfig.SetQuoteBlockStyle((settings.QuoteBlock ?? new QuoteBlockStyleSettings()).BackgroundColor);
            var headings = settings.Headings ?? new HeadingStyleSettings();
            for (int i = 1; i <= 6; i++) StyleConfig.SetHeadingFont(i, headings.GetFontSize(i));
            MathSettings = settings.Math ?? new MathSettings();
            var mermaid = settings.Mermaid ?? new MermaidSettings();
            MermaidOptions = new MermaidRenderOptions {
                Theme = mermaid.Theme ?? "default", MaxWidth = mermaid.MaxWidth > 0 ? mermaid.MaxWidth : 1920,
                MaxHeight = mermaid.MaxHeight > 0 ? mermaid.MaxHeight : 1080
            };
            StyleConfig.SetHorizontalRuleStyle((settings.HorizontalRule?.UseImage ?? true)
                ? OneNoteStyleConfig.HorizontalRuleMode.Image : OneNoteStyleConfig.HorizontalRuleMode.Character,
                "#888888", 90, '─', 2325);
            StyleConfig.SetReverseConversionOptions((settings.ReverseConversion ?? new ReverseConversionSettings()).TryRecognizeNonTeXShiftFormats);
        }

        public ServiceContainer()
        {
            // Initialize singletons lazily
            _styleConfig = new Lazy<OneNoteStyleConfig>(() => new OneNoteStyleConfig());

            // NOTE: We explicitly list extensions instead of UseAdvancedExtensions() because
            // UseGenericAttributes() (included in UseAdvancedExtensions) interprets {...} as
            // HTML attributes, conflicting with LaTeX subscript/superscript syntax like _{total}.
            // UseMathematics() must be added FIRST to take parsing precedence.
            _markdownPipeline = new Lazy<MarkdownPipeline>(() =>
                new MarkdownPipelineBuilder()
                    .UseMathematics()        // Enable $...$ and $$...$$ math syntax - MUST BE FIRST
                    .UseAbbreviations()
                    .UseAutoIdentifiers()
                    .UseCitations()
                    .UseCustomContainers()
                    .UseDefinitionLists()
                    .UseEmphasisExtras()     // Strikethrough, subscript, superscript, etc.
                    .UseFigures()
                    .UseFooters()
                    .UseFootnotes()
                    .UseGridTables()
                    .UseMediaLinks()
                    .UsePipeTables()
                    // .UseGenericAttributes() - EXCLUDED: conflicts with LaTeX {...} syntax
                    .UseAutoLinks()
                    .UseTaskLists()
                    .UseDiagrams()
                    .UseListExtras()         // Add-on for more flexible list parsing
                    .Build()
            );

            _mathEngine = new Lazy<IMathEngine>(() => new MathJaxEngine());

            _mermaidService = new Lazy<IMermaidService>(() =>
            {
                var service = new MermaidService();
                // Note: InitializeAsync() should be called before first use
                // This is handled in MermaidBlockHandler
                return service;
            });
        }

        /// <summary>
        /// Gets the shared OneNoteStyleConfig instance.
        /// </summary>
        public OneNoteStyleConfig StyleConfig => _styleConfig.Value;

        /// <summary>
        /// Gets the shared MarkdownPipeline instance.
        /// Thread-safe and expensive to create, so we cache it.
        /// </summary>
        public MarkdownPipeline MarkdownPipeline => _markdownPipeline.Value;

        /// <summary>
        /// Gets the shared MathEngine instance for LaTeX to MathML conversion.
        /// Thread-safe and requires WebView2 initialization, so we cache it.
        /// </summary>
        public IMathEngine MathEngine => _mathEngine.Value;

        /// <summary>
        /// Gets the shared MermaidService instance for Mermaid diagram rendering.
        /// Thread-safe and requires WebView2 initialization, so we cache it.
        /// </summary>
        public IMermaidService MermaidService => _mermaidService.Value;

        /// <summary>
        /// Creates a new IContentReader instance.
        /// Transient lifetime: new instance per call.
        /// </summary>
        public IContentReader CreateContentReader(OneNoteApp.Application oneNoteApp)
        {
            if (oneNoteApp == null)
                throw new ArgumentNullException(nameof(oneNoteApp));

            return new OneNotePageReader(GetOneNoteDispatcher(oneNoteApp));
        }

        internal IContentReader CreateContentReader(OneNoteComDispatcher dispatcher)
        {
            return new OneNotePageReader(dispatcher ?? throw new ArgumentNullException(nameof(dispatcher)));
        }

        /// <summary>
        /// Creates a new IMarkdownConverter instance.
        /// Transient lifetime: new instance per call.
        /// Uses singleton StyleConfig, MarkdownPipeline, and MathEngine for efficiency.
        /// </summary>
        public IMarkdownConverter CreateMarkdownConverter(double? sourceOutlineWidth = null)
        {
            return new MarkdownToOneNoteConverter(StyleConfig, MarkdownPipeline, MathEngine, MermaidService, MermaidOptions, sourceOutlineWidth, MathSettings);
        }

        /// <summary>
        /// Creates a new OneNote -> Markdown converter.
        /// Transient lifetime: new instance per call.
        /// </summary>
        public IOneNoteToMarkdownConverter CreateOneNoteToMarkdownConverter()
        {
            return new OneNoteToMarkdownConverter(StyleConfig);
        }

        internal IContentWriter CreateContentWriter(OneNoteComDispatcher dispatcher, IDebugLogger logger = null)
        {
            return new OneNotePageWriter(dispatcher ?? throw new ArgumentNullException(nameof(dispatcher)), logger);
        }

        internal OneNotePublisher CreateOneNotePublisher(OneNoteComDispatcher dispatcher)
        {
            return new OneNotePublisher(dispatcher ?? throw new ArgumentNullException(nameof(dispatcher)));
        }

        /// <summary>
        /// Creates a new paste conversion orchestrator.
        /// Transient lifetime: new instance per call.
        /// </summary>
        public PasteConversionOrchestrator CreatePasteConversionOrchestrator(OneNoteApp.Application oneNoteApp)
        {
            return new PasteConversionOrchestrator(this, GetOneNoteDispatcher(oneNoteApp));
        }

        public PasteConversionOrchestrator CreatePasteConversionOrchestrator(OneNoteComDispatcher dispatcher)
        {
            return new PasteConversionOrchestrator(this, dispatcher ?? throw new ArgumentNullException(nameof(dispatcher)));
        }

        public ConversionOrchestrator CreateConversionOrchestrator(OneNoteApp.Application oneNoteApp)
        {
            if (oneNoteApp == null)
                throw new ArgumentNullException(nameof(oneNoteApp));

            return new ConversionOrchestrator(this, GetOneNoteDispatcher(oneNoteApp));
        }

        /// <summary>Uses a caller-owned dispatcher. The caller releases it after the services.</summary>
        public ConversionOrchestrator CreateConversionOrchestrator(OneNoteComDispatcher dispatcher)
        {
            if (dispatcher == null)
                throw new ArgumentNullException(nameof(dispatcher));

            return new ConversionOrchestrator(this, dispatcher);
        }

        /// <summary>
        /// Creates a new ReverseConversionOrchestrator instance.
        /// Transient lifetime: new instance per call.
        /// </summary>
        public ReverseConversionOrchestrator CreateReverseConversionOrchestrator(OneNoteApp.Application oneNoteApp)
        {
            if (oneNoteApp == null)
                throw new ArgumentNullException(nameof(oneNoteApp));

            return new ReverseConversionOrchestrator(this, GetOneNoteDispatcher(oneNoteApp));
        }

        /// <summary>Uses a caller-owned dispatcher. The caller releases it after the services.</summary>
        public ReverseConversionOrchestrator CreateReverseConversionOrchestrator(OneNoteComDispatcher dispatcher)
        {
            if (dispatcher == null)
                throw new ArgumentNullException(nameof(dispatcher));

            return new ReverseConversionOrchestrator(this, dispatcher);
        }

        /// <summary>
        /// Creates a new IDebugLogger instance.
        /// Transient lifetime: new instance per call.
        /// </summary>
        /// <param name="customOutputPath">Custom output directory path. If null or empty, uses default location.</param>
        public IDebugLogger CreateDebugLogger(string customOutputPath = null)
        {
            return new DebugLogger(customOutputPath);
        }

        /// <summary>
        /// Disposes managed resources, including the MathEngine's WebView2 instance.
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try
            {
                if (_mathEngine.IsValueCreated && _mathEngine.Value is IDisposable disposable)
                    disposable.Dispose();
            }
            finally
            {
                try
                {
                    if (_mermaidService.IsValueCreated && _mermaidService.Value is IDisposable mermaidDisposable)
                        mermaidDisposable.Dispose();
                }
                finally { _oneNoteDispatcher?.Dispose(); }
            }
        }

        private OneNoteComDispatcher GetOneNoteDispatcher(OneNoteApp.Application oneNoteApp)
        {
            if (oneNoteApp == null) throw new ArgumentNullException(nameof(oneNoteApp));
            lock (_oneNoteDispatcherLock)
            {
                if (_oneNoteDispatcher == null)
                    _oneNoteDispatcher = new OneNoteComDispatcher(oneNoteApp);
                return _oneNoteDispatcher;
            }
        }
    }
}
