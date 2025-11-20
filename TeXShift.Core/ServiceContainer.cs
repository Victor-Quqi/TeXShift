using System;
using Markdig;
using TeXShift.Core.Logging;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace TeXShift.Core
{
    /// <summary>
    /// Simple dependency injection container for managing service lifetimes.
    /// Implements the Service Locator pattern for COM Add-in scenarios.
    /// </summary>
    public class ServiceContainer
    {
        // Singleton instances (shared for entire add-in lifetime)
        private readonly Lazy<OneNoteStyleConfig> _styleConfig;
        private readonly Lazy<MarkdownPipeline> _markdownPipeline;

        public ServiceContainer()
        {
            // Initialize singletons lazily
            _styleConfig = new Lazy<OneNoteStyleConfig>(() => new OneNoteStyleConfig());

            _markdownPipeline = new Lazy<MarkdownPipeline>(() =>
                new MarkdownPipelineBuilder()
                    .UseAdvancedExtensions() // Includes most common extensions
                    .UseListExtras()         // Add-on for more flexible list parsing (e.g., different indentations)
                    .Build()
            );
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
        /// Creates a new IContentReader instance.
        /// Transient lifetime: new instance per call.
        /// </summary>
        public IContentReader CreateContentReader(OneNote.Application oneNoteApp)
        {
            if (oneNoteApp == null)
                throw new ArgumentNullException(nameof(oneNoteApp));

            return new ContentReader(oneNoteApp);
        }

        /// <summary>
        /// Creates a new IMarkdownConverter instance.
        /// Transient lifetime: new instance per call.
        /// Uses singleton StyleConfig and MarkdownPipeline for efficiency.
        /// </summary>
        public IMarkdownConverter CreateMarkdownConverter(double? sourceOutlineWidth = null)
        {
            return new MarkdownConverter(StyleConfig, MarkdownPipeline, sourceOutlineWidth);
        }

        /// <summary>
        /// Creates a new IContentWriter instance.
        /// Transient lifetime: new instance per call.
        /// </summary>
        public IContentWriter CreateContentWriter(OneNote.Application oneNoteApp)
        {
            if (oneNoteApp == null)
                throw new ArgumentNullException(nameof(oneNoteApp));

            return new ContentWriter(oneNoteApp);
        }

        /// <summary>
        /// Creates a new IDebugLogger instance.
        /// Transient lifetime: new instance per call.
        /// </summary>
        public IDebugLogger CreateDebugLogger()
        {
            return new DebugLogger();
        }
    }
}
