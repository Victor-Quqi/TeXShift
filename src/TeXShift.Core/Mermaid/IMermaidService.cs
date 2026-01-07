using System;
using System.Threading.Tasks;

namespace TeXShift.Core.Mermaid
{
    /// <summary>
    /// Service interface for converting Mermaid diagram code to an image.
    /// Uses WebView2 + Mermaid.js for rendering.
    /// </summary>
    public interface IMermaidService : IDisposable
    {
        /// <summary>
        /// Gets whether the service has been initialized.
        /// </summary>
        bool IsInitialized { get; }

        /// <summary>
        /// Initializes the WebView2 environment and loads Mermaid.js.
        /// Must be called before any rendering operations.
        /// </summary>
        Task InitializeAsync();

        /// <summary>
        /// Renders Mermaid diagram code to a PNG image (base64).
        /// </summary>
        /// <param name="mermaidCode">Mermaid diagram code.</param>
        /// <param name="options">Render options, or null to use defaults.</param>
        Task<MermaidRenderResult> RenderToImageAsync(string mermaidCode, MermaidRenderOptions options = null);
    }

    public class MermaidRenderResult
    {
        public bool Success { get; set; }
        public string Base64PngData { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public string ErrorMessage { get; set; }
    }

    public class MermaidRenderOptions
    {
        public int MaxWidth { get; set; } = 1920;
        public int MaxHeight { get; set; } = 1080;
        public string Theme { get; set; } = "default";
    }
}

