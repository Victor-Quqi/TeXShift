using System;
using System.Threading.Tasks;

namespace TeXShift.Core.Math
{
    /// <summary>
    /// Service interface for converting LaTeX math to MathML format.
    /// Uses WebView2 + MathJax for conversion.
    /// </summary>
    public interface IMathService : IDisposable
    {
        /// <summary>
        /// Gets whether the service has been initialized.
        /// </summary>
        bool IsInitialized { get; }

        /// <summary>
        /// Initializes the WebView2 environment and loads MathJax.
        /// Must be called before any conversion operations.
        /// </summary>
        Task InitializeAsync();

        /// <summary>
        /// Converts LaTeX math expression to MathML string.
        /// </summary>
        /// <param name="latex">The LaTeX expression (without $ delimiters)</param>
        /// <param name="displayMode">True for block-level (display) math, false for inline math</param>
        /// <returns>MathML string formatted for OneNote (with mml: namespace prefix)</returns>
        Task<string> LatexToMathMLAsync(string latex, bool displayMode);

        /// <summary>
        /// Wraps MathML in OneNote conditional comment format.
        /// </summary>
        /// <param name="mathml">The MathML string</param>
        /// <returns>MathML wrapped in &lt;!--[if mathML]&gt;...&lt;![endif]--&gt;</returns>
        string WrapMathMLForOneNote(string mathml);
    }
}
