using System;
using System.IO;
using System.Threading.Tasks;
using TeXShift.Core.Logging;
using TeXShift.Core.OneNote;
using OneNoteInterop = Microsoft.Office.Interop.OneNote;

namespace TeXShift.Core.Services
{
    /// <summary>
    /// Options for executing the conversion pipeline.
    /// </summary>
    public class ConversionOptions
    {
        /// <summary>
        /// Whether to write debug artifacts to disk.
        /// </summary>
        public bool WriteDebugFiles { get; set; }

        /// <summary>
        /// Whether to export the current page as a PDF after conversion.
        /// </summary>
        public bool ExportPdf { get; set; }

        /// <summary>
        /// Output directory for debug files and optional PDF export.
        /// </summary>
        public string OutputDirectory { get; set; }
    }

    /// <summary>
    /// Represents the result of a conversion pipeline run.
    /// </summary>
    public class ConversionResult
    {
        /// <summary>
        /// Whether the pipeline completed successfully.
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// The read result returned by the content reader.
        /// </summary>
        public ReadResult ReadResult { get; set; }

        /// <summary>
        /// The folder where debug artifacts were written.
        /// </summary>
        public string DebugOutputFolder { get; set; }

        /// <summary>
        /// The full path of the exported PDF, if requested.
        /// </summary>
        public string PdfPath { get; set; }

        /// <summary>
        /// The exception captured during execution, if any.
        /// </summary>
        public Exception Error { get; set; }
    }

    /// <summary>
    /// Orchestrates the Read -> Convert -> Write pipeline.
    /// </summary>
    public class ConversionOrchestrator
    {
        private readonly ServiceContainer _serviceContainer;
        private readonly OneNoteInterop.Application _oneNoteApp;

        public ConversionOrchestrator(ServiceContainer serviceContainer, OneNoteInterop.Application oneNoteApp)
        {
            if (serviceContainer == null)
                throw new ArgumentNullException(nameof(serviceContainer));
            if (oneNoteApp == null)
                throw new ArgumentNullException(nameof(oneNoteApp));

            _serviceContainer = serviceContainer;
            _oneNoteApp = oneNoteApp;
        }

        /// <summary>
        /// Executes the conversion pipeline with optional debug logging and PDF export.
        /// </summary>
        public async Task<ConversionResult> ExecuteAsync(ConversionOptions options)
        {
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            IDebugLogger logger = null;
            var result = new ConversionResult();

            try
            {
                if (options.WriteDebugFiles)
                {
                    logger = _serviceContainer.CreateDebugLogger(options.OutputDirectory);
                    logger.StartSession();
                    result.DebugOutputFolder = logger.DebugSessionFolder;
                }

                var reader = _serviceContainer.CreateContentReader(_oneNoteApp);
                var readResult = await reader.ExtractContentAsync().ConfigureAwait(false);
                result.ReadResult = readResult;

                if (!readResult.IsSuccess)
                {
                    var error = new InvalidOperationException(readResult.ErrorMessage ?? "Content extraction failed.");
                    result.Error = error;
                    if (options.WriteDebugFiles)
                    {
                        await logger.LogErrorAsync(error).ConfigureAwait(false);
                    }
                    return result;
                }

                if (options.WriteDebugFiles)
                {
                    await logger.LogInputMarkdownAsync(readResult.ExtractedText).ConfigureAwait(false);
                    await logger.LogOriginalXmlAsync(readResult.OriginalXmlNode).ConfigureAwait(false);
                }

                var converter = _serviceContainer.CreateMarkdownConverter(readResult.SourceOutlineWidth);
                var oneNoteXml = await converter.ConvertToOneNoteXmlAsync(readResult.ExtractedText).ConfigureAwait(false);

                if (options.WriteDebugFiles)
                {
                    await logger.LogConvertedXmlAsync(oneNoteXml).ConfigureAwait(false);
                }

                var writer = _serviceContainer.CreateContentWriter(_oneNoteApp);
                await writer.ReplaceContentAsync(readResult, oneNoteXml).ConfigureAwait(false);

                if (options.WriteDebugFiles)
                {
                    string finalPageXml = await Task.Run(() =>
                    {
                        _oneNoteApp.GetPageContent(readResult.PageId, out string xml, OneNoteInterop.PageInfo.piAll);
                        return xml;
                    }).ConfigureAwait(false);

                    await logger.LogFinalPageXmlAsync(finalPageXml).ConfigureAwait(false);
                }

                if (options.ExportPdf)
                {
                    string pdfPath = ResolvePdfOutputPath(options.OutputDirectory, result.DebugOutputFolder);
                    var publisher = _serviceContainer.CreateOneNotePublisher(_oneNoteApp);
                    bool exported = await publisher.ExportToPdfAsync(readResult.PageId, pdfPath).ConfigureAwait(false);
                    if (!exported)
                    {
                        var error = new IOException("PDF export failed.");
                        result.Error = error;
                        if (options.WriteDebugFiles && logger != null)
                        {
                            await logger.LogErrorAsync(error).ConfigureAwait(false);
                        }
                        return result;
                    }
                    result.PdfPath = pdfPath;
                }

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Error = ex;
                if (options.WriteDebugFiles && logger != null)
                {
                    await logger.LogErrorAsync(ex).ConfigureAwait(false);
                }
                return result;
            }
        }

        private string ResolvePdfOutputPath(string outputDirectory, string debugOutputFolder)
        {
            var baseDirectory = outputDirectory;
            if (string.IsNullOrWhiteSpace(baseDirectory))
            {
                baseDirectory = debugOutputFolder;
            }

            if (string.IsNullOrWhiteSpace(baseDirectory))
            {
                throw new InvalidOperationException("OutputDirectory is required for PDF export.");
            }

            string fileName = $"TeXShift_Export_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
            return Path.Combine(baseDirectory, fileName);
        }
    }
}
