using System;
using System.IO;
using System.Threading.Tasks;
using OneNoteInterop = Microsoft.Office.Interop.OneNote;

namespace TeXShift.Core.OneNote
{
    /// <summary>
    /// Publishes OneNote pages to external formats.
    /// </summary>
    public class OneNotePublisher
    {
        private readonly OneNoteInterop.Application _oneNoteApp;

        public OneNotePublisher(OneNoteInterop.Application oneNoteApp)
        {
            if (oneNoteApp == null)
                throw new ArgumentNullException(nameof(oneNoteApp));

            _oneNoteApp = oneNoteApp;
        }

        /// <summary>
        /// Asynchronously exports a OneNote page to a PDF file.
        /// </summary>
        /// <param name="pageId">The OneNote page ID to export.</param>
        /// <param name="outputPath">The full path to the output PDF file.</param>
        /// <returns>True if the PDF file was created successfully.</returns>
        public async Task<bool> ExportToPdfAsync(string pageId, string outputPath)
        {
            if (string.IsNullOrWhiteSpace(pageId))
                throw new ArgumentException("Page ID is required.", nameof(pageId));
            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("Output path is required.", nameof(outputPath));

            return await Task.Run(() =>
            {
                string directory = Path.GetDirectoryName(outputPath);
                if (string.IsNullOrWhiteSpace(directory))
                {
                    throw new ArgumentException("Output path must include a directory.", nameof(outputPath));
                }

                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }

                _oneNoteApp.Publish(pageId, outputPath, OneNoteInterop.PublishFormat.pfPDF, string.Empty);
                return File.Exists(outputPath);
            }).ConfigureAwait(false);
        }
    }
}
