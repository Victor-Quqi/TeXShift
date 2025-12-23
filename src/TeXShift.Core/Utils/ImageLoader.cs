using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

namespace TeXShift.Core.Utils
{
    /// <summary>
    /// Utility class for loading images from local files or URLs and converting to base64.
    /// </summary>
    public static class ImageLoader
    {
        private static readonly HttpClient _httpClient;
        private const int MaxFileSizeBytes = 10 * 1024 * 1024; // 10MB
        private const int TimeoutSeconds = 30;

        static ImageLoader()
        {
            // Enable TLS 1.2 for .NET Framework 4.8
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;

            var handler = new HttpClientHandler
            {
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };

            _httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(TimeoutSeconds)
            };

            // Add User-Agent to avoid being blocked by servers
            _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) TeXShift/1.0");
        }

        /// <summary>
        /// Result of an image load operation.
        /// </summary>
        public class ImageLoadResult
        {
            public bool Success { get; set; }
            public string Base64Data { get; set; }
            public string Format { get; set; }
            public string ErrorMessage { get; set; }
        }

        /// <summary>
        /// Loads an image from a local path or URL and returns base64-encoded data.
        /// </summary>
        /// <param name="source">Local file path or URL</param>
        /// <returns>ImageLoadResult containing success status and data</returns>
        public static async Task<ImageLoadResult> LoadImageAsync(string source)
        {
            if (string.IsNullOrWhiteSpace(source))
            {
                return new ImageLoadResult { Success = false, ErrorMessage = "Empty source" };
            }

            // Determine if it's a URL or local path
            if (Uri.TryCreate(source, UriKind.Absolute, out var uri))
            {
                if (uri.Scheme == "http" || uri.Scheme == "https")
                {
                    return await LoadFromUrlAsync(uri);
                }
                else if (uri.Scheme == "file" || uri.IsFile)
                {
                    return LoadFromFile(uri.LocalPath);
                }
            }

            // Treat as local path
            return LoadFromFile(source);
        }

        /// <summary>
        /// Synchronous wrapper for LoadImageAsync.
        /// </summary>
        public static ImageLoadResult LoadImage(string source)
        {
            return LoadImageAsync(source).GetAwaiter().GetResult();
        }

        private static ImageLoadResult LoadFromFile(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    return new ImageLoadResult { Success = false, ErrorMessage = "File not found" };
                }

                var fileInfo = new FileInfo(filePath);
                if (fileInfo.Length > MaxFileSizeBytes)
                {
                    return new ImageLoadResult { Success = false, ErrorMessage = "File too large" };
                }

                var format = GetImageFormat(filePath);
                if (format == null)
                {
                    return new ImageLoadResult { Success = false, ErrorMessage = "Unsupported format" };
                }

                var bytes = File.ReadAllBytes(filePath);
                var base64 = Convert.ToBase64String(bytes);

                return new ImageLoadResult
                {
                    Success = true,
                    Base64Data = base64,
                    Format = format
                };
            }
            catch (Exception ex)
            {
                return new ImageLoadResult { Success = false, ErrorMessage = ex.Message };
            }
        }

        private static async Task<ImageLoadResult> LoadFromUrlAsync(Uri uri)
        {
            try
            {
                using (var response = await _httpClient.GetAsync(uri, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(false))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        return new ImageLoadResult { Success = false, ErrorMessage = $"HTTP {response.StatusCode}" };
                    }

                    // Check content length if available
                    if (response.Content.Headers.ContentLength > MaxFileSizeBytes)
                    {
                        return new ImageLoadResult { Success = false, ErrorMessage = "File too large" };
                    }

                    var bytes = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);

                    if (bytes.Length > MaxFileSizeBytes)
                    {
                        return new ImageLoadResult { Success = false, ErrorMessage = "File too large" };
                    }

                    var format = GetImageFormat(uri.AbsolutePath);
                    if (format == null)
                    {
                        // Try to detect from content
                        format = DetectFormatFromBytes(bytes);
                    }
                    if (format == null)
                    {
                        return new ImageLoadResult { Success = false, ErrorMessage = "Unsupported format" };
                    }

                    var base64 = Convert.ToBase64String(bytes);

                    return new ImageLoadResult
                    {
                        Success = true,
                        Base64Data = base64,
                        Format = format
                    };
                }
            }
            catch (TaskCanceledException)
            {
                return new ImageLoadResult { Success = false, ErrorMessage = "Request timeout" };
            }
            catch (Exception ex)
            {
                return new ImageLoadResult { Success = false, ErrorMessage = ex.Message };
            }
        }

        /// <summary>
        /// Gets the image format from file extension.
        /// </summary>
        private static string GetImageFormat(string path)
        {
            var ext = Path.GetExtension(path)?.ToLowerInvariant();
            switch (ext)
            {
                case ".png": return "png";
                case ".jpg":
                case ".jpeg": return "jpg";
                case ".gif": return "gif";
                case ".bmp": return "bmp";
                case ".webp": return "webp";
                case ".avif": return "avif";
                default: return null;
            }
        }

        /// <summary>
        /// Detects image format from file header bytes (magic numbers).
        /// </summary>
        private static string DetectFormatFromBytes(byte[] bytes)
        {
            if (bytes == null || bytes.Length < 8) return null;

            // PNG: 89 50 4E 47 0D 0A 1A 0A
            if (bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
                return "png";

            // JPEG: FF D8 FF
            if (bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
                return "jpg";

            // GIF: 47 49 46 38
            if (bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x38)
                return "gif";

            // BMP: 42 4D
            if (bytes[0] == 0x42 && bytes[1] == 0x4D)
                return "bmp";

            // WebP: 52 49 46 46 ... 57 45 42 50
            if (bytes[0] == 0x52 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x46 &&
                bytes.Length > 11 && bytes[8] == 0x57 && bytes[9] == 0x45 && bytes[10] == 0x42 && bytes[11] == 0x50)
                return "webp";

            // AVIF: ISOBMFF container with "ftyp" at offset 4 and "avif"/"avis"/"mif1" brand
            if (bytes.Length > 11 && bytes[4] == 0x66 && bytes[5] == 0x74 && bytes[6] == 0x79 && bytes[7] == 0x70)
            {
                // Check for "avif" brand at offset 8
                if (bytes[8] == 0x61 && bytes[9] == 0x76 && bytes[10] == 0x69 && bytes[11] == 0x66)
                    return "avif";
                // Check for "avis" brand at offset 8
                if (bytes[8] == 0x61 && bytes[9] == 0x76 && bytes[10] == 0x69 && bytes[11] == 0x73)
                    return "avif";
                // Check for "mif1" brand at offset 8 (HEIF/AVIF)
                if (bytes[8] == 0x6D && bytes[9] == 0x69 && bytes[10] == 0x66 && bytes[11] == 0x31)
                    return "avif";
            }

            return null;
        }
    }
}
