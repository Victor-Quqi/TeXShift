using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using TeXShift.Core.Errors;
using TeXShift.Core.Localization;

namespace TeXShift.Core.Utils
{
    /// <summary>
    /// Utility class for loading images from local files or URLs and converting to base64.
    /// </summary>
    public static class ImageLoader
    {
        private static readonly object HttpClientGate = new object();
        private static HttpClient _httpClient;
        private static Exception _httpClientInitError;
        private const int MaxFileSizeBytes = 10 * 1024 * 1024; // 10MB
        private const int TimeoutSeconds = 30;

        private static HttpClient GetHttpClient()
        {
            if (_httpClient != null)
            {
                return _httpClient;
            }

            lock (HttpClientGate)
            {
                if (_httpClient != null)
                {
                    return _httpClient;
                }

                _httpClient = CreateHttpClientSafe(out _httpClientInitError);
                return _httpClient;
            }
        }

        private static HttpClient CreateHttpClientSafe(out Exception error)
        {
            error = null;

            try
            {
                // Enable TLS 1.2 for .NET Framework 4.8
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
            }
            catch (Exception ex)
            {
                // Don't fail type initialization if TLS configuration is blocked in a given environment.
                System.Diagnostics.Trace.WriteLine(ex);
            }

            try
            {
                var handler = new HttpClientHandler
                {
                    AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
                };

                var client = new HttpClient(handler)
                {
                    Timeout = TimeSpan.FromSeconds(TimeoutSeconds)
                };

                // Add User-Agent to avoid being blocked by servers
                client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) TeXShift/1.0");
                return client;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex);
                error = ex;
            }

            // Fallback: try a minimal HttpClient without a custom handler.
            try
            {
                var client = new HttpClient
                {
                    Timeout = TimeSpan.FromSeconds(TimeoutSeconds)
                };

                client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) TeXShift/1.0");
                return client;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex);
                error = ex;
                return null;
            }
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
                return new ImageLoadResult { Success = false, ErrorMessage = Resources.GetString("Error_Image_EmptySource") };
            }

            // Support data: URLs (used by reverse conversion to keep images self-contained).
            if (source.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            {
                return LoadFromDataUrl(source);
            }

            // Determine if it's a URL or local path
            if (Uri.TryCreate(source, UriKind.Absolute, out var uri))
            {
                if (uri.Scheme == "http" || uri.Scheme == "https")
                {
                    return await LoadFromUrlAsync(uri).ConfigureAwait(false);
                }
                else if (uri.Scheme == "file" || uri.IsFile)
                {
                    return LoadFromFile(uri.LocalPath);
                }
            }

            // Treat as local path
            return LoadFromFile(source);
        }

        private static ImageLoadResult LoadFromDataUrl(string dataUrl)
        {
            try
            {
                // data:[<mediatype>][;base64],<data>
                int comma = dataUrl.IndexOf(',');
                if (comma < 0)
                {
                    return new ImageLoadResult { Success = false, ErrorMessage = "Invalid data URL (missing comma)" };
                }

                string meta = dataUrl.Substring("data:".Length, comma - "data:".Length);
                string payload = dataUrl.Substring(comma + 1);

                bool isBase64 = meta.IndexOf(";base64", StringComparison.OrdinalIgnoreCase) >= 0;
                string mediaType = meta.Split(new[] { ';' }, 2)[0];

                if (!isBase64)
                {
                    return new ImageLoadResult { Success = false, ErrorMessage = "Unsupported data URL (not base64)" };
                }

                if (string.IsNullOrWhiteSpace(mediaType) || !mediaType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
                {
                    return new ImageLoadResult { Success = false, ErrorMessage = "Unsupported data URL (not an image)" };
                }

                // Some writers percent-encode the payload; decode only when needed.
                var base64 = (payload.IndexOf('%') >= 0) ? Uri.UnescapeDataString(payload) : payload;
                base64 = RemoveWhitespace(base64);

                // Approximate size check before decoding.
                var approxBytes = ApproximateDecodedSize(base64);
                if (approxBytes > MaxFileSizeBytes)
                {
                    return new ImageLoadResult { Success = false, ErrorMessage = Resources.GetString("Error_Image_FileTooLarge") };
                }

                // Validate base64 and re-check actual size.
                var bytes = Convert.FromBase64String(base64);
                if (bytes.Length > MaxFileSizeBytes)
                {
                    return new ImageLoadResult { Success = false, ErrorMessage = Resources.GetString("Error_Image_FileTooLarge") };
                }

                var format = GetImageFormatFromMime(mediaType) ?? DetectFormatFromBytes(bytes);
                if (format == null)
                {
                    return new ImageLoadResult { Success = false, ErrorMessage = Resources.GetString("Error_Image_UnsupportedFormat") };
                }

                return new ImageLoadResult
                {
                    Success = true,
                    Base64Data = base64,
                    Format = format
                };
            }
            catch (FormatException)
            {
                return new ImageLoadResult { Success = false, ErrorMessage = "Invalid data URL (bad base64)" };
            }
            catch (Exception ex)
            {
                throw new ImageLoadException(
                    Resources.GetString("Error_Image_LoadFailed"),
                    $"Failed to load image from data URL. {ex.GetType().Name}: {ex.Message}",
                    ex);
            }
        }

        private static string GetImageFormatFromMime(string mime)
        {
            if (string.IsNullOrWhiteSpace(mime))
            {
                return null;
            }

            switch (mime.Trim().ToLowerInvariant())
            {
                case "image/png": return "png";
                case "image/jpeg":
                case "image/jpg": return "jpg";
                case "image/gif": return "gif";
                case "image/bmp": return "bmp";
                case "image/webp": return "webp";
                case "image/avif": return "avif";
                default: return null;
            }
        }

        private static string RemoveWhitespace(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            var sb = new StringBuilder(text.Length);
            foreach (var c in text)
            {
                if (!char.IsWhiteSpace(c))
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        private static int ApproximateDecodedSize(string base64)
        {
            if (string.IsNullOrEmpty(base64))
            {
                return 0;
            }

            int len = base64.Length;
            int padding = 0;
            if (len >= 2 && base64[len - 1] == '=')
            {
                padding++;
                if (base64[len - 2] == '=')
                {
                    padding++;
                }
            }

            // Each 4 base64 chars represent up to 3 bytes.
            long bytes = ((long)len * 3) / 4 - padding;
            if (bytes < 0)
            {
                return 0;
            }
            if (bytes > int.MaxValue)
            {
                return int.MaxValue;
            }
            return (int)bytes;
        }

        /// <summary>
        /// Synchronous wrapper for LoadImageAsync.
        /// </summary>
        public static ImageLoadResult LoadImage(string source)
        {
            // Avoid deadlocks when called from a UI thread; prefer LoadImageAsync for non-blocking behavior.
            return LoadImageAsync(source).ConfigureAwait(false).GetAwaiter().GetResult();
        }

        private static ImageLoadResult LoadFromFile(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    return new ImageLoadResult { Success = false, ErrorMessage = Resources.GetString("Error_Image_FileNotFound") };
                }

                var fileInfo = new FileInfo(filePath);
                if (fileInfo.Length > MaxFileSizeBytes)
                {
                    return new ImageLoadResult { Success = false, ErrorMessage = Resources.GetString("Error_Image_FileTooLarge") };
                }

                var format = GetImageFormat(filePath);
                if (format == null)
                {
                    return new ImageLoadResult { Success = false, ErrorMessage = Resources.GetString("Error_Image_UnsupportedFormat") };
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
                throw new ImageLoadException(
                    Resources.GetString("Error_Image_LoadFailed"),
                    $"Failed to load image from file '{filePath}'. {ex.GetType().Name}: {ex.Message}",
                    ex);
            }
        }

        private static async Task<ImageLoadResult> LoadFromUrlAsync(Uri uri)
        {
            try
            {
                var httpClient = GetHttpClient();
                if (httpClient == null)
                {
                    System.Diagnostics.Trace.WriteLine(_httpClientInitError);
                    return new ImageLoadResult { Success = false, ErrorMessage = Resources.GetString("Error_Image_LoadFailed") };
                }

                using (var response = await httpClient.GetAsync(uri, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(false))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        return new ImageLoadResult { Success = false, ErrorMessage = $"HTTP {response.StatusCode}" };
                    }

                    // Check content length if available
                    if (response.Content.Headers.ContentLength > MaxFileSizeBytes)
                    {
                        return new ImageLoadResult { Success = false, ErrorMessage = Resources.GetString("Error_Image_FileTooLarge") };
                    }

                    var bytes = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);

                    if (bytes.Length > MaxFileSizeBytes)
                    {
                        return new ImageLoadResult { Success = false, ErrorMessage = Resources.GetString("Error_Image_FileTooLarge") };
                    }

                    var format = GetImageFormat(uri.AbsolutePath);
                    if (format == null)
                    {
                        // Try to detect from content
                        format = DetectFormatFromBytes(bytes);
                    }
                    if (format == null)
                    {
                        return new ImageLoadResult { Success = false, ErrorMessage = Resources.GetString("Error_Image_UnsupportedFormat") };
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
                return new ImageLoadResult { Success = false, ErrorMessage = Resources.GetString("Error_Image_RequestTimeout") };
            }
            catch (HttpRequestException ex)
            {
                return new ImageLoadResult { Success = false, ErrorMessage = ex.Message };
            }
            catch (Exception ex)
            {
                throw new ImageLoadException(
                    Resources.GetString("Error_Image_LoadFailed"),
                    $"Failed to load image from URL '{uri}'. {ex.GetType().Name}: {ex.Message}",
                    ex);
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
