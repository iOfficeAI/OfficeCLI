// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Resolves file sources from local paths, HTTP(S) URLs, or data URIs into a seekable stream.
/// Unified counterpart to <see cref="ImageSource"/> for non-image binary files (media, 3D models, CSV, etc.).
///
/// Supports:
///   - Local file path: "/tmp/model.glb", "C:\media\video.mp4"
///   - HTTP(S) URL: "https://example.com/video.mp4"
///   - Data URI: "data:video/mp4;base64,AAAA..."
///
/// Returns a MemoryStream (always seekable) and the detected file extension.
/// </summary>
internal static class FileSource
{
    /// <summary>
    /// Resolve a source string into a seekable MemoryStream and file extension (with dot, e.g. ".glb").
    /// Caller is responsible for disposing the returned stream.
    /// </summary>
    public static (MemoryStream Stream, string Extension) Resolve(string source)
    {
        if (string.IsNullOrWhiteSpace(source))
            throw new ArgumentException("File source cannot be empty");

        if (source.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            return ResolveDataUri(source);

        if (source.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            source.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            return ResolveUrl(source);

        return ResolveFile(source);
    }

    /// <summary>
    /// Check whether a string looks like a resolvable source (URL, data URI, or existing local file).
    /// Useful for distinguishing file/URL sources from inline data (e.g. CSV inline vs file path).
    /// </summary>
    public static bool IsResolvable(string source)
    {
        if (string.IsNullOrWhiteSpace(source)) return false;
        if (source.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return true;
        if (source.StartsWith("http://", StringComparison.OrdinalIgnoreCase)) return true;
        if (source.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) return true;
        return File.Exists(source);
    }

    /// <summary>
    /// Resolve a source to text lines (for CSV/text data).
    /// </summary>
    public static string[] ResolveLines(string source)
    {
        var (stream, _) = Resolve(source);
        using (stream)
        {
            using var reader = new StreamReader(stream);
            var text = reader.ReadToEnd();
            return text.Split('\n')
                .Select(l => l.TrimEnd('\r'))
                .ToArray();
        }
    }

    private static (MemoryStream, string) ResolveFile(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"File not found: {path}");
        return (new MemoryStream(File.ReadAllBytes(path)), Path.GetExtension(path).ToLowerInvariant());
    }

    private static (MemoryStream, string) ResolveUrl(string url)
    {
        using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
        client.DefaultRequestHeaders.Add("User-Agent", "OfficeCLI");

        var response = client.GetAsync(url).GetAwaiter().GetResult();
        response.EnsureSuccessStatusCode();

        var bytes = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();

        // Try extension from URL path
        var uri = new Uri(url);
        var ext = Path.GetExtension(uri.AbsolutePath).ToLowerInvariant();

        // Fallback: infer from content-type header
        if (string.IsNullOrEmpty(ext))
        {
            var mime = response.Content.Headers.ContentType?.MediaType;
            ext = MimeToExtension(mime);
        }

        return (new MemoryStream(bytes), ext);
    }

    private static (MemoryStream, string) ResolveDataUri(string dataUri)
    {
        var commaIdx = dataUri.IndexOf(',');
        if (commaIdx < 0)
            throw new ArgumentException("Invalid data URI: missing comma separator");

        var header = dataUri[..commaIdx];
        var data = dataUri[(commaIdx + 1)..];

        if (!header.Contains("base64", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Only base64-encoded data URIs are supported");

        var mimeStart = header.IndexOf(':') + 1;
        var mimeEnd = header.IndexOf(';');
        var mime = mimeEnd > mimeStart ? header[mimeStart..mimeEnd] : header[mimeStart..];

        var ext = MimeToExtension(mime);
        return (new MemoryStream(Convert.FromBase64String(data)), ext);
    }

    private static string MimeToExtension(string? mime)
    {
        if (string.IsNullOrEmpty(mime)) return "";
        return mime.ToLowerInvariant() switch
        {
            // Video
            "video/mp4" => ".mp4",
            "video/quicktime" => ".mov",
            "video/x-msvideo" or "video/avi" => ".avi",
            "video/x-ms-wmv" => ".wmv",
            "video/mpeg" => ".mpg",
            "video/webm" => ".webm",
            // Audio
            "audio/mpeg" or "audio/mp3" => ".mp3",
            "audio/wav" or "audio/x-wav" => ".wav",
            "audio/mp4" or "audio/x-m4a" => ".m4a",
            "audio/x-ms-wma" => ".wma",
            "audio/ogg" => ".ogg",
            // 3D
            "model/gltf-binary" => ".glb",
            // Text/data
            "text/csv" => ".csv",
            "text/plain" => ".txt",
            _ => ""
        };
    }
}
