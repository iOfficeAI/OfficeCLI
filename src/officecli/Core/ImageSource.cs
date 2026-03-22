// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Net.Http;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeCli.Core;

/// <summary>
/// Resolves image sources from file paths, data URIs, or HTTP(S) URLs into a stream and content type.
/// Supports:
///   - Local file path: "/tmp/logo.png", "C:\images\photo.jpg"
///   - Data URI: "data:image/png;base64,iVBOR..."
///   - HTTP(S) URL: "https://example.com/image.png"
///
/// Returns a content type string compatible with OpenXmlPart.AddImagePart() (e.g. ImagePartType.Png).
/// </summary>
public static class ImageSource
{
    /// <summary>
    /// Resolve an image source string into a stream and content type string.
    /// Caller is responsible for disposing the returned stream.
    /// The returned contentType can be passed directly to AddImagePart().
    /// </summary>
    public static (Stream Stream, PartTypeInfo ContentType) Resolve(string source)
    {
        if (string.IsNullOrWhiteSpace(source))
            throw new ArgumentException("Image source cannot be empty");

        // Data URI: data:image/png;base64,iVBOR...
        if (source.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            return ResolveDataUri(source);

        // HTTP(S) URL
        if (source.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            source.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            return ResolveUrl(source);

        // Local file path
        return ResolveFile(source);
    }

    /// <summary>
    /// Determine content type string from a file extension (with or without dot).
    /// Returns a value usable with AddImagePart().
    /// </summary>
    public static PartTypeInfo ExtensionToContentType(string extension)
    {
        var ext = extension.TrimStart('.').ToLowerInvariant();
        return ext switch
        {
            "png" => ImagePartType.Png,
            "jpg" or "jpeg" => ImagePartType.Jpeg,
            "gif" => ImagePartType.Gif,
            "bmp" => ImagePartType.Bmp,
            "tif" or "tiff" => ImagePartType.Tiff,
            "emf" => ImagePartType.Emf,
            "wmf" => ImagePartType.Wmf,
            "svg" => ImagePartType.Svg,
            _ => throw new ArgumentException($"Unsupported image format: .{ext}. Supported: png, jpg, gif, bmp, tiff, emf, wmf, svg")
        };
    }

    private static (Stream, PartTypeInfo) ResolveFile(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"Image file not found: {path}");

        var contentType = ExtensionToContentType(Path.GetExtension(path));
        return (File.OpenRead(path), contentType);
    }

    private static (Stream, PartTypeInfo) ResolveDataUri(string dataUri)
    {
        // Format: data:[<mediatype>][;base64],<data>
        var commaIdx = dataUri.IndexOf(',');
        if (commaIdx < 0)
            throw new ArgumentException("Invalid data URI: missing comma separator");

        var header = dataUri[..commaIdx]; // e.g. "data:image/png;base64"
        var data = dataUri[(commaIdx + 1)..];

        if (!header.Contains("base64", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Only base64-encoded data URIs are supported");

        // Extract MIME type
        var mimeStart = header.IndexOf(':') + 1;
        var mimeEnd = header.IndexOf(';');
        var mime = mimeEnd > mimeStart ? header[mimeStart..mimeEnd] : header[mimeStart..];

        var contentType = MimeToContentType(mime);
        var bytes = Convert.FromBase64String(data);
        return (new MemoryStream(bytes), contentType);
    }

    private static (Stream, PartTypeInfo) ResolveUrl(string url)
    {
        using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
        client.DefaultRequestHeaders.Add("User-Agent", "OfficeCLI");

        var response = client.GetAsync(url).GetAwaiter().GetResult();
        response.EnsureSuccessStatusCode();

        var bytes = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
        var stream = new MemoryStream(bytes);

        // Try content-type header first
        var serverMime = response.Content.Headers.ContentType?.MediaType;
        if (!string.IsNullOrEmpty(serverMime) && TryMimeToContentType(serverMime, out var ct))
            return (stream, ct);

        // Fallback: extract extension from URL path (strip query string)
        var uri = new Uri(url);
        var ext = Path.GetExtension(uri.AbsolutePath);
        if (!string.IsNullOrEmpty(ext))
            return (stream, ExtensionToContentType(ext));

        // Last resort: sniff magic bytes
        if (TrySniffContentType(bytes, out var sniffed))
            return (stream, sniffed);

        throw new ArgumentException($"Cannot determine image type from URL: {url}. Specify format via file extension or content-type header.");
    }

    private static PartTypeInfo MimeToContentType(string mime)
    {
        if (TryMimeToContentType(mime, out var ct)) return ct;
        throw new ArgumentException($"Unsupported MIME type: {mime}. Supported: image/png, image/jpeg, image/gif, image/bmp, image/tiff, image/svg+xml");
    }

    private static bool TryMimeToContentType(string mime, out PartTypeInfo contentType)
    {
        contentType = mime.ToLowerInvariant() switch
        {
            "image/png" => ImagePartType.Png,
            "image/jpeg" or "image/jpg" => ImagePartType.Jpeg,
            "image/gif" => ImagePartType.Gif,
            "image/bmp" => ImagePartType.Bmp,
            "image/tiff" => ImagePartType.Tiff,
            "image/svg+xml" => ImagePartType.Svg,
            "image/emf" or "image/x-emf" => ImagePartType.Emf,
            "image/wmf" or "image/x-wmf" => ImagePartType.Wmf,
            _ => default
        };
        return contentType != default;
    }

    private static bool TrySniffContentType(byte[] bytes, out PartTypeInfo contentType)
    {
        contentType = default;
        if (bytes.Length < 4) return false;

        // PNG: 89 50 4E 47
        if (bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
        { contentType = ImagePartType.Png; return true; }

        // JPEG: FF D8 FF
        if (bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
        { contentType = ImagePartType.Jpeg; return true; }

        // GIF: GIF8
        if (bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x38)
        { contentType = ImagePartType.Gif; return true; }

        // BMP: BM
        if (bytes[0] == 0x42 && bytes[1] == 0x4D)
        { contentType = ImagePartType.Bmp; return true; }

        return false;
    }
}
