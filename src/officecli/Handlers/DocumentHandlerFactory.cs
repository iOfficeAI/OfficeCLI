// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static class DocumentHandlerFactory
{
    public static IDocumentHandler Open(string filePath, bool editable = false)
    {
        if (!File.Exists(filePath))
            throw new CliException($"File not found: {filePath}")
            {
                Code = "file_not_found",
                Suggestion = "Check the file path. Use an absolute path or a path relative to the current directory.",
                Help = "officecli create <path> --type docx|xlsx|pptx"
            };

        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        try
        {
            return OpenHandler(filePath, ext, editable);
        }
        catch (Exception ex) when (IsEncodingException(ex))
        {
            // Files created by python-pptx (lxml) use encoding="ascii" which Open XML SDK rejects.
            // Fix the XML declarations in-place and retry.
            FixXmlEncoding(filePath);
            return OpenHandler(filePath, ext, editable);
        }
        catch (DocumentFormat.OpenXml.Packaging.OpenXmlPackageException ex)
        {
            throw new InvalidOperationException($"Cannot open {Path.GetFileName(filePath)}: {ex.Message}", ex);
        }
    }

    private static IDocumentHandler OpenHandler(string filePath, string ext, bool editable)
    {
        return ext switch
        {
            ".docx" => new WordHandler(filePath, editable),
            ".xlsx" => new ExcelHandler(filePath, editable),
            ".pptx" => new PowerPointHandler(filePath, editable),
            ".hwpx" => new HwpxHandler(filePath, editable),
            ".hwp" => throw new CliException(
                "Binary .hwp files require pyhwp for reading. " +
                "Convert to .hwpx in Hancom Office, or install pyhwp and convert: " +
                "pip install pyhwp && hwp5txt input.hwp > output.txt")
            {
                Code = "unsupported_format",
                Suggestion = "Save as .hwpx in Hancom Office (File → Save As → HWPX)",
                Help = "pip install pyhwp && hwp5txt input.hwp > output.txt"
            },
            _ => throw new CliException($"Unsupported file type: {ext}. Supported: .docx, .xlsx, .pptx, .hwpx")
            {
                Code = "unsupported_type",
                ValidValues = [".docx", ".xlsx", ".pptx", ".hwpx"]
            }
        };
    }

    private static bool IsEncodingException(Exception ex)
    {
        // The exception may be thrown directly or wrapped inside another exception
        for (var e = ex; e != null; e = e.InnerException)
        {
            if (e.Message.Contains("Encoding format is not supported", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    /// <summary>
    /// Rewrite XML declarations inside an OOXML package that use unsupported encodings
    /// (e.g. encoding="ascii") to encoding="UTF-8".
    /// </summary>
    private static void FixXmlEncoding(string filePath)
    {
        using var zip = ZipFile.Open(filePath, ZipArchiveMode.Update);
        foreach (var entry in zip.Entries.ToList())
        {
            if (!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                !entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                continue;

            string content;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                content = reader.ReadToEnd();

            // Match <?xml ... encoding="xxx" ?> and replace non-standard encodings
            var fixed_ = Regex.Replace(content,
                @"(<\?xml\b[^?]*?\bencoding\s*=\s*"")(?!UTF-8|utf-8|UTF-16|utf-16)[^""]*("")",
                "${1}UTF-8${2}");

            if (fixed_ == content) continue;

            // Rewrite the entry
            entry.Delete();
            var newEntry = zip.CreateEntry(entry.FullName, CompressionLevel.Optimal);
            using var writer = new StreamWriter(newEntry.Open(), new UTF8Encoding(false));
            writer.Write(fixed_);
        }
    }
}
