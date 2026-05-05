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

        // CONSISTENCY(corrupt-file-rejection): a 0-byte file is silently
        // accepted by Open XML SDK 3.x in read-write mode (it materialises an
        // empty Package), but the resulting handler returns a fake root node
        // with no parts. CLI commands that follow then report success and
        // exit 0 even though the document is unusable. Reject the file
        // up-front so the same file_not_found / corrupt_file UX applies that
        // direct-mode (read-only) Open already gave for 0-byte files.
        if (new FileInfo(filePath).Length == 0)
            throw new CliException($"Cannot open {Path.GetFileName(filePath)}: file is 0 bytes (not a valid Office document).")
            {
                Code = "corrupt_file",
                Suggestion = "Recreate the file with: officecli create <path>"
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
            throw new CliException($"Cannot open {Path.GetFileName(filePath)}: {ex.Message}", ex)
            {
                Code = "corrupt_file",
                Suggestion = "Verify the file is a valid .docx/.xlsx/.pptx (e.g. unzip -t)."
            };
        }
        catch (System.IO.FileFormatException ex)
        {
            // Thrown by System.IO.Packaging when the file is not a valid OOXML zip container.
            throw new CliException($"Cannot open {Path.GetFileName(filePath)}: {ex.Message}", ex)
            {
                Code = "corrupt_file",
                Suggestion = "Verify the file is a valid .docx/.xlsx/.pptx (e.g. unzip -t)."
            };
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
