// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using OfficeCli.Core;
using OfficeCli.Core.Plugins;

namespace OfficeCli.Handlers;

public static class DocumentHandlerFactory
{
    public static IDocumentHandler Open(string filePath, bool editable = false, string? password = null)
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

        // Password-protected OOXML is not a zip but an OLE/CDFV2 compound file
        // (D0 CF 11 E0…) wrapping an encrypted package. Detect it up front and
        // route through the decrypt-to-temp path; otherwise the zip-based guards
        // and the Open XML SDK below would only ever see "corrupted data".
        if (ext is ".docx" or ".xlsx" or ".pptx" && StartsWithCfbMagic(filePath))
            return OpenEncrypted(filePath, ext, editable, password);

        // CONSISTENCY(dos-hardening): reject decompression bombs before the
        // Open XML SDK / System.IO.Packaging touches the package. A few KB of
        // zip can inflate to many gigabytes and OOM the process (or, on a
        // 32-bit size-field overflow, surface only as a raw "Arithmetic
        // operation resulted in an overflow"). Only the native zip formats are
        // inspected; plugin-handled formats may not be zips and are left to
        // their own handler. See DocumentLimits for the thresholds.
        if (ext is ".docx" or ".xlsx" or ".pptx")
            GuardDecompressionBomb(filePath);

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
        catch (Exception ex) when (IsDanglingPartException(ex))
        {
            // Some producers strip a part (e.g. a sensitivity-label
            // docMetadata/LabelInfo.xml) but leave its relationship behind.
            // The SDK throws "Part: X doesn't exist in the package" on the
            // first part-graph walk, so EVERY command failed on the file even
            // though Word opens it fine (it ignores the dangling rel). Remove
            // the dangling internal relationships in-place and retry — mirrors
            // the FixXmlEncoding repair-and-retry path above.
            StripDanglingPackageRels(filePath);
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

    /// <summary>
    /// CONSISTENCY(dos-hardening): inspect the zip container's entry count and
    /// total/ratio of uncompressed bytes and reject obvious decompression
    /// bombs. Uses the central <see cref="ZipFile.OpenRead"/> directory only
    /// (no entry is actually inflated), so this is cheap. A file that is not a
    /// valid zip is left untouched — the normal open path produces the existing
    /// corrupt_file error for it.
    /// </summary>
    private static void GuardDecompressionBomb(string filePath)
    {
        ZipArchive archive;
        try
        {
            archive = ZipFile.OpenRead(filePath);
        }
        catch (InvalidDataException)
        {
            // Not a valid zip container — let OpenHandler surface corrupt_file.
            return;
        }

        using (archive)
        {
            if (archive.Entries.Count > DocumentLimits.MaxZipEntries)
                throw new CliException(
                    $"Cannot open {Path.GetFileName(filePath)}: package has {archive.Entries.Count} entries " +
                    $"(limit {DocumentLimits.MaxZipEntries}); rejected as a potential decompression bomb.")
                {
                    Code = "decompression_bomb",
                    Suggestion = "Verify the file is a genuine .docx/.xlsx/.pptx and not a crafted archive."
                };

            long totalUncompressed = 0;
            long totalCompressed = 0;
            foreach (var entry in archive.Entries)
            {
                totalUncompressed += entry.Length;
                totalCompressed += entry.CompressedLength;

                if (totalUncompressed > DocumentLimits.MaxUncompressedBytes)
                    throw new CliException(
                        $"Cannot open {Path.GetFileName(filePath)}: uncompressed size exceeds " +
                        $"{DocumentLimits.MaxUncompressedBytes / (1024 * 1024 * 1024)} GiB; " +
                        $"rejected as a potential decompression bomb.")
                    {
                        Code = "decompression_bomb",
                        Suggestion = "Verify the file is a genuine .docx/.xlsx/.pptx and not a crafted archive."
                    };
            }

            // Ratio check only once there is enough data to be meaningful (a few
            // KB of highly-compressible XML legitimately exceeds the ratio).
            if (totalCompressed > 64 * 1024 &&
                totalUncompressed / Math.Max(1, totalCompressed) > DocumentLimits.MaxCompressionRatio)
                throw new CliException(
                    $"Cannot open {Path.GetFileName(filePath)}: compression ratio " +
                    $"{totalUncompressed / Math.Max(1, totalCompressed)}× exceeds {DocumentLimits.MaxCompressionRatio}×; " +
                    $"rejected as a potential decompression bomb.")
                {
                    Code = "decompression_bomb",
                    Suggestion = "Verify the file is a genuine .docx/.xlsx/.pptx and not a crafted archive."
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
            _      => TryOpenViaPlugin(filePath, ext, editable)
                   ?? throw UnsupportedTypeException(ext)
        };
    }

    /// <summary>True if the file begins with the OLE/CDFV2 magic — i.e. it is a
    /// compound file (a password-encrypted OOXML, or an OLE-wrapped legacy doc),
    /// not a plain zip-based .docx/.xlsx/.pptx.</summary>
    private static bool StartsWithCfbMagic(string filePath)
    {
        Span<byte> head = stackalloc byte[8];
        try
        {
            using var fs = File.OpenRead(filePath);
            if (fs.Read(head) < 8) return false;
        }
        catch { return false; }
        ReadOnlySpan<byte> magic = stackalloc byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
        return head.SequenceEqual(magic);
    }

    /// <summary>
    /// Open a password-protected (Agile-encrypted) OOXML file: decrypt the
    /// package to a temporary plaintext copy, open a normal handler on it, and
    /// wrap it so any save re-encrypts back to the original path. The temp copy
    /// is deleted when the handler is disposed.
    /// </summary>
    private static IDocumentHandler OpenEncrypted(string filePath, string ext, bool editable, string? password)
    {
        byte[] cfb = File.ReadAllBytes(filePath);
        switch (MsOffCrypto.Classify(cfb))
        {
            case MsOffCrypto.Scheme.StandardUnsupported:
                throw new CliException(
                    $"Cannot open {Path.GetFileName(filePath)}: it uses the legacy 'Standard' " +
                    $"encryption scheme, which officecli does not support (only modern 'Agile' encryption).")
                {
                    Code = "unsupported_encryption",
                    Suggestion = "Re-save it with a password from Office 2013+ (Agile), or remove the password."
                };
            case MsOffCrypto.Scheme.NotEncrypted:
                // CDFV2 but not an encrypted OOXML (e.g. a real OLE/.doc binary
                // mislabeled .docx). Fall through to the normal corrupt-file UX.
                throw new CliException(
                    $"Cannot open {Path.GetFileName(filePath)}: not a valid OOXML package " +
                    $"(OLE compound file without an EncryptionInfo stream).")
                {
                    Code = "corrupt_file",
                    Suggestion = "Verify the file is a genuine .docx/.xlsx/.pptx."
                };
        }

        // --password wins; otherwise fall back to the OFFICECLI_PASSWORD env var.
        // The env path keeps the secret out of argv (visible in `ps`) — better
        // for unattended/agent use, where the password comes from a secret store.
        if (string.IsNullOrEmpty(password))
            password = Environment.GetEnvironmentVariable("OFFICECLI_PASSWORD");

        if (string.IsNullOrEmpty(password))
            throw new CliException(
                $"{Path.GetFileName(filePath)} is password-protected. Re-run with --password <password> (or set OFFICECLI_PASSWORD).")
            {
                Code = "password_required",
                Suggestion = "officecli <command> <file> --password <password>   # or: OFFICECLI_PASSWORD=<password> officecli <command> <file>"
            };

        byte[] plain;
        try
        {
            plain = MsOffCrypto.Decrypt(cfb, password);
        }
        catch (MsOffCrypto.WrongPasswordException)
        {
            throw new CliException($"Incorrect password for {Path.GetFileName(filePath)}.")
            {
                Code = "wrong_password",
                Suggestion = "Check the password and try again."
            };
        }

        // Decrypted plaintext goes to a private temp file with the right
        // extension so the existing handlers + bomb guard work unchanged.
        var tempPlain = Path.Combine(Path.GetTempPath(), $"officecli-dec-{Guid.NewGuid():N}{ext}");
        File.WriteAllBytes(tempPlain, plain);

        try
        {
            GuardDecompressionBomb(tempPlain);   // the decrypted package is a real zip
            var inner = OpenHandler(tempPlain, ext, editable);
            return new EncryptedDocumentHandler(inner, tempPlain, filePath, password, editable);
        }
        catch
        {
            try { File.Delete(tempPlain); } catch { /* best-effort */ }
            throw;
        }
    }

    /// <summary>
    /// Look for an installed plugin that handles <paramref name="ext"/> and, if
    /// found, return a handler that delegates to it. Returns null when no
    /// plugin is installed — callers fall back to the unsupported-type error.
    ///
    /// dump-reader: per plugins/plugin-protocol.md §2.1, the plugin emits a batch
    /// of officecli commands describing the foreign source; main replays them
    /// into a fresh native file whose extension comes from the plugin's
    /// manifest <c>target</c> (docx/xlsx/pptx). The result is cached as a
    /// sibling file <c>&lt;source-stem&gt;.&lt;target&gt;</c> next to the
    /// source so subsequent invocations skip the plugin entirely (regenerated
    /// when the source mtime is newer than the sibling's, or when the sibling
    /// has been deleted). All edits target the sibling file, not the original
    /// source.
    ///
    /// format-handler: not yet wired; resolved plugins produce a clear
    /// "found but not yet wired" exception until the proxy lands.
    /// </summary>
    private static IDocumentHandler? TryOpenViaPlugin(string filePath, string ext, bool editable)
    {
        var dumpReader = PluginRegistry.FindFor(PluginKind.DumpReader, ext);
        if (dumpReader is not null)
        {
            var targetExt = dumpReader.Manifest.ResolveTargetExtension();
            var sibling = Path.ChangeExtension(filePath, targetExt);
            var needRegen = !File.Exists(sibling)
                || File.GetLastWriteTimeUtc(filePath) > File.GetLastWriteTimeUtc(sibling);

            if (needRegen)
            {
                var converted = DumpReaderInvoker.Run(filePath, ext);

                // Some plugins (e.g. Word interop on .doc) inherently write a
                // converted native file in the source directory as a side
                // effect of their conversion path. If the sibling now exists
                // and is current, prefer it over the batch-replayed copy:
                // it's the plugin's direct conversion, higher fidelity than
                // going through batch round-trip serialization.
                var siblingFresh = File.Exists(sibling)
                    && File.GetLastWriteTimeUtc(sibling) >= File.GetLastWriteTimeUtc(filePath);

                if (siblingFresh)
                {
                    try { File.Delete(converted.ConvertedPath); } catch { /* tmp will age out */ }
                }
                else
                {
                    try
                    {
                        var bytes = File.ReadAllBytes(converted.ConvertedPath);
                        File.WriteAllBytes(sibling, bytes);
                        try { File.Delete(converted.ConvertedPath); } catch { /* tmp will age out */ }
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine(
                            $"[note] could not write sibling {Path.GetFileName(sibling)} ({ex.Message}); falling back to temp file (will reconvert next run)");
                        return OpenHandlerWithRetry(converted.ConvertedPath, targetExt, editable);
                    }
                }
                Console.Error.WriteLine(
                    $"[note] generated {Path.GetFileName(sibling)} from {Path.GetFileName(filePath)}; reusing on future runs (delete or rename it to force reconversion)");
            }

            // The sibling file may be transiently locked right after a fresh
            // plugin run (Word/Excel COM server lingering, Defender scan,
            // OneDrive sync). Retry briefly before surfacing the lock to the
            // user.
            return OpenHandlerWithRetry(sibling, targetExt, editable);
        }

        var formatHandler = PluginRegistry.FindFor(PluginKind.FormatHandler, ext);
        if (formatHandler is not null)
        {
            var session = new FormatHandlerSession(filePath, formatHandler);
            try
            {
                session.Start(editable);
                return new FormatHandlerProxy(session);
            }
            catch
            {
                session.Dispose();
                throw;
            }
        }

        return null;
    }

    private static IDocumentHandler OpenHandlerWithRetry(string path, string ext, bool editable)
    {
        Exception? last = null;
        for (int attempt = 0; attempt < 8; attempt++)
        {
            try { return OpenHandler(path, ext, editable); }
            catch (IOException ex) { last = ex; Thread.Sleep(150 * (attempt + 1)); }
        }
        throw last!;
    }

    private static CliException UnsupportedTypeException(string ext) =>
        new CliException(
            $"Unsupported file type: {ext}. Supported: .docx, .xlsx, .pptx. " +
            $"Other formats may be opened via plugins — run `officecli plugins list` to see installed plugins, " +
            $"or see plugins/plugin-protocol.md for installation paths.")
        {
            Code = "unsupported_type",
            ValidValues = [".docx", ".xlsx", ".pptx"]
        };

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

    private static bool IsDanglingPartException(Exception ex)
    {
        for (var e = ex; e != null; e = e.InnerException)
        {
            if (e.Message.Contains("doesn't exist in the package", StringComparison.OrdinalIgnoreCase)
                || e.Message.Contains("Specified part does not exist", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    /// <summary>
    /// Remove internal (non-External) relationships whose target part is
    /// missing from the package. Word tolerates such dangling rels; the SDK
    /// refuses to open the file. Only .rels entries are touched, and only the
    /// dangling Relationship nodes are dropped.
    /// </summary>
    private static void StripDanglingPackageRels(string filePath)
    {
        using var zip = ZipFile.Open(filePath, ZipArchiveMode.Update);
        var names = new HashSet<string>(zip.Entries.Select(e => e.FullName), StringComparer.OrdinalIgnoreCase);
        foreach (var entry in zip.Entries.ToList())
        {
            if (!entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;
            string content;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                content = reader.ReadToEnd();

            System.Xml.Linq.XDocument xdoc;
            try { xdoc = System.Xml.Linq.XDocument.Parse(content); }
            catch { continue; }
            var ns = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/package/2006/relationships";
            // Base directory the .rels' relative targets resolve against:
            // "x/_rels/y.xml.rels" -> "x"; the root "_rels/.rels" -> "".
            var relsDir = (Path.GetDirectoryName(entry.FullName) ?? "").Replace('\\', '/');
            var baseDir = relsDir.EndsWith("_rels", StringComparison.OrdinalIgnoreCase)
                ? relsDir.Substring(0, relsDir.Length - "_rels".Length).TrimEnd('/')
                : relsDir;
            bool changed = false;
            foreach (var rel in xdoc.Root!.Elements(ns + "Relationship").ToList())
            {
                if (string.Equals((string?)rel.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase))
                    continue;
                var target = (string?)rel.Attribute("Target");
                if (string.IsNullOrEmpty(target)) continue;
                var resolved = target.StartsWith("/")
                    ? target.TrimStart('/')
                    : (baseDir.Length > 0 ? baseDir + "/" + target : target);
                var segs = new List<string>();
                foreach (var seg in resolved.Split('/'))
                {
                    if (seg == "..") { if (segs.Count > 0) segs.RemoveAt(segs.Count - 1); }
                    else if (seg != "." && seg.Length > 0) segs.Add(seg);
                }
                resolved = string.Join("/", segs);
                if (!names.Contains(resolved))
                {
                    rel.Remove();
                    changed = true;
                }
            }
            if (!changed) continue;
            entry.Delete();
            var newEntry = zip.CreateEntry(entry.FullName, CompressionLevel.Optimal);
            using var writer = new StreamWriter(newEntry.Open(), new UTF8Encoding(false));
            writer.Write(xdoc.Declaration != null
                ? xdoc.Declaration + xdoc.ToString(System.Xml.Linq.SaveOptions.DisableFormatting)
                : xdoc.ToString(System.Xml.Linq.SaveOptions.DisableFormatting));
        }
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
