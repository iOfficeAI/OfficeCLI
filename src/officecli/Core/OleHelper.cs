// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;

namespace OfficeCli.Core;

/// <summary>
/// Shared helpers for OLE (Object Linking and Embedding) support across
/// Word/Excel/PowerPoint handlers. Covers:
/// - ProgID auto-detection from file extension
/// - Mapping src file extensions to the right embedded PartTypeInfo
/// - A tiny placeholder PNG used as the visual icon for new OLE objects
/// - Populating canonical DocumentNode.Format fields from an embedded part
///
/// Design: all three handlers consume the same helper so that a single call
/// site governs progId defaults, content-type decisions, and node shape.
/// This keeps the "ole" node schema consistent across .docx/.xlsx/.pptx.
/// </summary>
internal static class OleHelper
{
    /// <summary>
    /// Detect the OLE ProgID to use when the caller did not supply one.
    /// Returns identifiers that match what Word/Excel/PowerPoint register
    /// at install time on Windows; all three are version-12 ProgIDs that
    /// real Office uses for embedded round-tripping. Unknown extensions
    /// fall back to "Package", the generic "wrapper for an opaque file"
    /// ProgID that any Office host will open via its registered handler.
    /// </summary>
    public static string DetectProgId(string srcPath)
    {
        var ext = Path.GetExtension(srcPath).TrimStart('.').ToLowerInvariant();
        return ext switch
        {
            "docx" or "docm" or "dotx" or "dotm" => "Word.Document.12",
            "doc" => "Word.Document.8",
            "xlsx" or "xlsm" or "xlsb" or "xltx" or "xltm" => "Excel.Sheet.12",
            "xls" => "Excel.Sheet.8",
            "pptx" or "pptm" or "ppsx" or "ppsm" or "potx" or "potm" => "PowerPoint.Show.12",
            "ppt" => "PowerPoint.Show.8",
            "pdf" => "AcroExch.Document",
            "vsdx" or "vsdm" or "vsd" => "Visio.Drawing",
            _ => "Package",
        };
    }

    /// <summary>
    /// Classifier for the content-type axis: Office files get an
    /// <see cref="EmbeddedPackagePart"/> with the matching OOXML MIME,
    /// everything else gets a generic <see cref="EmbeddedObjectPart"/>.
    /// This mirrors how real Office writes OLE objects — OOXML documents
    /// embed as x/vnd.openxmlformats-* package parts, binary or legacy
    /// content lands in the generic "oleObject" bucket.
    /// </summary>
    public enum EmbeddingKind
    {
        /// <summary>Use EmbeddedPackagePart (for .docx/.xlsx/.pptx and their macro/template siblings).</summary>
        Package,
        /// <summary>Use EmbeddedObjectPart (for arbitrary binaries — PDF, Visio, .bin, etc.).</summary>
        Object,
    }

    /// <summary>
    /// Decide whether a source file should be embedded as a Package part
    /// (strongly-typed OOXML container) or a generic Object part.
    /// </summary>
    public static EmbeddingKind ClassifyKind(string srcPath)
    {
        var ext = Path.GetExtension(srcPath).TrimStart('.').ToLowerInvariant();
        return ext switch
        {
            "docx" or "docm" or "dotx" or "dotm"
            or "xlsx" or "xlsm" or "xlsb" or "xltx" or "xltm"
            or "pptx" or "pptm" or "ppsx" or "ppsm" or "potx" or "potm"
            or "sldx" or "sldm" or "xlam" or "ppam" or "thmx"
                => EmbeddingKind.Package,
            _ => EmbeddingKind.Object,
        };
    }

    /// <summary>
    /// Map an OOXML-family extension to its EmbeddedPackagePartType entry.
    /// Returns null if the extension is not a recognized Office format,
    /// in which case the caller should use <see cref="EmbeddedObjectPart"/>
    /// with a generic content type.
    /// </summary>
    public static PartTypeInfo? GetPackagePartTypeInfo(string srcPath)
    {
        var ext = Path.GetExtension(srcPath).TrimStart('.').ToLowerInvariant();
        return ext switch
        {
            "docx" => EmbeddedPackagePartType.Docx,
            "docm" => EmbeddedPackagePartType.Docm,
            "dotx" => EmbeddedPackagePartType.Dotx,
            "dotm" => EmbeddedPackagePartType.Dotm,
            "xlsx" => EmbeddedPackagePartType.Xlsx,
            "xlsm" => EmbeddedPackagePartType.Xlsm,
            "xlsb" => EmbeddedPackagePartType.Xlsb,
            "xltx" => EmbeddedPackagePartType.Xltx,
            "xltm" => EmbeddedPackagePartType.Xltm,
            "xlam" => EmbeddedPackagePartType.Xlam,
            "pptx" => EmbeddedPackagePartType.Pptx,
            "pptm" => EmbeddedPackagePartType.Pptm,
            "ppsx" => EmbeddedPackagePartType.Ppsx,
            "ppsm" => EmbeddedPackagePartType.Ppsm,
            "potx" => EmbeddedPackagePartType.Potx,
            "potm" => EmbeddedPackagePartType.Potm,
            "ppam" => EmbeddedPackagePartType.Ppam,
            "sldx" => EmbeddedPackagePartType.Sldx,
            "sldm" => EmbeddedPackagePartType.Sldm,
            "thmx" => EmbeddedPackagePartType.Thmx,
            _ => null,
        };
    }

    /// <summary>
    /// Add an embedded part (package or generic object) to the given host
    /// part, feed it the source file bytes, and return the rel id.
    /// Works for any parent that supports embedded parts: MainDocumentPart,
    /// WorksheetPart, SlidePart.
    /// </summary>
    public static (string RelId, OpenXmlPart Part) AddEmbeddedPart(OpenXmlPart host, string srcPath, string? hostDocumentPath = null)
    {
        if (!File.Exists(srcPath))
            throw new FileNotFoundException($"OLE source file not found: {srcPath}");

        // Warn (don't throw) when the source file is zero bytes and it is NOT
        // a self-embed. Self-embed intentionally writes a zero-byte placeholder
        // (see CONSISTENCY(ole-self-embed) block below) and should stay silent.
        // Non-self-embed 0-byte files usually indicate a truncated or missing
        // payload — the user deserves a visible warning so they know the
        // embedded bytes are empty. We still proceed with the embed to match
        // the existing "silently ignored → visibly ignored" contract.
        var isSelfEmbed = hostDocumentPath != null && IsSameFile(srcPath, hostDocumentPath);
        if (!isSelfEmbed && new FileInfo(srcPath).Length == 0)
        {
            Console.Error.WriteLine(
                $"Warning: OLE source file is empty (0 bytes): {srcPath}. Document will embed an empty payload.");
        }

        var kind = ClassifyKind(srcPath);
        OpenXmlPart part;
        if (kind == EmbeddingKind.Package)
        {
            var pt = GetPackagePartTypeInfo(srcPath)
                ?? EmbeddedPackagePartType.Xlsx; // should never hit, classified as Package
            part = host switch
            {
                MainDocumentPart mdp => mdp.AddEmbeddedPackagePart(pt),
                WorksheetPart wp => wp.AddEmbeddedPackagePart(pt),
                SlidePart sp => sp.AddEmbeddedPackagePart(pt),
                HeaderPart hp => hp.AddEmbeddedPackagePart(pt),
                FooterPart fp => fp.AddEmbeddedPackagePart(pt),
                _ => throw new InvalidOperationException(
                    $"Host part type {host.GetType().Name} does not support embedded packages"),
            };
        }
        else
        {
            // Generic: use content-type that Office writes for "Package" OLE.
            // The literal OOXML content type for an oleObject is documented as
            // "application/vnd.openxmlformats-officedocument.oleObject".
            var ct = "application/vnd.openxmlformats-officedocument.oleObject";
            part = host switch
            {
                MainDocumentPart mdp => mdp.AddEmbeddedObjectPart(ct),
                WorksheetPart wp => wp.AddEmbeddedObjectPart(ct),
                SlidePart sp => sp.AddEmbeddedObjectPart(ct),
                HeaderPart hp => hp.AddEmbeddedObjectPart(ct),
                FooterPart fp => fp.AddEmbeddedObjectPart(ct),
                _ => throw new InvalidOperationException(
                    $"Host part type {host.GetType().Name} does not support embedded objects"),
            };
        }

        // CONSISTENCY(ole-self-embed): when srcPath refers to the host
        // document itself, the SDK holds an exclusive package lock and any
        // FileStream.Open() against srcPath fails with IOException. In that
        // case feed a zero-byte placeholder payload so the OLE element and
        // relationship are still created — callers can Get() the resulting
        // node and reopen the document without corruption. The user-facing
        // contract is: "self-embed is allowed and does not crash, but the
        // embedded bytes are a placeholder rather than the host's literal
        // snapshot" (which would require cloning the in-memory package).
        if (hostDocumentPath != null && IsSameFile(srcPath, hostDocumentPath))
        {
            using var emptyMs = new MemoryStream(Array.Empty<byte>());
            part.FeedData(emptyMs);
            var selfRelId = host.GetIdOfPart(part);
            return (selfRelId, part);
        }

        // First try FileShare.ReadWrite so concurrent writers do not crash;
        // if that still fails (exclusive package lock / non-self-embed race),
        // surface the exception to the caller with an actionable hint —
        // commonly it is an officecli resident/watch process holding the
        // source file open, in which case `officecli close <path>` unblocks
        // the embed. We keep the detection-free approach (just add the hint
        // to every IOException) so the helper stays dependency-free and the
        // message is useful even for non-officecli holders.
        //
        // CONSISTENCY(ole-orphan-cleanup): if FileStream.Open() or FeedData()
        // fails after the host part has been created, delete the dangling
        // part so we don't leave an orphan EmbeddedPackagePart/EmbeddedObjectPart
        // on the host (which would inflate part counts and survive into
        // the saved file). The part was just added by AddEmbeddedPackagePart/
        // AddEmbeddedObjectPart above — at this point nothing else references
        // it, so DeletePart is safe.
        try
        {
            FileStream src;
            try
            {
                src = new FileStream(srcPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            }
            catch (IOException ioEx)
            {
                throw new IOException(
                    $"Cannot read OLE source file '{srcPath}': the file is locked by another process. " +
                    $"If an officecli resident or watch process has this file open, run " +
                    $"'officecli close {srcPath}' first, then retry.", ioEx);
            }
            using (src)
            {
                part.FeedData(src);
            }
        }
        catch
        {
            try { host.DeletePart(part); } catch { /* best effort */ }
            throw;
        }
        var relId = host.GetIdOfPart(part);
        return (relId, part);
    }

    /// <summary>
    /// Returns true if <paramref name="candidatePath"/> resolves to the same
    /// file as <paramref name="hostDocumentPath"/>. Used by handlers to
    /// detect self-embed Set(src=hostPath) so they can substitute a
    /// zero-byte or placeholder payload instead of crashing when the SDK
    /// holds an exclusive package lock on the host file.
    /// </summary>
    public static bool IsSameFile(string candidatePath, string hostDocumentPath)
    {
        if (string.IsNullOrEmpty(candidatePath) || string.IsNullOrEmpty(hostDocumentPath))
            return false;
        try
        {
            var a = Path.GetFullPath(candidatePath);
            var b = Path.GetFullPath(hostDocumentPath);
            return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }


    /// <summary>
    /// Populate canonical OLE fields on a DocumentNode from the backing
    /// embedded part. Reads content type and byte length so consumers see
    /// the same shape regardless of whether the part was EmbeddedObject or
    /// EmbeddedPackage.
    /// </summary>
    public static void PopulateFromPart(DocumentNode node, OpenXmlPart part, string? progId = null)
    {
        node.Type = "ole";
        node.Format["objectType"] = "ole";
        if (!string.IsNullOrEmpty(progId))
        {
            node.Format["progId"] = progId;
            if (string.IsNullOrEmpty(node.Text))
                node.Text = progId;
        }
        node.Format["contentType"] = part.ContentType;
        try
        {
            using var s = part.GetStream();
            node.Format["fileSize"] = s.Length;
        }
        catch
        {
            // part stream may be transient during write; ignore
        }
    }

    /// <summary>
    /// Minimal valid 1x1 transparent PNG used as the icon preview for
    /// newly-inserted OLE objects. Office requires a visual placeholder;
    /// the size is irrelevant because the host shape's explicit extents
    /// govern display dimensions. This is the same byte sequence used by
    /// <c>PowerPointHandler.AddMedia</c> for its poster fallback, known
    /// to decode cleanly in every consumer we test against.
    /// </summary>
    public static byte[] PlaceholderIconPng => _placeholderPng;

    // 1x1 transparent PNG, precomputed. Verified valid by the existing
    // PowerPointHandler media poster path.
    private static readonly byte[] _placeholderPng =
    {
        0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
        0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
        0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,0x89,
        0x00,0x00,0x00,0x0D,0x49,0x44,0x41,0x54,
        0x08,0xD7,0x63,0x60,0x60,0x60,0x60,0x00,0x00,0x00,0x05,0x00,0x01,0x87,0xA1,0x4E,0xD4,
        0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,0x42,0x60,0x82,
    };

    /// <summary>
    /// Compute default icon dimensions in EMU when the caller didn't supply
    /// width/height. 2 inches × 0.75 inches matches what Office uses for a
    /// default "show as icon" OLE frame, sized to fit the file-type label.
    /// </summary>
    public const long DefaultOleWidthEmu = 1828800;  // 2 inches
    public const long DefaultOleHeightEmu = 685800;   //  0.75 inches

    /// <summary>
    /// Validate a COM ProgID string against the well-known Windows COM
    /// constraints: the identifier must be 1..39 characters long and must
    /// not start with a digit. OLE spec (MSDN "ProgID") is explicit on both
    /// rules. Handlers previously accepted arbitrary strings silently; this
    /// method gives users an early, actionable error instead of writing an
    /// invalid OLE element that Office refuses to open.
    /// </summary>
    public static void ValidateProgId(string progId)
    {
        if (progId == null) return;
        if (progId.Length > 39)
            throw new ArgumentException(
                $"progId '{progId}' exceeds 39 characters (limit: 39, actual: {progId.Length}).");
        if (progId.Length > 0 && char.IsDigit(progId[0]))
            throw new ArgumentException(
                $"progId '{progId}' cannot start with a digit.");
        // COM ProgID character set: letters, digits, '.', '_', '-'. Anything
        // else (notably XML-unsafe characters like '<', '>', '&', '"') would
        // either corrupt the OOXML progId attribute or be rejected by Office
        // on reopen. Reject early with an actionable error instead of letting
        // bad bytes land in the package.
        foreach (var ch in progId)
        {
            if (!(char.IsLetterOrDigit(ch) || ch == '.' || ch == '_' || ch == '-'))
                throw new ArgumentException(
                    $"progId '{progId}' contains invalid characters. Only letters, digits, '.', '_', '-' are allowed.");
        }
    }

    /// <summary>
    /// Normalize and validate the caller-supplied <c>display</c> property
    /// for an OLE object. Canonical values are <c>"icon"</c> (show the file
    /// as a clickable icon preview) and <c>"content"</c> (show the embedded
    /// file's first page as a live picture). Any other value — including
    /// ambiguous synonyms like <c>"embed"</c>, <c>"invisible"</c>, numbers,
    /// or boolean strings — is rejected with <see cref="ArgumentException"/>
    /// so the user is told their input was wrong instead of silently
    /// falling back to "icon". Used by Word/PPT Add and Set.
    /// </summary>
    public static string NormalizeOleDisplay(string value)
    {
        if (value == null)
            throw new ArgumentException(
                "Invalid display value ''. Expected 'icon' or 'content'.");
        var v = value.Trim().ToLowerInvariant();
        if (v == "icon") return "icon";
        if (v == "content") return "content";
        throw new ArgumentException(
            $"Invalid display value '{value}'. Expected 'icon' or 'content'.");
    }

    /// <summary>
    /// Known OLE Add/Set property keys shared across Word/PPT/Excel. Used by
    /// <see cref="WarnOnUnknownOleProps"/> to surface silently-ignored
    /// properties via stderr. Kept as a single union so the three handlers
    /// stay consistent — per-handler differences (e.g. Excel's "anchor"
    /// range string) are all represented here.
    /// </summary>
    private static readonly HashSet<string> KnownOleProps = new(StringComparer.OrdinalIgnoreCase)
    {
        "src", "path", "progId", "progid",
        "width", "height", "x", "y",
        "icon", "display", "name",
        "anchor",
    };

    /// <summary>
    /// Emit a single-line stderr warning for every property key in
    /// <paramref name="properties"/> that is not in <see cref="KnownOleProps"/>.
    /// The Add handler signature returns a string and cannot carry a
    /// structured warning list back to the caller, so we surface unknown
    /// keys via Console.Error to match the "silently ignored → visibly
    /// ignored" expectation. No-op when <paramref name="properties"/> is
    /// null or empty.
    /// </summary>
    public static void WarnOnUnknownOleProps(Dictionary<string, string>? properties)
    {
        if (properties == null || properties.Count == 0) return;
        foreach (var key in properties.Keys)
        {
            if (!KnownOleProps.Contains(key))
                Console.Error.WriteLine($"warning: unknown ole property '{key}' — ignored");
        }
    }

    // ==================== Shared Add helpers ====================
    //
    // The following methods extract duplicated boilerplate that previously
    // appeared verbatim in Word/Excel/PowerPoint AddOle handlers.

    /// <summary>
    /// Validate and extract the required <c>src</c> (or <c>path</c>) property
    /// from the caller-supplied dictionary. Throws
    /// <see cref="ArgumentException"/> when neither key is present or the
    /// value is blank.
    /// </summary>
    public static string RequireSource(Dictionary<string, string>? properties)
    {
        properties ??= new Dictionary<string, string>();
        if (!properties.TryGetValue("src", out var srcPath)
            && !properties.TryGetValue("path", out srcPath))
            throw new ArgumentException("'src' property is required for ole type");
        if (string.IsNullOrWhiteSpace(srcPath))
            throw new ArgumentException("'src' property for ole type cannot be empty");
        return srcPath;
    }

    /// <summary>
    /// Resolve the ProgID from explicit property → auto-detected from
    /// extension, then validate. Replaces the 4-line fallback chain that
    /// was duplicated in every handler.
    /// </summary>
    public static string ResolveProgId(Dictionary<string, string> properties, string srcPath)
    {
        var progId = properties.GetValueOrDefault("progId")
            ?? properties.GetValueOrDefault("progid")
            ?? DetectProgId(srcPath);
        ValidateProgId(progId);
        return progId;
    }

    /// <summary>
    /// Create the icon preview <see cref="ImagePart"/> on the given host
    /// part — either from the user-supplied <c>icon</c> property or the
    /// default 1×1 placeholder PNG. Returns the relationship id.
    /// </summary>
    public static (ImagePart Part, string RelId) CreateIconPart(OpenXmlPart host, Dictionary<string, string> properties)
    {
        ImagePart iconPart;
        if (properties.TryGetValue("icon", out var iconPath) && !string.IsNullOrWhiteSpace(iconPath))
        {
            var (iconStream, iconType) = ImageSource.Resolve(iconPath);
            using var _ = iconStream;
            iconPart = AddImagePartTo(host, iconType);
            iconPart.FeedData(iconStream);
        }
        else
        {
            iconPart = AddImagePartTo(host, ImagePartType.Png);
            using var ms = new MemoryStream(PlaceholderIconPng);
            iconPart.FeedData(ms);
        }
        return (iconPart, host.GetIdOfPart(iconPart));
    }

    /// <summary>
    /// Dispatch <see cref="OpenXmlPart.AddImagePart"/> to the correct
    /// concrete host type. Covers all part types that can own OLE objects.
    /// </summary>
    private static ImagePart AddImagePartTo(OpenXmlPart host, PartTypeInfo type)
        => host switch
        {
            MainDocumentPart mdp => mdp.AddImagePart(type),
            HeaderPart hp => hp.AddImagePart(type),
            FooterPart fp => fp.AddImagePart(type),
            WorksheetPart wp => wp.AddImagePart(type),
            SlidePart sp => sp.AddImagePart(type),
            DrawingsPart dp => dp.AddImagePart(type),
            _ => throw new InvalidOperationException(
                $"Host part type {host.GetType().Name} does not support image parts"),
        };
}
