// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Represents where to insert an element: by index, after an anchor, or before an anchor.
/// At most one field is set. All null = append to end.
/// </summary>
public class InsertPosition
{
    public int? Index { get; init; }
    public string? After { get; init; }
    public string? Before { get; init; }

    public static InsertPosition AtIndex(int idx) => new() { Index = idx };
    public static InsertPosition AfterElement(string path) => new() { After = path };
    public static InsertPosition BeforeElement(string path) => new() { Before = path };

    /// <summary>
    /// Resolve After/Before anchor to a 0-based index among children.
    /// If this is already an Index or null, returns Index as-is.
    /// anchorFinder: given the anchor path, returns the 0-based index of that element among siblings, or throws.
    /// childCount: total number of children of the relevant type.
    /// </summary>
    public int? Resolve(Func<string, int> anchorFinder, int childCount)
    {
        if (Index.HasValue) return Index;
        if (After != null)
        {
            var anchorIdx = anchorFinder(After);
            return anchorIdx + 1 >= childCount ? null : anchorIdx + 1; // null = append
        }
        if (Before != null)
        {
            return anchorFinder(Before);
        }
        return null; // append
    }
}

/// <summary>
/// Common interface for all document types (Word/Excel/PowerPoint).
/// Each handler implements the three-layer architecture:
///   - Semantic layer: view (text/annotated/outline/stats/issues)
///   - Query layer: get, query, set
///   - Raw layer: raw XML access
/// </summary>
public interface IDocumentHandler : IDisposable
{
    // === Semantic Layer ===
    string ViewAsText(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null);
    string ViewAsAnnotated(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null);
    string ViewAsOutline();
    string ViewAsStats();

    // === Structured JSON variants (for --json mode) ===
    System.Text.Json.Nodes.JsonNode ViewAsStatsJson();
    System.Text.Json.Nodes.JsonNode ViewAsOutlineJson();
    System.Text.Json.Nodes.JsonNode ViewAsTextJson(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null);
    List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null);

    // === Query Layer ===
    DocumentNode Get(string path, int depth = 1);
    List<DocumentNode> Query(string selector);
    /// <summary>
    /// Returns list of prop names that were not applied (unsupported for this element type).
    /// </summary>
    List<string> Set(string path, Dictionary<string, string> properties);
    string Add(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties);
    /// <summary>
    /// Remove element at path. Returns an optional warning message (e.g. formula cells affected by shift).
    /// </summary>
    string? Remove(string path);
    string Move(string sourcePath, string? targetParentPath, InsertPosition? position);
    string CopyFrom(string sourcePath, string targetParentPath, InsertPosition? position);

    // === Raw Layer ===
    string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null);
    void RawSet(string partPath, string xpath, string action, string? xml);

    /// <summary>
    /// Create a new part (chart, header, footer, etc.) and return its relationship ID and accessible path.
    /// </summary>
    (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null);

    /// <summary>
    /// Validate the document against OpenXML schema and return any errors.
    /// </summary>
    List<ValidationError> Validate();

    /// <summary>
    /// Extract the binary payload backing a node (ole/picture/media/embedded)
    /// to <paramref name="destPath"/>. Returns <c>true</c> if the node has a
    /// backing part and the bytes were written, <c>false</c> if the node has
    /// no binary payload (e.g. it is a text paragraph or table cell).
    /// <paramref name="contentType"/> receives the part's MIME type on success;
    /// <paramref name="byteCount"/> receives the number of bytes written.
    /// </summary>
    bool TryExtractBinary(string path, string destPath, out string? contentType, out long byteCount);
}

/// <summary>
/// Standardized validation error/warning codes (aligned with kordoc v2.2.6).
/// </summary>
public static class ValidationCodes
{
    // Errors (critical — document may not open correctly)
    public const string Encrypted = "ENCRYPTED";
    public const string DrmProtected = "DRM_PROTECTED";
    public const string ZipBomb = "ZIP_BOMB";
    public const string Corrupted = "CORRUPTED";
    public const string NoSections = "NO_SECTIONS";
    public const string ZipEmpty = "ZIP_EMPTY";
    public const string ZipCorrupt = "ZIP_CORRUPT";
    public const string OpfMissing = "OPF_MISSING";
    public const string XmlMalformed = "XML_MALFORMED";
    public const string IdRefOrphan = "IDREF_ORPHAN";
    public const string TableStructure = "TABLE_STRUCTURE";
    public const string BinDataMissing = "BINDATA_MISSING";
    public const string BinDataOrphan = "BINDATA_ORPHAN";
    public const string FieldPairMismatch = "FIELD_PAIR_MISMATCH";
    public const string SectionMismatch = "SECTION_MISMATCH";

    // Warnings (non-critical — document opens but may have issues)
    public const string TruncatedTable = "TRUNCATED_TABLE";
    public const string MalformedXml = "MALFORMED_XML_MINOR";
    public const string PartialParse = "PARTIAL_PARSE";
    public const string NamespaceMissing = "NAMESPACE_MISSING";
    public const string NamespaceMismatch = "NAMESPACE_MISMATCH";
    public const string StaleIdRef = "STALE_IDREF";
    public const string EmptySection = "EMPTY_SECTION";
    public const string LargeFile = "LARGE_FILE";
    public const string DeprecatedElement = "DEPRECATED_ELEMENT";
    public const string MergedCellOverlap = "MERGED_CELL_OVERLAP";
}

public record ValidationError(
    string ErrorType,
    string Description,
    string? Path,
    string? Part,
    IssueSeverity Severity = IssueSeverity.Error);
