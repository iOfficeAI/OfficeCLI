// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
/// Shared guard for handlers that do not support the `view text --range`
/// cell-range subset (docx/pptx/plugins). Kept next to the interface so the
/// error text stays identical across handlers.
/// </summary>
public static class ViewRangeGuard
{
    public static void RejectTextRange(string? range, string format)
    {
        if (range == null) return;
        throw new CliException(
            $"--range on view text is only supported for xlsx (cell ranges like 'Sheet1!A1:C10'). For {format}, use --start/--end to bound the output.")
        { Code = "invalid_value" };
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
    // range: xlsx-only cell-range subset ('Sheet1!A1:C10' or '/Sheet1/A1:C10');
    // docx/pptx throw invalid_value when non-null (use --start/--end there).
    string ViewAsText(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null, string? range = null);
    string ViewAsAnnotated(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null);
    string ViewAsOutline();
    string ViewAsStats();

    // === Structured JSON variants (for --json mode) ===
    System.Text.Json.Nodes.JsonNode ViewAsStatsJson();
    System.Text.Json.Nodes.JsonNode ViewAsOutlineJson();
    System.Text.Json.Nodes.JsonNode ViewAsTextJson(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null, string? range = null);
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
    /// When <paramref name="properties"/> carries trackChange.* keys (Word only, Run/Paragraph in Phase 4),
    /// the removal is recorded as a w:del revision instead of physically deleted.
    /// </summary>
    string? Remove(string path, Dictionary<string, string>? properties = null);
    string Move(string sourcePath, string? targetParentPath, InsertPosition? position, Dictionary<string, string>? properties = null);
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

    /// <summary>
    /// Flush the in-memory OOXML package to disk without ending the session.
    /// Only meaningful when the handler was opened with editable=true.
    /// </summary>
    void Save();

    /// <summary>
    /// Write every pending change to disk now and <b>let a failure propagate</b>.
    /// This is the whole Dispose write path (including the Word-only whole-part
    /// flush and self-closing normalization), minus closing the session: after it
    /// succeeds, Dispose does not serialize the document again.
    ///
    /// Callers are the mutating commands: persist first, then report. A command
    /// that announces success and leaves the write to the using-scope Dispose
    /// cannot take that announcement back — the atomic write-back throws, Dispose
    /// swallows it as best-effort, and the user is told an unsaved edit was saved
    /// (issue #517).
    /// </summary>
    void Persist();
}

public record ValidationError(string ErrorType, string Description, string? Path, string? Part);
