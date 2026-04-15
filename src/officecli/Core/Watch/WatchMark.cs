// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0
//
// CONSISTENCY(watch-isolation): 本文件不引用 OfficeCli.Handlers,不打开文件,不写盘。
// 见 CLAUDE.md "Watch Server Rules"。要放宽这条红线,
// grep "CONSISTENCY(watch-isolation)" 找全 watch 子系统所有文件项目级一起评审。

using System.Text.Json.Serialization;

namespace OfficeCli.Core;

/// <summary>
/// In-memory mark stored on the WatchServer. Marks are advisory annotations
/// (find/expect/note/color) attached to a document path. They live only in
/// the watch process — never persisted to disk, never written into the
/// underlying OOXML file. The watch server stores them; browsers re-locate
/// the find target in the live DOM after each refresh.
///
/// Find supports two forms (matching Set's vocabulary verbatim):
///   • literal:  find = "hello"
///   • regex:    find = r"[abc]"  OR  find = "[abc]" with regex=true flag
/// The flag is normalized into the r"..." form on insert (see WatchServer).
///
/// Tofix is a free-form display label rendered in the mark tooltip alongside
/// the find pattern. It does NOT participate in matching or staleness — when
/// a mark goes stale (find no longer hits), tofix is the human hint for
/// "what should be done about it".
/// </summary>
internal class WatchMark
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = "";

    [JsonPropertyName("path")]
    public string Path { get; set; } = "";

    [JsonPropertyName("find")]
    public string? Find { get; set; }

    [JsonPropertyName("color")]
    public string? Color { get; set; }

    [JsonPropertyName("note")]
    public string? Note { get; set; }

    [JsonPropertyName("tofix")]
    public string? Tofix { get; set; }

    /// <summary>
    /// Always an array. For literal find: 0 entries (no match → stale)
    /// or 1 entry (the literal text). For regex find: 0..N entries.
    /// Server stores whatever the client reports back; default = empty.
    /// </summary>
    [JsonPropertyName("matched_text")]
    public string[] MatchedText { get; set; } = Array.Empty<string>();

    [JsonPropertyName("stale")]
    public bool Stale { get; set; }

    [JsonPropertyName("created_at")]
    public DateTime CreatedAt { get; set; }
}

/// <summary>Request payload for the "mark" pipe command.</summary>
internal class MarkRequest
{
    [JsonPropertyName("path")]
    public string Path { get; set; } = "";

    [JsonPropertyName("find")]
    public string? Find { get; set; }

    [JsonPropertyName("color")]
    public string? Color { get; set; }

    [JsonPropertyName("note")]
    public string? Note { get; set; }

    [JsonPropertyName("tofix")]
    public string? Tofix { get; set; }
}

/// <summary>Request payload for the "unmark" pipe command.</summary>
internal class UnmarkRequest
{
    [JsonPropertyName("path")]
    public string? Path { get; set; }

    [JsonPropertyName("all")]
    public bool All { get; set; }
}

/// <summary>
/// Response payload for "mark". On success, <see cref="Id"/> is the assigned
/// mark id. On server-side rejection (invalid color, invalid path, malformed
/// request), <see cref="Error"/> carries the reason and Id is empty.
/// BUG-BT-001: callers MUST check Error first — an empty Id is not the same
/// as a null pipe response.
/// </summary>
internal class MarkResponse
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = "";

    [JsonPropertyName("error")]
    public string? Error { get; set; }
}

/// <summary>Response payload for "unmark" — returns the removed count or error.</summary>
internal class UnmarkResponse
{
    [JsonPropertyName("removed")]
    public int Removed { get; set; }

    [JsonPropertyName("error")]
    public string? Error { get; set; }
}

/// <summary>
/// Thrown by <see cref="WatchNotifier.AddMark"/> / RemoveMarks when the
/// running watch process accepts the pipe call but rejects the request
/// (invalid color, invalid path, etc.). Distinct from "no watch running"
/// (which returns null) so the CLI can surface the actual error message
/// instead of silently treating an empty id as success.
/// </summary>
public sealed class MarkRejectedException : Exception
{
    public MarkRejectedException(string message) : base(message) { }
}

/// <summary>
/// Response payload for "get-marks" — carries the current marks list plus
/// a monotonic version counter so clients can CAS on top of the SSE
/// broadcast stream without missing updates.
/// </summary>
internal class MarksResponse
{
    [JsonPropertyName("version")]
    public int Version { get; set; }

    [JsonPropertyName("marks")]
    public WatchMark[] Marks { get; set; } = Array.Empty<WatchMark>();
}

[JsonSerializable(typeof(WatchMark))]
[JsonSerializable(typeof(WatchMark[]))]
[JsonSerializable(typeof(List<WatchMark>))]
[JsonSerializable(typeof(MarkRequest))]
[JsonSerializable(typeof(UnmarkRequest))]
[JsonSerializable(typeof(MarkResponse))]
[JsonSerializable(typeof(UnmarkResponse))]
[JsonSerializable(typeof(MarksResponse))]
internal partial class WatchMarkJsonContext : JsonSerializerContext { }

/// <summary>
/// Shared JSON serializer options for the watch subsystem. Uses
/// UnsafeRelaxedJsonEscaping so CJK / non-ASCII payloads round-trip as
/// literal characters (资钱) instead of \uXXXX escapes — A complained
/// these were unreadable during manual debugging.
///
/// "Unsafe" in the encoder name refers to HTML/attribute contexts: the
/// server emits these bytes inside SSE `data:` lines and a named pipe
/// where they are consumed as raw JSON, not embedded in HTML.
///
/// AOT-friendly pattern: we build Relaxed once by cloning the source-gen
/// context's baked-in Options and overriding only the encoder, then cache
/// typed <see cref="System.Text.Json.Serialization.Metadata.JsonTypeInfo{T}"/>
/// instances that production code uses directly. The typed overloads
/// satisfy the trimmer without IL2026 warnings.
/// </summary>
internal static class WatchMarkJsonOptions
{
    public static readonly System.Text.Json.JsonSerializerOptions Relaxed =
        new(WatchMarkJsonContext.Default.Options)
        {
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        };

    public static readonly System.Text.Json.Serialization.Metadata.JsonTypeInfo<WatchMark> WatchMarkInfo =
        (System.Text.Json.Serialization.Metadata.JsonTypeInfo<WatchMark>)Relaxed.GetTypeInfo(typeof(WatchMark));

    public static readonly System.Text.Json.Serialization.Metadata.JsonTypeInfo<WatchMark[]> WatchMarkArrayInfo =
        (System.Text.Json.Serialization.Metadata.JsonTypeInfo<WatchMark[]>)Relaxed.GetTypeInfo(typeof(WatchMark[]));

    public static readonly System.Text.Json.Serialization.Metadata.JsonTypeInfo<MarksResponse> MarksResponseInfo =
        (System.Text.Json.Serialization.Metadata.JsonTypeInfo<MarksResponse>)Relaxed.GetTypeInfo(typeof(MarksResponse));
}
