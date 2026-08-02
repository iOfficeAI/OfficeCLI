// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Collections.Generic;

namespace OfficeCli.Core.Markdown;

/// <summary>
/// Format-neutral markdown IR — the analogue of <c>Core/Diagram</c>'s
/// <c>DiagramGraph</c>. <see cref="MarkdownParser"/> produces it from text;
/// <see cref="MarkdownBatchEmitter"/> maps it onto a target document's
/// <c>BatchItem</c> stream. Nothing here knows about OOXML or a specific
/// handler — that mapping lives in the emitter, and only in terms of the
/// neutral command envelope (BatchItem), never the SDK.
///
/// Scope is deliberately the "AI-written markdown" subset (see MarkdownParser),
/// not full CommonMark. Anything unrecognized degrades to a plain paragraph;
/// the parser never throws.
/// </summary>
public sealed class MarkdownDocument
{
    public List<MdBlock> Blocks { get; } = new();
}

/// <summary>Block-level construct. Closed hierarchy — one record per supported block.</summary>
public abstract class MdBlock { }

/// <summary>ATX heading: <c># … ######</c>. <see cref="Level"/> is 1-6.</summary>
public sealed class MdHeading : MdBlock
{
    public int Level { get; init; }
    public List<MdSpan> Inlines { get; init; } = new();
}

/// <summary>A run of text with inline formatting.</summary>
public sealed class MdParagraph : MdBlock
{
    public List<MdSpan> Inlines { get; init; } = new();
}

/// <summary>Ordered or unordered list. Nesting is represented by child items carrying sub-lists.</summary>
public sealed class MdList : MdBlock
{
    public bool Ordered { get; init; }
    /// <summary>
    /// Starting ordinal for an ordered list (CommonMark: the first item's
    /// number). 1 for unordered lists and for ordered lists that begin at 1.
    /// A list beginning "3." renders 3., 4., 5. …
    /// </summary>
    public int Start { get; init; } = 1;
    public List<MdListItem> Items { get; init; } = new();
}

public sealed class MdListItem
{
    public List<MdSpan> Inlines { get; init; } = new();
    /// <summary>
    /// Nested sub-lists (indented children), in document order. A list, not a
    /// single link: one item can own several nested segments — a marker-type
    /// switch mid-nest (`- x` then `1. y`) or a partial dedent (4→2 spaces)
    /// closes one sub-list and opens another under the same parent item.
    /// A single `Child` slot silently overwrote the earlier segment.
    /// </summary>
    public List<MdList> Children { get; } = new();

    /// <summary>
    /// Fenced code block(s) whose opener shared the item's marker line
    /// (<c>1. ```py</c>). Kept as nested item content so the list stays intact
    /// (consistent numbering) instead of the fence being hoisted to a top-level
    /// block that orphaned sibling items.
    /// </summary>
    public List<MdCodeBlock> CodeBlocks { get; } = new();
}

/// <summary>Fenced code block (```lang). Content is verbatim, no inline parsing.</summary>
public sealed class MdCodeBlock : MdBlock
{
    public string? Language { get; init; }
    public string Code { get; init; } = "";
}

/// <summary>Blockquote (<c>&gt; …</c>). Flattened to its text lines in the MVP.</summary>
public sealed class MdBlockQuote : MdBlock
{
    public List<MdSpan> Inlines { get; init; } = new();
}

/// <summary>GFM pipe table.</summary>
public sealed class MdTable : MdBlock
{
    public List<List<MdSpan>> Header { get; init; } = new();
    public List<List<List<MdSpan>>> Rows { get; init; } = new();
}

/// <summary>Thematic break (<c>---</c>, <c>***</c>, <c>___</c>).</summary>
public sealed class MdHorizontalRule : MdBlock { }

// ─────────────────────────── inline spans ───────────────────────────

/// <summary>Inline text with cumulative formatting flags. A paragraph is a list of these.</summary>
public sealed class MdSpan
{
    public string Text { get; init; } = "";
    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public bool Code { get; init; }
    /// <summary>GFM strikethrough (<c>~~text~~</c>).</summary>
    public bool Strike { get; init; }
    /// <summary>Hyperlink target for <c>[text](url)</c>, or null.</summary>
    public string? Href { get; init; }

    public MdSpan With(string text) => new()
    {
        Text = text, Bold = Bold, Italic = Italic, Code = Code, Strike = Strike, Href = Href,
    };
}
