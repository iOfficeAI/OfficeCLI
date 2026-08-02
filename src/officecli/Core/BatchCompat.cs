// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// NEWLINE-SEMANTICS-V2 dump versioning. A v2 dump starts with
/// {"command":"meta","dumpVersion":2} and encodes docx soft line breaks as
/// '\v' in text props, with '\n' reserved for paragraph boundaries (unified
/// with pptx and the Google Docs API convention). A pre-v2 dump encoded a
/// docx soft break as '\n', so replaying one under v2 semantics would explode
/// every soft break into a paragraph split; such a dump opts back into the old
/// reading by declaring {"command":"meta","dumpVersion":1}, which makes
/// <see cref="PrepareForReplay"/> rewrite '\n' → '\v' in text props.
///
/// A batch that declares no version is read with CURRENT semantics. It used to
/// be read as legacy, on the theory that only dumps are ever replayed — but a
/// hand-written or generated batch carries no meta item either, so the shim
/// fired on ordinary input and one batch item behaved differently from the
/// identical single command: `text=A\nB` split into two paragraphs from the
/// CLI and produced one paragraph with a soft break through batch. Versioning
/// is opt-in for the file format that has a version; everything else follows
/// the documented rule that '\n' is a paragraph and '\v' is a line break.
/// </summary>
public static class BatchCompat
{
    public const int CurrentDumpVersion = 2;

    public static BatchItem MetaItem() => new()
    {
        Command = "meta",
        DumpVersion = CurrentDumpVersion,
    };

    /// <summary>
    /// Strip meta items and apply the legacy-newline shim when the batch
    /// targets a .docx and explicitly declares a dumpVersion below 2.
    /// A batch with no meta item is left alone. Call before executing any
    /// items. Idempotent.
    /// </summary>
    public static void PrepareForReplay(List<BatchItem> items, string targetFilePath)
    {
        int? declared = null;
        for (int i = items.Count - 1; i >= 0; i--)
        {
            if (string.Equals(items[i].Command, "meta", StringComparison.OrdinalIgnoreCase))
            {
                if (items[i].DumpVersion is { } v && v > (declared ?? 0)) declared = v;
                items.RemoveAt(i);
            }
        }
        if ((declared ?? CurrentDumpVersion) >= 2) return;
        if (!targetFilePath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)) return;

        foreach (var item in items)
        {
            // Legacy docx dumps carry soft breaks as '\n' inside text-bearing
            // fields. Rewrite to '\v' so v2 handlers rebuild <w:br/> instead
            // of splitting paragraphs. Scope: the "text" prop plus the
            // item-level Text field — the only carriers the v1 emitters used.
            if (item.Props != null && item.Props.TryGetValue("text", out var t)
                && t != null && t.IndexOf('\n') >= 0)
                item.Props["text"] = t.Replace("\n", "\v");
            if (item.Text != null && item.Text.IndexOf('\n') >= 0)
                item.Text = item.Text.Replace("\n", "\v");
        }
    }
}
