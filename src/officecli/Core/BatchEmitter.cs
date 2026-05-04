// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Handlers;

namespace OfficeCli.Core;

/// <summary>
/// Walks an opened handler's document tree and emits a sequence of BatchItem
/// rows that, when replayed against a blank document of the same format,
/// reconstruct the original document.
///
/// <para>
/// This is the core of the `officecli dump --format batch` pipeline. The
/// emit relies on the OOXML schema reflection fallback in
/// <see cref="TypedAttributeFallback"/> + <see cref="GenericXmlQuery"/>:
/// any leaf property that Get reads can be re-applied via Add/Set, so
/// emit just transcribes Format keys directly without per-property
/// allowlisting.
/// </para>
///
/// <para>
/// Scope (v0): docx body paragraphs and their runs. Resources (styles,
/// numbering, theme, headers, footers, sections, comments, footnotes,
/// endnotes) are NOT yet emitted — they will be added in subsequent
/// passes once round-trip is proven on the body skeleton.
/// </para>
/// </summary>
public static class BatchEmitter
{
    /// <summary>Emit a batch sequence for a Word document.</summary>
    public static List<BatchItem> EmitWord(WordHandler word)
    {
        var items = new List<BatchItem>();

        var bodyNode = word.Get("/body");
        if (bodyNode.Children == null) return items;

        // Phase: body — paragraphs and their runs.
        // Emit `add /body p --prop text=...` for each paragraph, then
        // optional `set /body/p[i]/r[j]` rows for run-level formatting that
        // doesn't fit on the parent paragraph alone.
        int pIndex = 0;
        foreach (var pChild in bodyNode.Children)
        {
            if (pChild.Type != "paragraph" && pChild.Type != "p") continue;
            pIndex++;
            EmitParagraph(word, pChild.Path, pIndex, items);
        }

        return items;
    }

    private static void EmitParagraph(WordHandler word, string sourcePath, int targetIndex, List<BatchItem> items)
    {
        var pNode = word.Get(sourcePath);

        // Drop OOXML-internal/derived keys that are not user-provided input
        // (paraId / rsid / textId regenerate on save; effective.* are computed
        // readbacks; basedOn.path is a Get-side derived pointer).
        var props = FilterEmittableProps(pNode.Format);

        // Collapse paragraph-level text only when the paragraph has a single
        // run (or no runs). When there are multiple runs with mixed
        // formatting, emit the paragraph empty and add runs individually.
        var runs = (pNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "run" || c.Type == "r")
            .ToList();

        if (runs.Count <= 1)
        {
            // Single-run shortcut: bake text into the paragraph add.
            if (runs.Count == 1)
            {
                var runProps = FilterEmittableProps(runs[0].Format);
                foreach (var (k, v) in runProps)
                {
                    if (!props.ContainsKey(k)) props[k] = v;
                }
                if (!string.IsNullOrEmpty(runs[0].Text))
                    props["text"] = runs[0].Text!;
            }
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = "/body",
                Type = "p",
                Props = props.Count > 0 ? props : null
            });
            return;
        }

        // Multi-run paragraph: emit the paragraph first (no text), then
        // each run as a child add. The paragraph index in target equals
        // targetIndex because emit is sequential and target is empty.
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = "/body",
            Type = "p",
            Props = props.Count > 0 ? props : null
        });

        int rIdx = 0;
        foreach (var run in runs)
        {
            rIdx++;
            var rProps = FilterEmittableProps(run.Format);
            if (!string.IsNullOrEmpty(run.Text))
                rProps["text"] = run.Text!;
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = $"/body/p[{targetIndex}]",
                Type = "r",
                Props = rProps.Count > 0 ? rProps : null
            });
        }
    }

    // Format keys that must NOT be emitted: derived (computed by Get, not
    // user-set), unstable (regenerate on save), or coordinate-system
    // (paths that only make sense in the source document).
    private static readonly HashSet<string> SkipKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        // Derived / Get-only
        "basedOn.path",
        // Unstable IDs (regenerate on save)
        "paraId", "textId", "rsidR", "rsidRDefault", "rsidRPr", "rsidP", "rsidTr",
        // Effective / cascade readbacks (already absorbed into base keys)
        // Anything starting with "effective." is dropped via prefix below.
    };

    private static Dictionary<string, string> FilterEmittableProps(Dictionary<string, object?> raw)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (key, val) in raw)
        {
            if (SkipKeys.Contains(key)) continue;
            if (key.StartsWith("effective.", StringComparison.OrdinalIgnoreCase)) continue;
            // Source-tracking metadata (e.g. fontSize.cs holds .cs source path)
            if (key.EndsWith(".cs.source", StringComparison.OrdinalIgnoreCase)) continue;

            if (val == null) continue;
            // Booleans serialize to lowercase "true"/"false" — what Add expects.
            var s = val switch
            {
                bool b => b ? "true" : "false",
                _ => val.ToString() ?? ""
            };
            if (s.Length > 0) result[key] = s;
        }
        return result;
    }
}
