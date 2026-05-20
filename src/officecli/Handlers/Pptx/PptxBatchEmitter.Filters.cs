// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class PptxBatchEmitter
{
    // Format keys that must NOT be emitted: derived (Get computes from cache),
    // diagnostic (relIds, cNvPr ids that resolve per package), or coordinate-
    // system (only meaningful in the source document). Same role as
    // WordBatchEmitter.SkipKeys.
    // CONSISTENCY(emit-filter-mirror): see WordBatchEmitter.Filters.cs:14.
    private static readonly HashSet<string> PptxSkipKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        // Internal relationship id — unstable across packages, see WordBatchEmitter.
        "relId",
        // OOXML cNvPr id — auto-accumulated by GenerateUniqueShapeId on Add.
        // Emitting the source id would force Add to honor it, which works for
        // free-form shapes but collides on slides where the layout already
        // contributes a placeholder at the same id slot (animations/video deck).
        // Mirrors docx WordBatchEmitter.Filters.cs treating paraId as derived:
        // emit uses positional /slide[N]/shape[K] for downstream addressing.
        "id",
        // Cached display content for unevaluated fields. The `evaluated`
        // protocol surfaces this for diagnostic Get only; replay would
        // re-emit an a:fld with stale text.
        "evaluated",
        // Aggregate child counts surface only on the Get tree (ChildCount).
        "shapeCount", "layoutCount",
        // Per-presentation metadata that auto-restamps (last-modified-by /
        // revision / created / modified). Mirrors Word's stance on
        // similar metadata.
        "revision", "lastModifiedBy", "created", "modified",
        // Default font + slide dimensions live at the root presentation
        // node, not slide-level — they roll up into a single root `set /`
        // bag in PR2 (or are already set on the blank-doc baseline).
        "defaultFont",
        // Slide `layoutType` is a derived Get-side descriptor (resolved from
        // the slide's layout relationship — "title", "twoContent", …). Replay
        // drives layout selection via `layout=<name>`; emitting layoutType
        // additionally would surface as UNSUPPORTED on AddSlide and confuse
        // users into thinking the slide lost something.
        "layoutType",
        // Speaker notes text is surfaced on the slide Format bag by
        // NodeBuilder, but AddSlide doesn't accept a `notes=` prop —
        // notes are replayed by EmitNotes as a separate add-paragraph
        // sequence under /slide[N]/notes. Without this filter, every
        // emitted slide carries a `notes=...` prop that AddSlide reports
        // as UNSUPPORTED, flipping the per-item success to false and
        // (per pre-R6 contract) the batch-level success too.
        "notes",
        // ReadShapeAnimation emits Format["animation"] / Format["animationN"]
        // on the shape node as a compound `effect-class-direction-duration`
        // string, originally used by the AddShape `animation=` prop. Dump
        // now emits a separate `add animation` row per effect
        // (EmitAnimationsForShape), so passing the compound through `add
        // shape` would double-add each effect on replay. Drop the
        // shape-level animation keys; the fine-grained rows carry trigger
        // / delay / direction / easing that the compound form loses.
        "animation",
    };

    // Shape-level `animation` is filtered above. The same readback emits
    // `animation2`, `animation3`, ... for shapes carrying multiple effects;
    // strip those alongside the singular form.
    private static bool IsShapeLevelAnimationKey(string key)
    {
        if (key.StartsWith("animation", StringComparison.OrdinalIgnoreCase)
            && key.Length > "animation".Length
            && int.TryParse(key.AsSpan("animation".Length), out _))
            return true;
        return false;
    }

    private static Dictionary<string, string> FilterEmittableProps(Dictionary<string, object?> raw)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var (key, val) in raw)
        {
            if (PptxSkipKeys.Contains(key)) continue;
            if (IsShapeLevelAnimationKey(key)) continue;
            // CONSISTENCY(effective-X-mirror): docx WordBatchEmitter.Filters.cs
            // applies the same `effective.*` prefix filter — those are read-only
            // cascade snapshots, never user-settable.
            if (key.StartsWith("effective.", StringComparison.OrdinalIgnoreCase)) continue;
            if (val == null) continue;
            string s = val switch
            {
                bool b => b ? "true" : "false",
                _ => val.ToString() ?? ""
            };
            if (s.Length > 0) result[key] = s;
        }
        // Get emits both fill=gradient (type marker) and gradient=<spec> (params).
        // ApplyShapeFill would try parsing "gradient" as a color and reject; the
        // spec via gradient= already drives the fill. Same logic for pattern.
        if (result.TryGetValue("fill", out var fillVal))
        {
            if (fillVal.Equals("gradient", StringComparison.OrdinalIgnoreCase) && result.ContainsKey("gradient"))
                result.Remove("fill");
            else if (fillVal.Equals("pattern", StringComparison.OrdinalIgnoreCase) && result.ContainsKey("pattern"))
                result.Remove("fill");
        }

        // Slide background="image" is a Get-side type marker — the embedded
        // image part is not surfaced as a re-importable path, so replay would
        // try to parse "image" as a color and reject. Drop the marker; the
        // slide will replay with default (inherited) background until image-
        // background round-trip is implemented end-to-end.
        if (result.TryGetValue("background", out var bgVal)
            && bgVal.Equals("image", StringComparison.OrdinalIgnoreCase))
            result.Remove("background");

        // Merge transitionSpeed into transition as a compound form
        // (e.g. `transition=fade` + `transitionSpeed=slow` → `transition=fade-slow`).
        // AddSlide/ApplyTransition only honor the compound form; emitting them
        // as two separate props would drop the speed on replay.
        if (result.TryGetValue("transitionSpeed", out var spd) && spd.Length > 0)
        {
            if (result.TryGetValue("transition", out var trans) && trans.Length > 0)
                result["transition"] = $"{trans}-{spd}";
            result.Remove("transitionSpeed");
        }

        // Same treatment for transitionDuration — ApplyTransition's parser
        // accepts an integer-ms modifier in the compound form (`fade-500`).
        // Emitting `transitionDuration=500` standalone fell through to the
        // generic prop-validator and the slide replay returned exit 2.
        if (result.TryGetValue("transitionDuration", out var dur) && dur.Length > 0)
        {
            if (result.TryGetValue("transition", out var trans) && trans.Length > 0)
                result["transition"] = $"{trans}-{dur}";
            result.Remove("transitionDuration");
        }

        // Shape image="true" is a NodeBuilder marker emitted for shapes
        // carrying a blipFill — Add has no shape-fill image importer, so
        // pass-through would fail prop validation. Mirror the
        // background="image" filter above; the shape replays with default
        // fill until shape image-fill round-trip is implemented.
        if (result.TryGetValue("image", out var imgVal)
            && imgVal.Equals("true", StringComparison.OrdinalIgnoreCase))
            result.Remove("image");

        return result;
    }
}
