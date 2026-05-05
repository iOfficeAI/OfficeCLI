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
/// Scope (v0.5): docx body paragraphs (with run formatting) + tables (single
/// paragraph + single run per cell, common case). Resources (styles,
/// numbering, theme, headers, footers, sections, comments, footnotes,
/// endnotes) and richer cell contents are NOT yet emitted — follow-up
/// passes will add them.
/// </para>
/// </summary>
public static class BatchEmitter
{
    /// <summary>Emit a batch sequence for a Word document.</summary>
    public static List<BatchItem> EmitWord(WordHandler word)
    {
        var items = new List<BatchItem>();

        // Phase order matters: resources first so body refs (style=Heading1,
        // numId=3, etc.) resolve when the paragraph adds reach them on replay.
        // Numbering must come BEFORE styles — list-style definitions
        // (Heading paragraphs with numPr) reference numId values, so style
        // adds that carry `numId=N` need /numbering to already hold N.
        EmitNumberingRaw(word, items);
        EmitStyles(word, items);
        EmitThemeRaw(word, items);
        EmitSettingsRaw(word, items);
        EmitSection(word, items);
        EmitHeadersFooters(word, items);
        var paraIdToTargetIdx = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        EmitBody(word, items, paraIdToTargetIdx);
        EmitComments(word, items, paraIdToTargetIdx);
        return items;
    }

    private static void EmitThemeRaw(WordHandler word, List<BatchItem> items)
    {
        // Theme carries clrScheme + fontScheme + fmtScheme — pure structured
        // XML that users rarely modify property-by-property; the natural
        // operation is "swap the entire theme block". Raw-set replace fits
        // that model exactly. Word.Raw returns the literal string
        // "(no theme)" when the part is missing — gate on a leading '<' so
        // we only emit when there's real XML to ship.
        string xml;
        try { xml = word.Raw("/theme"); }
        catch { return; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return;

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/theme",
            Xpath = "/a:theme",
            Action = "replace",
            Xml = xml
        });
    }

    private static void EmitSettingsRaw(WordHandler word, List<BatchItem> items)
    {
        // Settings carries dozens of feature flags + compat shims that
        // surface on root.Format only piecemeal — and not all of them are
        // wired through Set's case table. Wholesale raw-set is the simplest
        // way to keep Word feature toggles (evenAndOddHeaders, mirrorMargins,
        // schema-pegged compat options, …) round-tripped without
        // per-property allowlisting.
        string xml;
        try { xml = word.Raw("/settings"); }
        catch { return; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return;

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/settings",
            Xpath = "/w:settings",
            Action = "replace",
            Xml = xml
        });
    }

    private static void EmitNumberingRaw(WordHandler word, List<BatchItem> items)
    {
        // Numbering models list templates (abstractNum + num pairs, each
        // abstractNum holds 9 levels with their own pPr / numFmt / lvlText).
        // Reconstructing this through typed Add would mean another emitter
        // in itself; for v0.5 we ship the entire <w:numbering> XML wholesale
        // via raw-set. The blank document creates an empty numbering part,
        // so a single replace on the part root is sufficient.
        string xml;
        try { xml = word.Raw("/numbering"); }
        catch { return; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return;
        // Skip when numbering is empty (just `<w:numbering/>` with no children).
        if (!xml.Contains("<w:abstractNum") && !xml.Contains("<w:num "))
            return;

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/numbering",
            Xpath = "/w:numbering",
            Action = "replace",
            Xml = xml
        });
    }

    private static void EmitHeadersFooters(WordHandler word, List<BatchItem> items)
    {
        var root = word.Get("/");
        if (root.Children == null) return;
        // BUG-R4-T2: header/footer parts carry no `type` key on Get; the
        // section's `headerRef.default|first|even` (and `footerRef.*`)
        // entries are the only place the part's role is recorded. Build a
        // reverse lookup so EmitHeaderFooterPart can emit the right
        // `type` prop (default/first/even) instead of always emitting
        // "default" — which on a doc with both default + first headers
        // throws "Header of type 'default' already exists" on replay.
        var headerPathToType = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var footerPathToType = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        // BUG-R5-2 / R5-F2: headerRef.<type> / footerRef.<type> live on
        // **section** nodes (see WordHandler.Query.cs:902), not on root.
        // The earlier R4 fix scanned root.Format and silently found nothing,
        // so every emitted header/footer was typed "default" — round-trip
        // failed when a doc had both default + first headers. Walk all
        // section children to build the path→type map.
        void HarvestRefs(DocumentNode node)
        {
            foreach (var (key, val) in node.Format)
            {
                if (val == null) continue;
                var s = val.ToString();
                if (string.IsNullOrEmpty(s)) continue;
                if (key.StartsWith("headerRef.", StringComparison.OrdinalIgnoreCase))
                {
                    var t = key["headerRef.".Length..];
                    if (!headerPathToType.ContainsKey(s)) headerPathToType[s] = t;
                }
                else if (key.StartsWith("footerRef.", StringComparison.OrdinalIgnoreCase))
                {
                    var t = key["footerRef.".Length..];
                    if (!footerPathToType.ContainsKey(s)) footerPathToType[s] = t;
                }
            }
        }
        HarvestRefs(root);
        try
        {
            var sections = word.Query("section");
            if (sections != null)
            {
                foreach (var sec in sections) HarvestRefs(sec);
            }
        }
        catch { /* missing section info — fall through with default typing */ }

        int hIdx = 0, fIdx = 0;
        foreach (var child in root.Children)
        {
            if (child.Type == "header")
            {
                hIdx++;
                var t = headerPathToType.TryGetValue(child.Path, out var ht) ? ht : "default";
                EmitHeaderFooterPart(word, child.Path, "header", hIdx, items, t);
            }
            else if (child.Type == "footer")
            {
                fIdx++;
                var t = footerPathToType.TryGetValue(child.Path, out var ft) ? ft : "default";
                EmitHeaderFooterPart(word, child.Path, "footer", fIdx, items, t);
            }
        }
    }

    private static void EmitHeaderFooterPart(WordHandler word, string sourcePath, string kind,
                                             int targetIndex, List<BatchItem> items,
                                             string subTypeOverride = "default")
    {
        var partNode = word.Get(sourcePath);
        var paras = (partNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "paragraph" || c.Type == "p")
            .ToList();
        // partNode.Format does not expose `type`; the caller resolves the
        // role (default/first/even) from the section's headerRef.* / footerRef.*
        // map and passes it via subTypeOverride.
        var subType = subTypeOverride;

        // Create the part with just its role (default/first/even). AddHeader/
        // AddFooter seed an empty auto paragraph; EmitParagraph(autoPresent:
        // true) on paras[0] then routes through CollapseFieldChains so a
        // PAGE-field header (the canonical case) round-trips as a typed
        // `add field` row instead of being baked into static "1" text on the
        // seed paragraph (BUG-R4-T3). Run-level formatting on multi-run
        // first paragraphs is preserved by the per-run emit path below.
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = "/",
            Type = kind,
            Props = new Dictionary<string, string> { ["type"] = subType }
        });

        var partTargetPath = $"/{kind}[{targetIndex}]";
        for (int p = 0; p < paras.Count; p++)
        {
            EmitParagraph(word, paras[p].Path, partTargetPath, p + 1, items, autoPresent: p == 0);
        }
    }

    private static void EmitComments(WordHandler word, List<BatchItem> items,
                                     Dictionary<string, int> paraIdToTargetIdx)
    {
        var comments = word.Query("comment");
        foreach (var c in comments)
        {
            var props = FilterEmittableProps(c.Format);
            if (!string.IsNullOrEmpty(c.Text))
                props["text"] = c.Text!;
            // Map anchoredTo (source paraId path) -> target paragraph index.
            // anchoredTo looks like "/body/p[@paraId=00100000]"; parse and
            // resolve via the paraId map we built during EmitBody.
            string parentTarget = "/body/p[1]";  // safe fallback to first body para
            if (props.TryGetValue("anchoredTo", out var anchor))
            {
                var pid = ExtractParaId(anchor);
                if (pid != null && paraIdToTargetIdx.TryGetValue(pid, out var idx))
                    parentTarget = $"/body/p[{idx}]";
                props.Remove("anchoredTo");
            }
            // The comment id is allocated by AddComment on the target side;
            // do not propagate the source id (would conflict on replay).
            props.Remove("id");
            // BUG-R7-04 (T-4): previously dropped `date` so dump→replay always
            // re-stamped the comment with the SDK's "now". That breaks
            // archival / audit-trail use cases where the source timestamp is
            // load-bearing. Preserve it; AddComment accepts an explicit
            // ISO-8601 date and the SDK will use it instead of stamping.

            items.Add(new BatchItem
            {
                Command = "add",
                Parent = parentTarget,
                Type = "comment",
                Props = props
            });
        }
    }

    // Emit a body-level SDT (Content Control) as a typed `add /body --type sdt`
    // row. Get exposes type, alias, tag, items (dropdown/combobox), editable,
    // and the visible text — all of which AddSdt round-trips. Without this,
    // SDTs were silently dropped from dump output (BUG-R2-06 / R2-3).
    private static void EmitSdt(WordHandler word, string sourcePath, List<BatchItem> items)
    {
        DocumentNode sdt;
        try { sdt = word.Get(sourcePath); }
        catch { return; }

        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        // Whitelist Get-canonical keys that AddSdt consumes. `editable` is a
        // Get readback (negation of `lock`), the source-side `id` is allocated
        // at creation, so neither is forwarded.
        foreach (var key in new[] { "type", "alias", "tag", "items", "format" })
        {
            if (sdt.Format.TryGetValue(key, out var v) && v != null)
            {
                var s = v.ToString() ?? "";
                if (s.Length > 0) props[key] = s;
            }
        }
        if (!string.IsNullOrEmpty(sdt.Text))
            props["text"] = sdt.Text!;

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = "/body",
            Type = "sdt",
            Props = props
        });
    }

    private static string? ExtractParaId(string anchorPath)
    {
        var m = System.Text.RegularExpressions.Regex.Match(anchorPath, @"@paraId=([0-9A-Fa-f]+)");
        return m.Success ? m.Groups[1].Value : null;
    }

    // Root-level keys that round-trip via `set /`. Includes section page
    // layout, document protection, doc-level grid + defaults. Excludes
    // metadata that auto-updates on save (created/modified timestamps,
    // lastModifiedBy, package author/title — those re-stamp anyway).
    private static readonly HashSet<string> RootScalarKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        // Section page layout (mirrors body's trailing sectPr)
        "pageWidth", "pageHeight", "orientation",
        "marginTop", "marginBottom", "marginLeft", "marginRight",
        "pageStart", "pageNumFmt",
        "titlePage", "direction", "rtlGutter",
        "lineNumbers", "lineNumberCountBy",
        // Multi-column section layout. Get exposes these as canonical keys
        // (columns, columnSpace, columns.equalWidth) and Set's case table
        // accepts all three (WordHandler.Set.SectionLayout.cs). Without them
        // here, multi-column documents silently revert to single column on
        // round-trip.
        "columns", "columnSpace",
        // Document-level final-section break type (oddPage / evenPage /
        // continuous). Set / accepts section.type but the canonical Get
        // surfaces it bare; emit so the trailing sectPr's type survives.
        "section.type",
        // Document protection
        "protection", "protectionEnforced",
        // Document grid (CJK-aware line layout)
        "charSpacingControl",
        // pPrDefault CJK toggles — without these, Word inserts an automatic
        // space between Latin runs and adjacent CJK glyphs ("2025年" →
        // "2025 年"). Templates that explicitly disable autoSpaceDE/DN
        // depend on these surviving the round-trip.
        "kinsoku", "overflowPunct", "autoSpaceDE", "autoSpaceDN",
    };

    // Dotted-prefix groups that round-trip wholesale via `set /`. Each
    // sub-key is forwarded as-is; the schema-reflection layer routes the
    // dotted path into the right OOXML target.
    private static readonly string[] RootPrefixGroups = new[]
    {
        "docDefaults.",
        "docGrid.",
        // columns.equalWidth / columns.separator etc. roundtrip via the
        // canonical dotted form Get already emits.
        "columns.",
    };

    private static void EmitSection(WordHandler word, List<BatchItem> items)
    {
        var root = word.Get("/");
        // protectionEnforced has no Set case in WordHandler — `set / protectionEnforced=...`
        // emits a WARNING on every replay. The only meaningful encoding is
        // when protection is non-default; for protection="none" the
        // enforced flag is implicitly false anyway. Drop the noisy
        // false-when-no-protection emit so round-trips stay clean.
        if (root.Format.TryGetValue("protection", out var protVal)
            && string.Equals(protVal?.ToString(), "none", StringComparison.OrdinalIgnoreCase))
        {
            root.Format.Remove("protectionEnforced");
            root.Format.Remove("protection");
        }
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (k, v) in root.Format)
        {
            bool include = RootScalarKeys.Contains(k);
            if (!include)
            {
                foreach (var pref in RootPrefixGroups)
                {
                    if (k.StartsWith(pref, StringComparison.OrdinalIgnoreCase))
                    {
                        include = true;
                        break;
                    }
                }
            }
            if (!include) continue;
            if (v == null) continue;
            var s = v switch { bool b => b ? "true" : "false", _ => v.ToString() ?? "" };
            if (s.Length > 0) props[k] = s;
        }
        // docDefaults.font side-effect: the bare TrySetDocDefaults("docdefaults.font", v)
        // case writes ALL four font slots (Ascii/HAnsi/EastAsia/ComplexScript)
        // — convenient for setup, harmful on round-trip. Source documents
        // commonly carry only Ascii/HAnsi (latin) in docDefaults; emitting
        // the bare key on replay would spuriously stamp the same value into
        // eastAsia and complexScript, drifting away from source.
        //
        // Rewrite the bare `docDefaults.font` into the targeted
        // `docDefaults.font.latin` (= Ascii+HAnsi only) so the round-trip
        // doesn't bleed into the other script slots. Per-slot eastAsia /
        // complexScript / hAnsi keys remain untouched and continue to
        // address only their own slot.
        if (props.TryGetValue("docDefaults.font", out var bareFont))
        {
            props.Remove("docDefaults.font");
            props["docDefaults.font.latin"] = bareFont;
        }
        // BUG-R6-05: BlankDocCreator stamps `Times New Roman` into
        // docDefaults RunFonts. Source docs that omit the slot (use theme
        // fonts) round-trip with the blank's TNR baked in. Force an
        // explicit empty `docDefaults.font.latin=` so the Set path clears
        // the blank's TNR back to absent. Same for docGrid.type which the
        // blank sets to `default`.
        if (!props.ContainsKey("docDefaults.font.latin")
            && !props.ContainsKey("docDefaults.font"))
        {
            props["docDefaults.font.latin"] = "";
        }
        if (props.Count == 0) return;
        items.Add(new BatchItem
        {
            Command = "set",
            Path = "/",
            Props = props
        });
    }

    private static void EmitStyles(WordHandler word, List<BatchItem> items)
    {
        // Use query() rather than walking Get("/styles").Children — the
        // positional /styles/style[N] children Get returns are not
        // addressable on the Get side (style paths resolve by id, not by
        // index). Query produces id-based paths and excludes docDefaults.
        var styles = word.Query("style");
        foreach (var stub in styles)
        {
            // CONSISTENCY(slash-in-style-id): style ids/names containing '/'
            // produce paths like /styles/Style/With/Slash that the path
            // parser splits on. Get fails. Fall back to the Query stub —
            // we lose pPr/rPr details but at least the style stub
            // (id/name/type/basedOn) round-trips, instead of dropping the
            // style entirely (BUG BT-3).
            DocumentNode full;
            try { full = word.Get(stub.Path); }
            catch { full = stub; }
            var props = FilterEmittableProps(full.Format);
            // Ensure id is present (Add requires it for /styles target).
            if (!props.ContainsKey("id") && !props.ContainsKey("styleId"))
            {
                if (props.TryGetValue("name", out var n)) props["id"] = n;
                else continue;
            }
            // BUG-R6-03: built-in style ids (Normal / Heading1-9 / Title /
            // …) collide with the blank template's reservations on a
            // fresh batch target. AddStyle is now idempotent for those
            // specific ids (upsert: drop existing + re-add). For non-
            // built-in ids the strict "already exists" check still
            // applies. Emit `add` uniformly so the wire format stays a
            // simple `add`-only stream regardless of style provenance.
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = "/styles",
                Type = "style",
                Props = props
            });
            // BUG-R4-T1: FilterEmittableProps drops the `tabs` scalar (it's a
            // List<Dict>, not stringable). EmitParagraph compensates by
            // emitting per-stop `add tab` rows; EmitStyles must do the same
            // or paragraph-level custom tab stops on a style (Heading TOC
            // leader tabs, etc.) silently disappear on round-trip.
            var styleId = props.TryGetValue("id", out var sid) ? sid
                : props.TryGetValue("styleId", out sid) ? sid : null;
            if (styleId != null && full.Format.TryGetValue("tabs", out var styleTabs))
            {
                EmitTabStops($"/styles/{styleId}", styleTabs, items);
            }
        }
    }

    private sealed class NoteCursor { public int Index; }

    private sealed record ChartSpec(Dictionary<string, object?> Format, IReadOnlyList<DocumentNode> Series);

    private sealed record BodyEmitContext(
        List<string> FootnoteTexts,
        List<string> EndnoteTexts,
        NoteCursor FootnoteCursor,
        NoteCursor EndnoteCursor,
        List<ChartSpec> ChartSpecs,
        NoteCursor ChartCursor,
        Dictionary<string, int>? ParaIdToTargetIdx);

    private static void EmitBody(WordHandler word, List<BatchItem> items,
                                 Dictionary<string, int>? paraIdToTargetIdx = null)
    {
        var bodyNode = word.Get("/body");
        if (bodyNode.Children == null) return;

        // Footnotes/endnotes are referenced by runs (rStyle=FootnoteReference)
        // inside body paragraphs but the run carries no id back to the
        // notes part. We assume notes are listed in document order matching
        // reference order — the typical case since AddFootnote/AddEndnote
        // allocate ids sequentially.
        // Charts: query("chart") returns /chart[N] in document order, which
        // matches the order chart-bearing runs appear in body. Pre-resolve
        // each chart's properties + series children so EmitParagraph can
        // emit a typed `add chart` row when it walks across each ref.
        var charts = word.Query("chart");
        var chartSpecs = charts.Select(c =>
        {
            var full = word.Get(c.Path);
            return new ChartSpec(full.Format, full.Children ?? new List<DocumentNode>());
        }).ToList();

        var ctx = new BodyEmitContext(
            FootnoteTexts: word.Query("footnote").Select(n => n.Text ?? "").ToList(),
            EndnoteTexts: word.Query("endnote").Select(n => n.Text ?? "").ToList(),
            FootnoteCursor: new NoteCursor(),
            EndnoteCursor: new NoteCursor(),
            ChartSpecs: chartSpecs,
            ChartCursor: new NoteCursor(),
            ParaIdToTargetIdx: paraIdToTargetIdx);

        int pIndex = 0, tblIndex = 0;
        foreach (var child in bodyNode.Children)
        {
            switch (child.Type)
            {
                case "paragraph":
                case "p":
                    // BUG-R4-FUZZ-1: display-mode equations surface in
                    // bodyNode.Children as type="paragraph" but the path
                    // resolver addresses them as /body/oMathPara[N], NOT as
                    // /body/p[N]. Incrementing pIndex for them would offset
                    // every subsequent inline-child path (hyperlink/footnote/
                    // run) by +1 per preceding equation, breaking round-trip.
                    // Detect the wrapper via path and route to EmitParagraph
                    // without bumping pIndex — EmitParagraph's equation branch
                    // re-emits the equation as `add /body --type equation`.
                    if (child.Path.Contains("/oMathPara[", StringComparison.OrdinalIgnoreCase))
                    {
                        EmitParagraph(word, child.Path, "/body", pIndex + 1, items, autoPresent: false, ctx);
                    }
                    else
                    {
                        pIndex++;
                        EmitParagraph(word, child.Path, "/body", pIndex, items, autoPresent: false, ctx);
                    }
                    break;
                case "table":
                    tblIndex++;
                    EmitTable(word, child.Path, tblIndex, items, ctx);
                    break;
                case "section":
                case "sectPr":
                    // The body always carries one trailing sectPr that the
                    // blank document already provides; for v0.5 we rely on
                    // that default and skip emitting section properties.
                    // Section emit is a follow-up.
                    break;
                case "sdt":
                    EmitSdt(word, child.Path, items);
                    break;
                default:
                    // Unknown body-level child types — skip for v0.5.
                    break;
            }
        }
    }

    /// <summary>
    /// Emit a paragraph at the target index under <paramref name="parentPath"/>.
    /// When <paramref name="autoPresent"/> is true, the parent already has a
    /// pre-existing paragraph at that index (e.g. an auto-created table cell
    /// paragraph); we issue a `set` instead of a fresh `add` so the existing
    /// paragraph gets reused rather than duplicated.
    /// </summary>
    private static void EmitParagraph(WordHandler word, string sourcePath, string parentPath,
                                      int targetIndex, List<BatchItem> items, bool autoPresent,
                                      BodyEmitContext? ctx = null)
    {
        var pNode = word.Get(sourcePath);

        // Display-mode equations (<m:oMathPara>) surface in EmitBody's
        // bodyNode.Children as type=paragraph, but a direct Get on the
        // path returns type=equation with the LaTeX-ish formula in
        // DocumentNode.Text. EmitParagraph would otherwise emit an empty
        // `add p` and lose the entire formula. Route to typed
        // `add /body --type equation` instead.
        if (pNode.Type == "equation" && parentPath == "/body" && !autoPresent)
        {
            var mode = pNode.Format.TryGetValue("mode", out var m) ? m?.ToString() : "display";
            var eqProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["mode"] = string.IsNullOrEmpty(mode) ? "display" : mode
            };
            if (!string.IsNullOrEmpty(pNode.Text))
                eqProps["formula"] = pNode.Text!;
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = "/body",
                Type = "equation",
                Props = eqProps
            });
            return;
        }

        // Track source paraId -> target index BEFORE any early-return path
        // (section break, TOC, …). Comments anchored on a section-break or
        // TOC paragraph would otherwise miss the mapping and fall back to
        // /body/p[1], silently retargeting the comment.
        if (ctx?.ParaIdToTargetIdx != null && parentPath == "/body" &&
            pNode.Format.TryGetValue("paraId", out var earlyParaId) && earlyParaId != null)
        {
            ctx.ParaIdToTargetIdx[earlyParaId.ToString()!] = targetIndex;
        }

        // Inline section break: a paragraph carrying <w:sectPr> is the
        // OOXML representation of a mid-document section boundary.
        // AddSection on /body produces this same shape, so we emit
        // `add /body --type section` (which creates a fresh break paragraph)
        // rather than emitting a regular `add p`. The companion
        // sectionBreak.* keys map back to AddSection's prop vocabulary.
        if (parentPath == "/body" &&
            pNode.Format.TryGetValue("sectionBreak", out var breakKind) && breakKind != null)
        {
            var sectProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["type"] = breakKind.ToString() ?? "nextPage"
            };
            foreach (var (k, v) in pNode.Format)
            {
                if (!k.StartsWith("sectionBreak.", StringComparison.OrdinalIgnoreCase)) continue;
                if (v == null) continue;
                var keyTail = k["sectionBreak.".Length..];
                var s = v switch { bool b => b ? "true" : "false", _ => v.ToString() ?? "" };
                if (s.Length > 0) sectProps[keyTail] = s;
            }
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = "/body",
                Type = "section",
                Props = sectProps
            });
            return;
        }

        // TOC field-bearing paragraph: a fldChar(begin) + instrText("TOC ...")
        // + fldChar(separate) + placeholder run + fldChar(end) chain. Get
        // exposes only the placeholder text on the parent paragraph, so
        // emitting a regular `add p text=...` would drop the field structure
        // entirely and Word would no longer auto-update the TOC on open.
        // Detect the chain and emit a typed `add /body --type toc` instead;
        // AddToc rebuilds the full fldChar wrapper with the same instruction.
        if (parentPath == "/body" && pNode.Children != null)
        {
            var instrChild = pNode.Children
                .FirstOrDefault(c => c.Type == "instrText"
                    && (c.Format.TryGetValue("instruction", out var iv)
                        && iv?.ToString()?.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase) == true));
            if (instrChild != null)
            {
                var instr = instrChild.Format["instruction"]!.ToString()!;
                var tocProps = ParseTocInstruction(instr);
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = "/body",
                    Type = "toc",
                    Props = tocProps
                });
                return;
            }
        }

        var props = FilterEmittableProps(pNode.Format);
        // When a paragraph carries numId, the abstractNum/num pair is already
        // in /numbering (raw-set wholesale by EmitNumberingRaw). Forwarding
        // numFmt/listStyle/start to AddParagraph triggers ad-hoc
        // numbering-definition creation in WordHandler.Add — Word allocates
        // a fresh numId (1→9, 2→16, …) and the paragraph references the
        // new one, orphaning the original abstract numbering's level rPr
        // (color, bold, custom marker text). Drop those keys so the
        // paragraph just attaches by numId+numLevel to the existing def.
        if (props.ContainsKey("numId"))
        {
            props.Remove("numFmt");
            props.Remove("listStyle");
            props.Remove("start");
        }
        // Collapse non-TOC field chains (fldChar(begin) + instrText(" PAGE ")
        // + fldChar(separate) + display run(s) + fldChar(end)) into a single
        // synthetic "field" entry. Without this collapse, the subsequent
        // `runs` filter sees only the cached display run and emits the field
        // value as static text — PAGE/REF/SEQ/HYPERLINK/NUMPAGES degrade to
        // their evaluated string and stop auto-updating (BUG-R2-05 / R2-1).
        var fieldEntries = CollapseFieldChains(pNode.Children ?? new List<DocumentNode>());
        var runs = fieldEntries
            .Where(c => c.Type == "run" || c.Type == "r" || c.Type == "picture" || c.Type == "field" || c.Type == "ptab")
            .ToList();
        var breaks = (pNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "break")
            .ToList();

        // Single-run / no-run paragraph: collapse run formatting into the
        // paragraph's prop bag (the schema-reflection layer accepts run-level
        // keys on a paragraph and routes them through ApplyRunFormatting).
        // Picture runs need their own typed `add picture` row, so the
        // collapse only applies when the sole run is a regular text run.
        // Break-only paragraphs (e.g. <w:p><w:r><w:br type=page/></w:r></w:p>)
        // also fall out of collapse — they need an explicit `add pagebreak`
        // child after the empty paragraph is created.
        // A run carrying `url` (or `anchor`) was a <w:hyperlink>-wrapped
        // run in source; collapsing it into a paragraph-level prop bag
        // would drop the hyperlink wrapper because `add p` does not
        // consume url/anchor. Force the multi-run path so the run gets
        // re-emitted as `add hyperlink` below.
        bool singleRunIsHyperlink = runs.Count == 1 &&
            (runs[0].Format.ContainsKey("url") || runs[0].Format.ContainsKey("anchor"));
        // BUG-R4-FUZZ-2: when a paragraph's sole run is a footnote/endnote
        // reference (rStyle=FootnoteReference / EndnoteReference), collapsing
        // the run into the paragraph prop bag emits `add p props={rStyle=...}`
        // and drops the typed `add footnote/endnote` row entirely (Add does
        // not consume rStyle on a paragraph; the note text is lost). Force
        // the multi-run path so the dedicated note-emit branch below fires.
        // BUG-R6-6: w14 text effects (textOutline / textFill / w14shadow /
        // w14glow / w14reflection) live on a run but AddParagraph's
        // ApplyRunFormatting fallback has no case for them — collapsing
        // the single run would route the keys to the paragraph prop bag
        // and they'd surface as UNSUPPORTED on replay (effect lost).
        // Force the multi-run path so the effects ride along on `add r`.
        bool singleRunHasW14 = runs.Count == 1 &&
            (runs[0].Format.ContainsKey("w14shadow")
             || runs[0].Format.ContainsKey("textOutline")
             || runs[0].Format.ContainsKey("textFill")
             || runs[0].Format.ContainsKey("w14glow")
             || runs[0].Format.ContainsKey("w14reflection"));
        bool singleRunIsNoteRef = runs.Count == 1 &&
            runs[0].Format.TryGetValue("rStyle", out var srStyle)
            && (string.Equals(srStyle?.ToString(), "FootnoteReference", StringComparison.OrdinalIgnoreCase)
                || string.Equals(srStyle?.ToString(), "EndnoteReference", StringComparison.OrdinalIgnoreCase));
        // BUG-R7-05: a synthetic field run (from CollapseFieldChains) carries
        // `instruction=PAGE` + `text="1"` — collapsing those onto the
        // paragraph emits `set /footer[1]/p[1] instruction=PAGE text=1` which
        // ApplyParagraphLevelProperty doesn't translate into an actual field
        // chain (paragraph just becomes static text "1"). Force the multi-run
        // path so the field run is re-emitted as `add field` and the chain
        // is rebuilt on replay. Header parts hit this same code path; the
        // bug surfaces in footers because header documents in earlier rounds
        // happened to have multiple runs that already forced the multi-run
        // branch.
        bool singleRunIsField = runs.Count == 1 && runs[0].Type == "field";
        bool collapseSingleRun = runs.Count <= 1 &&
            !(runs.Count == 1 && runs[0].Type == "picture") &&
            !(runs.Count == 1 && runs[0].Type == "ptab") &&
            !singleRunIsHyperlink &&
            !singleRunIsNoteRef &&
            !singleRunHasW14 &&
            !singleRunIsField &&
            breaks.Count == 0;
        // Pull paragraph-level tab stops out for per-stop `add tab` emit
        // (FilterEmittableProps already drops the `tabs` scalar).
        pNode.Format.TryGetValue("tabs", out var pTabs);

        if (collapseSingleRun)
        {
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

            if (autoPresent)
            {
                // Replace the auto-created paragraph in place — only push the
                // set when there is something to apply, otherwise the empty
                // skeleton is already correct.
                if (props.Count > 0)
                {
                    items.Add(new BatchItem
                    {
                        Command = "set",
                        Path = $"{parentPath}/p[{targetIndex}]",
                        Props = props
                    });
                }
            }
            else
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = parentPath,
                    Type = "p",
                    Props = props.Count > 0 ? props : null
                });
            }
            EmitTabStops($"{parentPath}/p[{targetIndex}]", pTabs, items);
            return;
        }

        // Multi-run paragraph: emit/set the paragraph empty first, then add
        // each run as an explicit child.
        if (autoPresent)
        {
            if (props.Count > 0)
            {
                items.Add(new BatchItem
                {
                    Command = "set",
                    Path = $"{parentPath}/p[{targetIndex}]",
                    Props = props
                });
            }
        }
        else
        {
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = parentPath,
                Type = "p",
                Props = props.Count > 0 ? props : null
            });
        }

        var paraTargetPath = $"{parentPath}/p[{targetIndex}]";
        EmitTabStops(paraTargetPath, pTabs, items);

        // Emit any break runs (page/column/textWrapping/line) the paragraph
        // carries. Without this, a break-only paragraph (the OOXML idiom
        // for "page break here") collapsed to an empty paragraph and
        // subsequent content shifted up a page.
        foreach (var br in breaks)
        {
            var breakType = br.Format.TryGetValue("breakType", out var bt) ? bt?.ToString() : "page";
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "pagebreak",
                Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["type"] = string.IsNullOrEmpty(breakType) ? "page" : breakType
                }
            });
        }

        foreach (var run in runs)
        {
            // Positional tab — Navigation surfaces ptab as its own run type
            // with align/relativeTo/leader on Format. Without an explicit
            // emit branch the runs filter would drop it (BUG-R6-4) and the
            // round-trip would silently lose right-align/header-style tabs.
            if (run.Type == "ptab")
            {
                var ptabProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (run.Format.TryGetValue("align", out var pAlign) && pAlign != null)
                    ptabProps["alignment"] = pAlign.ToString() ?? "";
                if (run.Format.TryGetValue("relativeTo", out var pRel) && pRel != null)
                    ptabProps["relativeTo"] = pRel.ToString() ?? "";
                if (run.Format.TryGetValue("leader", out var pLead) && pLead != null)
                    ptabProps["leader"] = pLead.ToString() ?? "";
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = paraTargetPath,
                    Type = "ptab",
                    Props = ptabProps.Count > 0 ? ptabProps : null
                });
                continue;
            }

            // Synthetic field entry from CollapseFieldChains. Format carries
            // `instruction` (the raw fldSimple/instrText string) and Text holds
            // the cached display value. AddField parses the instruction code
            // and rebuilds the fldChar chain on replay.
            if (run.Type == "field")
            {
                var instr = run.Format.TryGetValue("instruction", out var iv)
                    ? iv?.ToString() ?? "" : "";
                var fieldProps = BuildFieldAddProps(instr, run.Text ?? "");
                if (fieldProps != null)
                {
                    items.Add(new BatchItem
                    {
                        Command = "add",
                        Parent = paraTargetPath,
                        Type = "field",
                        Props = fieldProps
                    });
                }
                else if (!string.IsNullOrEmpty(run.Text))
                {
                    // Unparseable instruction — fall back to plain text so the
                    // paragraph still renders the cached value rather than going
                    // empty.
                    items.Add(new BatchItem
                    {
                        Command = "add",
                        Parent = paraTargetPath,
                        Type = "r",
                        Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["text"] = run.Text! }
                    });
                }
                continue;
            }

            // Drawing-bearing runs surface as type=="picture" regardless of
            // whether the Drawing wraps an image (Blip) or a chart
            // (c:chart). Try the image path first; if there's no embedded
            // image part the run is a chart anchor — pull the next
            // pre-resolved ChartSpec and emit a typed `add chart` row.
            if (run.Type == "picture")
            {
                var binary = word.GetImageBinary(run.Path);
                if (binary.HasValue)
                {
                    var (bytes, contentType) = binary.Value;
                    var dataUri = $"data:{contentType};base64,{Convert.ToBase64String(bytes)}";
                    var picProps = FilterEmittableProps(run.Format);
                    picProps.Remove("id");
                    picProps.Remove("contentType");
                    picProps.Remove("fileSize");
                    picProps["src"] = dataUri;
                    items.Add(new BatchItem
                    {
                        Command = "add",
                        Parent = paraTargetPath,
                        Type = "picture",
                        Props = picProps
                    });
                    continue;
                }

                // Only consume a ChartSpec if the run is genuinely a chart.
                // Picture-typed runs that aren't images can also be background
                // images, OLE objects, SmartArt, watermark anchors, etc. —
                // falling through unconditionally to chart consumption would
                // misalign chart positions for every subsequent chart in the
                // document (e.g. a Background anchor at p[1] would steal the
                // chart spec belonging to a real chart further down).
                if (ctx != null && word.IsChartRun(run.Path)
                    && ctx.ChartCursor.Index < ctx.ChartSpecs.Count)
                {
                    var spec = ctx.ChartSpecs[ctx.ChartCursor.Index];
                    ctx.ChartCursor.Index++;
                    var chartProps = BuildChartProps(spec);
                    items.Add(new BatchItem
                    {
                        Command = "add",
                        Parent = paraTargetPath,
                        Type = "chart",
                        Props = chartProps
                    });
                    continue;
                }
                // Drawing without image part and not a chart — most likely a
                // wps shape (background rectangle, watermark anchor) drawn
                // with prstGeom + solidFill. No typed Add path exists yet,
                // but the XML is self-contained (no rId/embed back-references)
                // so round-trip via raw-set append is safe. Targets the
                // already-created paragraph by xpath positional index.
                // Caveats: drawings with embedded image references (a:blipFill
                // with r:embed) would also land here and silently lose their
                // image part — for those we'd need rId remapping. Acceptable
                // v0.5 lossy mode: log nothing, round-trip survives for the
                // common decorative-shape case.
                var rawXml = word.GetElementXml(run.Path);
                if (!string.IsNullOrEmpty(rawXml) &&
                    parentPath == "/body" &&
                    !rawXml.Contains("r:embed") && !rawXml.Contains("r:id"))
                {
                    items.Add(new BatchItem
                    {
                        Command = "raw-set",
                        Part = "/document",
                        Xpath = $"/w:document/w:body/w:p[{targetIndex}]",
                        Action = "append",
                        Xml = rawXml
                    });
                }
                continue;
            }

            // Detect footnote/endnote reference runs. The OOXML model marks
            // them with a w:rStyle = FootnoteReference / EndnoteReference;
            // the run itself carries no visible text. Emit them as a
            // typed footnote/endnote add anchored on the host paragraph and
            // pull the body text from the pre-resolved ordered list — see
            // BodyEmitContext for the document-order assumption.
            var rStyle = run.Format.TryGetValue("rStyle", out var rs) ? rs?.ToString() : null;
            if (ctx != null && rStyle == "FootnoteReference")
            {
                var noteText = ctx.FootnoteCursor.Index < ctx.FootnoteTexts.Count
                    ? ctx.FootnoteTexts[ctx.FootnoteCursor.Index]
                    : "";
                ctx.FootnoteCursor.Index++;
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = paraTargetPath,
                    Type = "footnote",
                    Props = new() { ["text"] = noteText }
                });
                continue;
            }
            if (ctx != null && rStyle == "EndnoteReference")
            {
                var noteText = ctx.EndnoteCursor.Index < ctx.EndnoteTexts.Count
                    ? ctx.EndnoteTexts[ctx.EndnoteCursor.Index]
                    : "";
                ctx.EndnoteCursor.Index++;
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = paraTargetPath,
                    Type = "endnote",
                    Props = new() { ["text"] = noteText }
                });
                continue;
            }

            var rProps = FilterEmittableProps(run.Format);
            if (!string.IsNullOrEmpty(run.Text))
                rProps["text"] = run.Text!;

            // Hyperlink-wrapped run: Get flattens a <w:hyperlink>'s child run
            // into a regular run-typed node, but copies the hyperlink's
            // r:id-resolved URL onto the run via Format["url"]. AddRun does
            // not consume `url` — emitting type="r" would silently drop the
            // hyperlink wrapper. Re-emit as a typed `add hyperlink` so the
            // <w:hyperlink>+rel-relationship round-trip rebuilds correctly.
            // CONSISTENCY(docx-hyperlink-canonical-url): canonical key is
            // `url` on both Get readback and Add input.
            if (rProps.ContainsKey("url") || rProps.ContainsKey("anchor"))
            {
                // AddHyperlink writes its own color/underline defaults from
                // theme; drop the inferred `color: hyperlink` /
                // `underline: single` Get echoes back so we don't override
                // those defaults with stringly-typed values that the
                // AddHyperlink color path doesn't recognize.
                if (rProps.TryGetValue("color", out var hlColor)
                    && string.Equals(hlColor, "hyperlink", StringComparison.OrdinalIgnoreCase))
                    rProps.Remove("color");
                if (rProps.TryGetValue("underline", out var hlUl)
                    && string.Equals(hlUl, "single", StringComparison.OrdinalIgnoreCase))
                    rProps.Remove("underline");
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = paraTargetPath,
                    Type = "hyperlink",
                    Props = rProps,
                });
                continue;
            }
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "r",
                Props = rProps.Count > 0 ? rProps : null
            });
        }
    }

    private static void EmitTable(WordHandler word, string sourcePath, int targetIndex,
                                  List<BatchItem> items, BodyEmitContext? ctx = null,
                                  string? parentTablePath = null)
    {
        var tableNode = word.Get(sourcePath);
        var rows = (tableNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "row")
            .ToList();
        if (rows.Count == 0) return;

        // Column count must cover the widest row including colspan effects.
        // Format["cols"] reflects gridCol; per-row effective width is
        // sum(colspan or 1) over each cell. Take the max so a first row
        // with merged cells (visible cell count < grid width) doesn't
        // truncate the table shape and break later `set tc[N]` rows.
        var rowEffectiveWidths = new List<int>(rows.Count);
        var rowCellNodes = new List<List<DocumentNode>>(rows.Count);
        var rowNodes = new List<DocumentNode>(rows.Count);
        foreach (var rowChild in rows)
        {
            var rowNode = word.Get(rowChild.Path);
            rowNodes.Add(rowNode);
            var cells = (rowNode.Children ?? new List<DocumentNode>())
                .Where(c => c.Type == "cell")
                .ToList();
            rowCellNodes.Add(cells);
            int width = 0;
            foreach (var cell in cells)
            {
                int span = 1;
                if (cell.Format.TryGetValue("colspan", out var sp) &&
                    int.TryParse(sp?.ToString(), out var n) && n > 0)
                {
                    span = n;
                }
                width += span;
            }
            rowEffectiveWidths.Add(width);
        }
        int colsFromRows = rowEffectiveWidths.Count > 0 ? rowEffectiveWidths.Max() : 0;
        int colsFromGrid = 0;
        if (tableNode.Format.TryGetValue("cols", out var gridColObj) &&
            int.TryParse(gridColObj?.ToString(), out var gridCols))
        {
            colsFromGrid = gridCols;
        }
        int cols = Math.Max(colsFromGrid, colsFromRows);
        if (cols == 0) return;

        var tableProps = FilterEmittableProps(tableNode.Format);
        tableProps["rows"] = rows.Count.ToString();
        tableProps["cols"] = cols.ToString();
        // Nested tables sit inside a parent table cell; AddTable accepts
        // /body/tbl[N]/tr[M]/tc[K] as a parent. Outer-level tables target
        // /body. parentTablePath, when set, is a cell target path
        // (/body/tbl[X]/tr[Y]/tc[Z]) that we emit nested tables under.
        var tableParentPath = parentTablePath ?? "/body";
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = tableParentPath,
            Type = "table",
            Props = tableProps
        });

        // For nested tables, the target path is parent_cell/tbl[1] (first
        // table in the cell). For outer tables, it's /body/tbl[N].
        var tablePath = parentTablePath != null
            ? $"{parentTablePath}/tbl[1]"
            : $"/body/tbl[{targetIndex}]";
        for (int r = 0; r < rows.Count; r++)
        {
            // Emit row-level properties (header / height / height.rule) as a
            // `set` on the row path — `add table` only seeds rows, it doesn't
            // surface per-row props (BUG-R6-2). Without this, `dump→batch`
            // silently strips repeating-header rows and explicit row heights.
            var rowNode = rowNodes[r];
            var rowProps = ExtractRowOnlyProps(rowNode.Format);
            if (rowProps.Count > 0)
            {
                items.Add(new BatchItem
                {
                    Command = "set",
                    Path = $"{tablePath}/tr[{r + 1}]",
                    Props = rowProps
                });
            }
            var cells = rowCellNodes[r];
            for (int c = 0; c < cells.Count; c++)
            {
                var cellNode = word.Get(cells[c].Path);
                var cellTargetPath = $"{tablePath}/tr[{r + 1}]/tc[{c + 1}]";

                // Cell-level tcPr properties (fill, valign, width, borders,
                // padding, colspan, …) are surfaced on cellNode.Format but
                // were previously dropped — only the inner paragraph was
                // emitted. Push them via a `set` on the cell path before
                // the paragraph emits so cell shading / merges / widths
                // round-trip. Skip keys that EmitParagraph will re-apply
                // to the first paragraph (align/direction/run leak-throughs)
                // to avoid double-application.
                var cellProps = ExtractCellOnlyProps(cellNode.Format);
                if (cellProps.Count > 0)
                {
                    items.Add(new BatchItem
                    {
                        Command = "set",
                        Path = cellTargetPath,
                        Props = cellProps
                    });
                }

                // Each cell carries auto-generated paragraphs (Add table seeds
                // one empty paragraph per cell). Update the first one in place
                // and append further paragraphs as fresh adds. Nested tables
                // and paragraphs are emitted in document order so footnote/
                // chart cursors (carried in ctx) advance correctly through
                // the table cell content. Without ctx threading, body-level
                // footnote/chart references after a table would resolve
                // against the wrong note text.
                var cellChildren = cellNode.Children ?? new List<DocumentNode>();
                int cellParaIdx = 0;
                int nestedTblIdx = 0;
                bool firstParaSeen = false;
                foreach (var cc in cellChildren)
                {
                    if (cc.Type == "paragraph" || cc.Type == "p")
                    {
                        cellParaIdx++;
                        EmitParagraph(word, cc.Path, cellTargetPath, cellParaIdx, items,
                                      autoPresent: !firstParaSeen, ctx);
                        firstParaSeen = true;
                    }
                    else if (cc.Type == "table")
                    {
                        nestedTblIdx++;
                        EmitTable(word, cc.Path, nestedTblIdx, items, ctx,
                                  parentTablePath: cellTargetPath);
                    }
                }
            }
        }
    }

    // Collapse OOXML complex field chains (fldChar(begin) + instrText + …
    // + fldChar(end)) into a single synthetic "field" DocumentNode with
    // Format["instruction"] (raw code) and Text (cached display value).
    // Non-field children pass through untouched in original order. The TOC
    // chain is handled by the dedicated EmitParagraph branch above and never
    // reaches this collapsing step (early-return in that branch).
    private static List<DocumentNode> CollapseFieldChains(List<DocumentNode> children)
    {
        var result = new List<DocumentNode>();
        for (int i = 0; i < children.Count; i++)
        {
            var c = children[i];
            bool isBegin = c.Type == "fieldChar"
                && c.Format.TryGetValue("fieldCharType", out var fct)
                && string.Equals(fct?.ToString(), "begin", StringComparison.OrdinalIgnoreCase);
            if (!isBegin)
            {
                result.Add(c);
                continue;
            }

            // Walk forward to find instruction text and end marker.
            string instruction = "";
            string display = "";
            int end = -1;
            for (int j = i + 1; j < children.Count; j++)
            {
                var k = children[j];
                if (k.Type == "instrText")
                {
                    if (k.Format.TryGetValue("instruction", out var iv) && iv != null)
                        instruction += iv.ToString();
                    else if (!string.IsNullOrEmpty(k.Text))
                        instruction += k.Text;
                }
                else if (k.Type == "fieldChar"
                    && k.Format.TryGetValue("fieldCharType", out var ft)
                    && string.Equals(ft?.ToString(), "end", StringComparison.OrdinalIgnoreCase))
                {
                    end = j;
                    break;
                }
                else if (k.Type == "run" || k.Type == "r")
                {
                    // Cached display segments after fldChar(separate). Concatenate
                    // their text — formatting on the display run is dropped (the
                    // field renders fresh on replay).
                    if (!string.IsNullOrEmpty(k.Text)) display += k.Text;
                }
            }
            if (end < 0)
            {
                // Malformed (no end marker) — fall back to passing through.
                result.Add(c);
                continue;
            }
            var synth = new DocumentNode
            {
                Path = c.Path,
                Type = "field",
                Text = display,
                Format = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                {
                    ["instruction"] = instruction.Trim()
                }
            };
            result.Add(synth);
            i = end;
        }
        return result;
    }

    // Build the prop bag AddField consumes from a parsed field instruction.
    // Returns null when the instruction is empty or its first token is not a
    // known field code; the caller falls back to a plain-text run for the
    // cached display value so the paragraph still renders.
    private static Dictionary<string, string>? BuildFieldAddProps(string instruction, string display)
    {
        if (string.IsNullOrWhiteSpace(instruction)) return null;
        var trimmed = instruction.Trim();
        // First whitespace-separated token is the field code.
        var firstSpace = trimmed.IndexOfAny(new[] { ' ', '\t' });
        var code = (firstSpace < 0 ? trimmed : trimmed[..firstSpace]).ToUpperInvariant();
        var rest = firstSpace < 0 ? "" : trimmed[(firstSpace + 1)..].Trim();

        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["fieldType"] = code
        };
        switch (code)
        {
            case "PAGE":
            case "NUMPAGES":
            case "AUTHOR":
            case "TITLE":
            case "SUBJECT":
            case "FILENAME":
            case "SECTION":
            case "SECTIONPAGES":
                break;
            case "DATE":
            case "TIME":
            case "CREATEDATE":
            case "SAVEDATE":
            case "PRINTDATE":
            {
                // Preserve the `\@ "MMMM d, yyyy"` format switch so dump
                // round-trips Word's locale-formatted date fields. Without
                // this, BuildFieldAddProps dropped `rest` and replay
                // produced a bare DATE field rendered in the default
                // locale (BUG-R6-3). AddField consumes the value via
                // --prop format=…
                var fmtMatch = System.Text.RegularExpressions.Regex.Match(
                    rest ?? "", "\\\\@\\s+\"([^\"]+)\"");
                if (fmtMatch.Success)
                    props["format"] = fmtMatch.Groups[1].Value;
                break;
            }
            case "REF":
            case "PAGEREF":
            case "NOTEREF":
            {
                // First arg is the bookmark name (may be quoted).
                var name = ExtractFirstArg(rest);
                if (string.IsNullOrEmpty(name)) return null;
                props["bookmarkName"] = name;
                break;
            }
            case "SEQ":
            {
                var ident = ExtractFirstArg(rest);
                if (string.IsNullOrEmpty(ident)) return null;
                props["identifier"] = ident;
                break;
            }
            case "MERGEFIELD":
            {
                var name = ExtractFirstArg(rest);
                if (string.IsNullOrEmpty(name)) return null;
                props["fieldName"] = name;
                break;
            }
            case "HYPERLINK":
            {
                // Either `\l "anchor"` or a URL as the first arg.
                var anchorMatch = System.Text.RegularExpressions.Regex.Match(rest, "\\\\l\\s+\"([^\"]+)\"");
                if (anchorMatch.Success) props["anchor"] = anchorMatch.Groups[1].Value;
                else
                {
                    var url = ExtractFirstArg(rest);
                    if (string.IsNullOrEmpty(url)) return null;
                    props["url"] = url;
                }
                break;
            }
            default:
                // Unknown field code — let the generic field add try.
                break;
        }
        if (!string.IsNullOrEmpty(display))
            props["text"] = display;
        return props;
    }

    private static string ExtractFirstArg(string s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        var t = s.TrimStart();
        if (t.StartsWith('"'))
        {
            var end = t.IndexOf('"', 1);
            return end > 0 ? t[1..end] : "";
        }
        var spc = t.IndexOfAny(new[] { ' ', '\t' });
        return spc < 0 ? t : t[..spc];
    }

    // Parse a TOC field instruction (` TOC \o "1-3" \h \u \z `) into the
    // prop bag AddToc accepts. AddToc emits the canonical instruction so
    // round-tripping the parsed props back through it lands at the same
    // OOXML even when the source instruction had extra whitespace or
    // switch ordering.
    private static Dictionary<string, string> ParseTocInstruction(string instruction)
    {
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var lvl = System.Text.RegularExpressions.Regex.Match(instruction, "\\\\o\\s+\"([^\"]+)\"");
        if (lvl.Success) props["levels"] = lvl.Groups[1].Value;
        // \h = hyperlinks (default true on AddToc, but emit explicitly for clarity)
        props["hyperlinks"] = System.Text.RegularExpressions.Regex.IsMatch(instruction, "\\\\h\\b")
            ? "true" : "false";
        // \z suppresses page numbers; absence means pageNumbers=true
        props["pageNumbers"] = System.Text.RegularExpressions.Regex.IsMatch(instruction, "\\\\z\\b")
            ? "false" : "true";
        // BUG-R5-03: \t = custom-style→level mapping ("Style;level,..."),
        // \b = bookmark scope. Capture the quoted argument so AddToc can
        // round-trip them; otherwise custom TOC switches were silently
        // dropped on dump.
        var ct = System.Text.RegularExpressions.Regex.Match(instruction, "\\\\t\\s+\"([^\"]+)\"");
        if (ct.Success) props["customStyles"] = ct.Groups[1].Value;
        var cb = System.Text.RegularExpressions.Regex.Match(instruction, "\\\\b\\s+\"([^\"]+)\"");
        if (cb.Success) props["bookmark"] = cb.Groups[1].Value;
        return props;
    }

    // Cell Format includes both true tcPr keys and "leaked" keys read from
    // the first inner paragraph/run (align, direction, font, size, bold, …).
    // EmitParagraph re-emits those for the first paragraph, so emitting them
    // here too would double-apply. Whitelist genuine cell-level keys only.
    private static readonly HashSet<string> CellOnlyKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "fill", "width", "valign", "vmerge", "colspan", "nowrap", "textDirection",
    };

    private static Dictionary<string, string> ExtractCellOnlyProps(Dictionary<string, object?> raw)
    {
        var filtered = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var (key, val) in raw)
        {
            if (CellOnlyKeys.Contains(key) ||
                key.StartsWith("border.", StringComparison.OrdinalIgnoreCase) ||
                key.StartsWith("padding.", StringComparison.OrdinalIgnoreCase))
            {
                filtered[key] = val;
            }
        }
        return FilterEmittableProps(filtered);
    }

    // Row-level keys surfaced by Navigation.ReadRowProps. Used by EmitTable
    // so dump→batch round-trips header rows / heights / cantSplit. Cell
    // children are emitted separately via ExtractCellOnlyProps.
    private static readonly HashSet<string> RowOnlyKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "header", "height", "cantSplit",
    };

    private static Dictionary<string, string> ExtractRowOnlyProps(Dictionary<string, object?> raw)
    {
        var filtered = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        bool heightExact = false;
        if (raw.TryGetValue("height.rule", out var ruleObj) &&
            string.Equals(ruleObj?.ToString(), "exact", StringComparison.OrdinalIgnoreCase))
        {
            heightExact = true;
        }
        foreach (var (key, val) in raw)
        {
            if (!RowOnlyKeys.Contains(key)) continue;
            // height + height.rule=exact → SetElementTableRow expects key
            // `height.exact`. Translate so dump output applies cleanly.
            if (heightExact && string.Equals(key, "height", StringComparison.OrdinalIgnoreCase))
            {
                filtered["height.exact"] = val;
            }
            else
            {
                filtered[key] = val;
            }
        }
        return FilterEmittableProps(filtered);
    }

    private static Dictionary<string, string> BuildChartProps(ChartSpec spec)
    {
        // AddChart ingests data series via a single `data="Name1:v1,v2;Name2:v1,v2"`
        // string. Reconstruct that string from the series children Get
        // exposes; categories come from the chart's own Format key.
        var props = FilterEmittableProps(spec.Format);
        // Strip Get-only / SDK-managed keys that AddChart neither expects
        // nor accepts.
        props.Remove("id");
        props.Remove("seriesCount");

        // Build data="Name:v1,v2;..." from series children.
        var seriesParts = new List<string>();
        foreach (var s in spec.Series)
        {
            if (s.Type != "series") continue;
            if (!s.Format.TryGetValue("name", out var nObj) || nObj == null) continue;
            if (!s.Format.TryGetValue("values", out var vObj) || vObj == null) continue;
            var name = nObj.ToString() ?? "";
            var vals = vObj.ToString() ?? "";
            if (name.Length == 0 || vals.Length == 0) continue;
            seriesParts.Add($"{name}:{vals}");
        }
        if (seriesParts.Count > 0)
        {
            props["data"] = string.Join(";", seriesParts);
        }
        return props;
    }

    // Format keys that must NOT be emitted: derived (computed by Get, not
    // user-set), unstable (regenerate on save), or coordinate-system
    // (paths that only make sense in the source document).
    private static readonly HashSet<string> SkipKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "basedOn.path",
        "paraId", "textId", "rsidR", "rsidRDefault", "rsidRPr", "rsidP", "rsidTr",
        // Paragraph Get emits `style`, `styleId`, and `styleName` — all three
        // carry the same value (style id, repeated). AddParagraph only
        // consumes `style`; emitting the other two would either re-process
        // the same value (no-op) or, if Add ever grows divergent semantics
        // for them, cause double-application. Drop the aliases so the
        // dump bag stays minimal.
        "styleId", "styleName",
        // CONSISTENCY(line-rule-inferred): style/paragraph Get emits `lineRule`
        // alongside `lineSpacing`, but lineSpacing already encodes the rule
        // (e.g. "1.5x" => Auto, "18pt" => Exact via SpacingConverter). Set has
        // no `lineRule` case, so emitting it produces "UNSUPPORTED" warnings
        // (104+ on a typical real-world doc). Drop it from dump.
        "lineRule",
    };

    // Serialize the tabs list (List<Dictionary<string,object?>>) to the
    // semicolon-separated form Add/Set's tab parser accepts:
    //   "pos=2160 val=left leader=dot;pos=4320 val=center"
    // ApplyTabsString in WordHandler.Helpers parses both space- and
    // comma-delimited fields per stop. Without this projection, the raw
    // .NET List<Dict> rendered as "System.Collections.Generic.List`1[…]"
    // and tab stops were silently lost on round-trip (BUG-R2-01).
    private static string SerializeTabsList(IEnumerable<Dictionary<string, object?>> tabs)
    {
        var stops = new List<string>();
        foreach (var t in tabs)
        {
            var parts = new List<string>();
            if (t.TryGetValue("pos", out var p) && p != null) parts.Add($"pos={p}");
            if (t.TryGetValue("val", out var v) && v != null) parts.Add($"val={v}");
            if (t.TryGetValue("leader", out var l) && l != null) parts.Add($"leader={l}");
            if (parts.Count > 0) stops.Add(string.Join(" ", parts));
        }
        return string.Join(";", stops);
    }

    private static Dictionary<string, string> FilterEmittableProps(Dictionary<string, object?> raw)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // CONSISTENCY(border-fold): Get emits `pbdr.bottom: single`,
        // `pbdr.bottom.sz: 6`, `pbdr.bottom.color: #FF0000`, `pbdr.bottom.space: 1`
        // as separate keys (mirrors `border.*` on Excel). Set accepts a single
        // colon-encoded value `pbdr.bottom=single:6:#FF0000:1`. Without folding,
        // the 2-segment key applies an empty-style border and the 3-segment
        // subkeys hit unsupported (BUG BT-6: Title/Intense Quote lose bottom
        // border on round-trip). Fold the 4 keys into one before validation.
        var pbdrFold = new Dictionary<string, (string? style, string? sz, string? color, string? space)>(
            StringComparer.OrdinalIgnoreCase);
        foreach (var (key, val) in raw)
        {
            if (val == null) continue;
            if (!key.StartsWith("pbdr.", StringComparison.OrdinalIgnoreCase)) continue;
            var parts = key.Split('.');
            if (parts.Length < 2) continue;
            var side = $"{parts[0]}.{parts[1]}"; // pbdr.bottom
            pbdrFold.TryGetValue(side, out var cur);
            var sval = val.ToString() ?? "";
            if (parts.Length == 2) cur.style = sval;
            else if (parts.Length == 3)
            {
                switch (parts[2].ToLowerInvariant())
                {
                    case "sz": cur.sz = sval; break;
                    case "color": cur.color = sval; break;
                    case "space": cur.space = sval; break;
                }
            }
            pbdrFold[side] = cur;
        }

        // BUG-R7-04: same fold for table `border.*` keys. Get emits
        // `border.top: single`, `border.top.sz: 12`, `border.top.color: #000000`
        // separately; Set accepts only the colon-encoded form
        // `border.top=single;12;#000000;1`. Without folding, dump strips the
        // 3-segment subkeys (see the explicit "drop them here" comment below)
        // and round-trip silently downgrades real borders to default thin
        // single. Fold sz/color/space into the 2-segment key.
        var borderFold = new Dictionary<string, (string? style, string? sz, string? color, string? space)>(
            StringComparer.OrdinalIgnoreCase);
        foreach (var (key, val) in raw)
        {
            if (val == null) continue;
            if (!key.StartsWith("border.", StringComparison.OrdinalIgnoreCase)) continue;
            var parts = key.Split('.');
            if (parts.Length < 2) continue;
            var side = $"{parts[0]}.{parts[1]}"; // border.top
            borderFold.TryGetValue(side, out var cur);
            var sval = val.ToString() ?? "";
            if (parts.Length == 2) cur.style = sval;
            else if (parts.Length == 3)
            {
                switch (parts[2].ToLowerInvariant())
                {
                    case "sz": cur.sz = sval; break;
                    case "color": cur.color = sval; break;
                    case "space": cur.space = sval; break;
                }
            }
            borderFold[side] = cur;
        }

        foreach (var (key, val) in raw)
        {
            if (SkipKeys.Contains(key)) continue;
            if (key.StartsWith("effective.", StringComparison.OrdinalIgnoreCase)) continue;
            if (key.EndsWith(".cs.source", StringComparison.OrdinalIgnoreCase)) continue;

            // pbdr fold: skip subkeys, rewrite the bare side key into colon form.
            if (key.StartsWith("pbdr.", StringComparison.OrdinalIgnoreCase))
            {
                var parts = key.Split('.');
                if (parts.Length >= 3) continue; // subkey already folded
                var side = $"{parts[0]}.{parts[1]}";
                if (pbdrFold.TryGetValue(side, out var folded) && folded.style != null)
                {
                    // ParseBorderValue format: STYLE[;SIZE[;COLOR[;SPACE]]] — empties
                    // for missing intermediates so positional parts stay aligned.
                    var sz = folded.sz ?? "";
                    var col = folded.color ?? "";
                    var sp = folded.space ?? "";
                    var v = folded.style!;
                    if (folded.sz != null || folded.color != null || folded.space != null)
                        v += ";" + sz;
                    if (folded.color != null || folded.space != null)
                        v += ";" + col;
                    if (folded.space != null)
                        v += ";" + sp;
                    result[key] = v;
                }
                continue;
            }

            // BUG-R7-04: fold border.* like pbdr.*. Skip the 3-segment subkeys
            // (folded into the 2-segment side key below) and rewrite the bare
            // side key into the colon-encoded form Set's ParseBorderValue
            // expects.
            if (key.StartsWith("border.", StringComparison.OrdinalIgnoreCase))
            {
                var bparts = key.Split('.');
                if (bparts.Length >= 3) continue; // subkey already folded
                var bside = $"{bparts[0]}.{bparts[1]}";
                if (borderFold.TryGetValue(bside, out var folded) && folded.style != null)
                {
                    var sz = folded.sz ?? "";
                    var col = folded.color ?? "";
                    var sp = folded.space ?? "";
                    var v = folded.style!;
                    if (folded.sz != null || folded.color != null || folded.space != null)
                        v += ";" + sz;
                    if (folded.color != null || folded.space != null)
                        v += ";" + col;
                    if (folded.space != null)
                        v += ";" + sp;
                    result[key] = v;
                }
                continue;
            }

            // tabs is a List<Dict>, not a flat scalar. Both Add and Set ingest
            // tab stops via the dedicated `add ... --type tab` command (one
            // row per stop), not as a paragraph/style scalar prop. Skipping
            // here avoids serializing the .NET list type name into the prop
            // string (BUG-R2-01); paragraph emitters layer per-stop add rows
            // separately.
            if (string.Equals(key, "tabs", StringComparison.OrdinalIgnoreCase)) continue;

            if (val == null) continue;
            string s = val switch
            {
                bool b => b ? "true" : "false",
                _ => val.ToString() ?? ""
            };
            if (s.Length > 0) result[key] = s;
        }
        return result;
    }

    // Layer per-stop `add tab` rows under a parent path that already has the
    // host paragraph/style created. tabs is the flat List<Dict> Get exposes.
    private static void EmitTabStops(string parentPath, object? tabsVal, List<BatchItem> items)
    {
        if (tabsVal is not IEnumerable<Dictionary<string, object?>> list) return;
        foreach (var t in list)
        {
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (t.TryGetValue("pos", out var p) && p != null) props["pos"] = p.ToString() ?? "";
            if (t.TryGetValue("val", out var v) && v != null) props["val"] = v.ToString() ?? "";
            if (t.TryGetValue("leader", out var l) && l != null) props["leader"] = l.ToString() ?? "";
            if (props.Count == 0 || !props.ContainsKey("pos")) continue;
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = parentPath,
                Type = "tab",
                Props = props
            });
        }
    }
}
