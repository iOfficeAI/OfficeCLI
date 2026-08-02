// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // PERF(nav-child-cache): keys whose Set only writes run/paragraph/cell/row/
    // table PROPERTIES (rPr / pPr / tcPr / trPr / tblPr) and therefore never
    // adds or removes a <w:r> / <w:tr> / <w:tc> element — the only structure the
    // run/row/cell nav caches index. A Set restricted to these keys leaves those
    // caches valid, so a per-run/per-cell attribute-set batch keeps hitting them
    // (arming the clear guard there would rebuild O(R) per set → O(R²)).
    //
    // JUDGMENT (why this is a safe-list, not a structural block-list): the
    // criterion "does this key add/remove an r/tr/tc element?" is provable per
    // key. `text` fails it (replaces all runs of a paragraph/cell). Anything not
    // provably attribute-only — a new key, a field/hyperlink/equation rewrite,
    // revision accept/reject (deletes inserted runs) — is absent here and thus
    // clears by default. Missing a genuinely-safe key only over-clears (perf);
    // wrongly listing an unsafe key would serve a detached element (silent
    // wrong-content bug), so the list stays conservative.
    private static readonly HashSet<string> NavCacheSafeSetKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        // run rPr
        "bold", "italic", "underline", "strike", "doublestrike", "color", "size",
        "fontsize", "font", "highlight", "caps", "smallcaps", "vanish", "hidden",
        "subscript", "superscript", "vertalign", "position", "kerning", "emboss",
        "outline", "shadow", "imprint", "rtl", "lang", "charscale", "spacing",
        "letterspacing",
        // paragraph pPr
        "alignment", "align", "halign", "textalignment", "spacebefore", "spaceafter",
        "linespacing", "linerule", "keepnext", "keeplines", "widowcontrol",
        "outlinelevel", "contextualspacing", "numid", "numlevel", "bidi",
        // cell tcPr
        "valign", "verticalalignment", "width", "nowrap", "textdirection",
        "gridspan", "vmerge", "hidemark", "fit",
        // row trPr
        "height", "rowheight", "cantsplit", "tblheader",
    };

    // A Set clears the nav child caches unless EVERY property key is provably
    // attribute-only (in NavCacheSafeSetKeys, or a dotted member of a safe
    // family like border.*/shading.*/margin.*/indent.*/padding.*). Empty props
    // (e.g. a pure revision.action routed elsewhere) is treated as "may
    // restructure" → clears.
    private static bool ShouldClearNavCacheAfterSet(Dictionary<string, string> properties)
    {
        if (properties.Count == 0) return true;
        foreach (var key in properties.Keys)
        {
            if (NavCacheSafeSetKeys.Contains(key)) continue;
            var lower = key.ToLowerInvariant();
            var dot = lower.IndexOf('.');
            var family = dot >= 0 ? lower[..dot] : lower;
            if (family is "border" or "shading" or "margin" or "indent"
                or "padding" or "cellmargin" or "font")
                continue;
            return true; // a key we can't prove is attribute-only → clear
        }
        return false;
    }

    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        Modified = true;
        LastSetWarnings = new List<string>();
        LastUnrecognizedLatex = new List<string>();
        LastSetNewPath = null;
        var unsupported = new List<string>();

        // PERF(nav-child-cache): drop the run/row/cell nav caches on exit when
        // this Set might rewrite a container's child set (`set text`, revision
        // accept/reject, field/hyperlink rewrites, …). Property-only sets whose
        // keys are all attribute-only leave it disarmed. Exit-timing so a set
        // that repopulates the cache mid-op (navigating through the mutated
        // segment) still ends clean. See NavCacheSafeSetKeys.
        using var _navClearGuard = ShouldClearNavCacheAfterSet(properties)
            ? new NavCacheClearGuard(this)
            : default;

        // Bare `revision=` key was retired when the namespace split into
        // revision.type (creation) / revision.action (accept-reject). Reject
        // up-front with a pointed error so migrating scripts get an actionable
        // message instead of a generic "unsupported property" deep in dispatch.
        RejectBareRevisionKey(properties);

        // Revision selector / indexed path → dedicated dispatcher. Handles
        // `revision`, `revision[@author=X][@type=Y]`, and `/revision[N]` with
        // a single `--prop revision.action=accept|reject`. Routed before the
        // generic selector-batch branch so the synthetic /revision[N] path
        // (which has no real DOM target) doesn't trip Query→recursive-Set.
        if (IsRevisionSelectorPath(path))
            return SetRevisionsBySelector(path, properties);

        // Disambiguation guard: revision.action + find is rejected up-front.
        // Without this, the IsRevisionActionRequest router below grabs the
        // call and ignores the find=… key entirely, surfacing as "no revision
        // markers found at <path>" — a confusing error. The right interpretation
        // is "you tried to combine a creation context (find) with an action verb
        // (revision.action) — pick one".
        if (properties.ContainsKey("find") && properties.ContainsKey("revision.action"))
            throw new ArgumentException(
                "revision.action is reserved for the /revision selector (e.g. "
                + "`set /revision --prop revision.action=accept`); it cannot be combined "
                + "with find. To accept/reject changes produced by a find, run find first "
                + "(emits ins/del/rPrChange/pPrChange) then use the selector dispatch.");

        // Native-path accept/reject: `set /body/p[N] --prop revision.action=accept`,
        // `set /body/p[N]/r[M] --prop revision.action=reject`, etc. — accept/reject
        // every revision marker structurally tied to that element. Routed
        // BEFORE the rest of the dispatch so the action key isn't shadowed
        // by the creation-side Set.TrackChange decorator (creation and action
        // live in disjoint key namespaces: `revision.type` vs `revision.action`).
        if (IsRevisionActionRequest(properties))
            return SetRevisionByNativePath(path, properties);

        // Batch Set: route to the shared filter engine when the path is a bare
        // selector (no `/`) OR a `/`-scoped path that carries a content filter
        // (e.g. `/body/p[1]/r[bold=true or size=20pt]`). docx Query resolves the
        // slash-filter path via its bridge, so the same FilterSelector → Set-each
        // branch works; structural `/`-paths (`/body/p[@paraId=…]/r[2]`) stay
        // false in IsContentFilterPath and take NavigateToElement below.
        if (!string.IsNullOrEmpty(path)
            && (!path.StartsWith("/") || OfficeCli.Core.AttributeFilter.IsContentFilterPath(path)))
        {
            // FilterSelector narrows with the shared engine: a pure-AND selector
            // takes the flat path with applyAll:false (comparison ops only — the
            // handler already matched = / != with its looser semantics, so
            // set "run[size>14pt]" no longer sets every run); a selector with `or`
            // is queried bracket-stripped (handler returns broadly) then narrowed
            // by the boolean expression tree, which applies every predicate.
            var (targets, _) = OfficeCli.Core.AttributeFilter.FilterSelector(
                path, Query, keyResolver: null, applyAll: false);
            if (targets.Count == 0)
                throw new ArgumentException($"No elements matched selector: {path}");
            LastSelectorSetCount = targets.Count;
            foreach (var target in targets)
            {
                var targetUnsupported = Set(target.Path, properties);
                foreach (var u in targetUnsupported)
                    if (!unsupported.Contains(u)) unsupported.Add(u);
            }
            return unsupported;
        }

        // find / range are two mutually exclusive addressing modes (text match vs
        // explicit character offsets). Reject the contradiction rather than pick a
        // silent winner — mirrors the revision-namespace ambiguity rule.
        if (properties.ContainsKey("find") && properties.ContainsKey("range"))
            throw new ArgumentException(
                "'find' and 'range' are mutually exclusive addressing modes — provide one, not both.");

        // Explicit character-range run formatting: same split-run engine as find,
        // but the [start,end) offsets are supplied directly instead of derived from
        // a text match — giving a caller with a live text selection unambiguous
        // run-level formatting even when the selected text repeats. Run-level only;
        // offsets are 0-based, half-open, relative to the concatenated run text of
        // the resolved scope (a /body/p[N] path is that paragraph; /body spans all).
        // Structured to be isomorphic with this handler's own find path
        // (ProcessWordRange ≡ ProcessFind minus the match step); the cross-handler
        // range surface itself is CONSISTENCY(char-range).
        if (properties.TryGetValue("range", out var rangeSpec))
        {
            if (properties.ContainsKey("replace") || properties.ContainsKey("text"))
                throw new ArgumentException(
                    "range currently supports formatting only (bold, color, size, …); " +
                    "text insertion/replacement via range is not yet supported.");
            if (properties.Keys.Any(k => k.StartsWith("revision.", StringComparison.OrdinalIgnoreCase)))
                throw new ArgumentException(
                    "range cannot be combined with revision.* — use find (which infers "
                    + "ins/del) or an explicit /body/p[N]/r[M] path for tracked-change formatting.");
            var rangeFormatProps = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
            rangeFormatProps.Remove("range");
            rangeFormatProps.Remove("scope");
            rangeFormatProps.Remove("regex");
            if (rangeFormatProps.Count == 0)
                throw new ArgumentException(
                    "'range' requires format properties (e.g. bold, color, size).");
            var ranges = ParseHelpers.ParseCharRanges(rangeSpec);
            var rangeEffectivePath = (path is "" or "/") ? "/body" : path;
            foreach (var u in ProcessWordRange(rangeEffectivePath, ranges, rangeFormatProps))
                if (!unsupported.Contains(u)) unsupported.Add(u);
            return unsupported;
        }

        // Unified find: if 'find' key is present (at any path level), route to ProcessFind
        if (properties.TryGetValue("find", out var findText))
        {
            var replace = properties.TryGetValue("replace", out var r) ? r : null;
            // Separate run-level format properties from paragraph-level properties.
            // revision.* creation keys go to a third bucket so each scope (run /
            // paragraph) can be wrapped in the right *Change marker downstream.
            var formatProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var paraProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var revisionProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var (key, value) in properties)
            {
                var k = key.ToLowerInvariant();
                if (k is "find" or "replace" or "regex") continue;
                // revision.* (creation) carries author/date/id/type for wrapping
                // matched runs (w:ins/w:del) or capturing rPrChange/pPrChange.
                // revision.action is for /revision selectors, never for find.
                if (k.StartsWith("revision.", StringComparison.OrdinalIgnoreCase))
                {
                    revisionProps[key] = value;
                    continue;
                }
                // Paragraph-level properties go to paraProps
                if (k is "style" or "align" or "alignment" or "firstlineindent" or "leftindent" or "indentleft"
                    or "indent" or "rightindent" or "indentright" or "hangingindent" or "spacebefore"
                    or "spaceafter" or "linespacing" or "keepnext" or "keeplines" or "pagebreakbefore"
                    or "widowcontrol" or "liststyle" or "start" or "text" or "formula"
                    or "contextualspacing"
                    // direction is paragraph-scope (writes <w:bidi/> on pPr +
                    // <w:rtl/> cascade to runs); routing it as run-level
                    // would only stamp the run flag and skip the pPr bidi.
                    or "direction" or "dir" or "bidi")
                    paraProps[key] = value;
                else
                    formatProps[key] = value;
            }

            // ---- find + revision validation ----
            // revision.action is reserved for the /revision selector dispatch
            // — combining it with find is ambiguous (action verb vs creation context).
            if (revisionProps.ContainsKey("revision.action"))
                throw new ArgumentException(
                    "revision.action is reserved for the /revision selector (e.g. "
                    + "`set /revision --prop revision.action=accept`); it cannot be combined "
                    + "with find. To accept/reject changes produced by a find, run find first "
                    + "(emits ins/del/rPrChange/pPrChange) then use the selector dispatch.");
            // moveFrom/moveTo demand pairing via revision.id across two separate
            // targets — find has no way to express that pairing, so reject up-front
            // rather than silently producing one half of a broken pair.
            if (revisionProps.TryGetValue("revision.type", out var rtype)
                && (rtype.Equals("moveFrom", StringComparison.OrdinalIgnoreCase)
                    || rtype.Equals("moveTo", StringComparison.OrdinalIgnoreCase)))
                throw new ArgumentException(
                    $"revision.type={rtype} cannot be combined with find: move pairs require "
                    + "two explicitly-addressed runs sharing a revision.id. "
                    + "Use `set /body/p[N]/r[M] --prop revision.type=moveFrom --prop revision.id=N` "
                    + "and the matching moveTo on the destination run.");
            // Reject revision.type=ins/del on find without an actual action:
            // there's no destination text to mark as inserted (find matches
            // existing text — wrapping it as ins would record a fake "added by"
            // event) and del-without-replace is expressible via `--prop replace=""`.
            if (rtype != null
                && (rtype.Equals("ins", StringComparison.OrdinalIgnoreCase)
                    || rtype.Equals("insertion", StringComparison.OrdinalIgnoreCase)
                    || rtype.Equals("del", StringComparison.OrdinalIgnoreCase)
                    || rtype.Equals("deletion", StringComparison.OrdinalIgnoreCase)))
                throw new ArgumentException(
                    $"revision.type={rtype} on find is ambiguous — the find path infers "
                    + "ins/del from the operation: pass `--prop replace=NEW` to record a "
                    + "del(old)+ins(new) pair, or `--prop replace=` to record a delete-only. "
                    + "Drop revision.type from --prop and the wrap shape is inferred.");
            // Explicit revision.id is rejected on find: a single find call can
            // produce multiple markers (one w:del per matched-run fragment + one
            // w:ins for the replacement; one w:rPrChange per matched run for
            // format-only; one w:pPrChange per matched paragraph). A single
            // shared id would collide and break accept/reject by-id. The handler
            // auto-allocates from the shared paraId pool — each marker gets a
            // distinct id, addressable via the @id= selector on the resulting
            // /revision[@id=N] paths.
            if (revisionProps.ContainsKey("revision.id"))
                throw new ArgumentException(
                    "revision.id cannot be combined with find — a find can produce "
                    + "multiple markers per match and a shared id would collide. "
                    + "Drop revision.id from --prop; the handler auto-allocates per marker. "
                    + "If you need to address a specific marker afterwards, run `query revision` "
                    + "to read the assigned ids.");

            // text= cannot be combined with find: find/replace rewrites only the
            // matched fragment, while text= would rewrite the whole paragraph
            // body — the two intents conflict. text= has no
            // ApplyParagraphLevelProperty case, so previously it was routed to
            // paraProps and silently dropped while the confirmation still
            // reported it as applied (a lie). Surface it as unsupported instead
            // (exit 2 warning) so the caller knows it was not honoured; the
            // find/replace still proceeds. Mirrors the handler-as-truth
            // self-report model used by the non-find Set path.
            // Read via the source dictionary's TryGetValue (not paraProps) so a
            // TrackingPropertyDictionary marks "text" as accessed — otherwise the
            // wrapper would also flag it unsupported and we'd double-report.
            paraProps.Remove("text");
            if (properties.TryGetValue("text", out _))
                unsupported.Add("text");

            // Final gate: at least one of replace / format / para / revision-attribution
            // must be present. Bare `find=x` with nothing to do is a usage error.
            // Bare revision attribution (e.g. revision.author with no other prop) is
            // also rejected — there's nothing for the wrap to attribute.
            if (replace == null && formatProps.Count == 0 && paraProps.Count == 0)
                throw new ArgumentException(
                    "'find' requires either 'replace' and/or a format/paragraph property "
                    + "(e.g. bold, highlight, color, align). When combined with revision.*, "
                    + "the same rule applies — revision attribution needs a real action to "
                    + "attribute.");

            // CONSISTENCY(find-regex): canonical site for the `regex=true` → `r"..."`
            // raw-string normalization. `mark` and the other handlers' Set paths all
            // copy this pattern verbatim. To change the find/regex protocol,
            // grep "CONSISTENCY(find-regex)" and update every site project-wide;
            // do not diverge in a single handler.
            if (properties.TryGetValue("regex", out var regexFlag) && ParseHelpers.IsTruthySafe(regexFlag) && !findText.StartsWith("r\"") && !findText.StartsWith("r'"))
                findText = $"r\"{findText}\"";

            var effectivePath = (path is "" or "/") ? "/body" : path;
            var matchCount = ProcessFind(
                effectivePath,
                findText,
                replace,
                formatProps.Count > 0 ? formatProps : new Dictionary<string, string>(),
                revisionProps.Count > 0 ? revisionProps : null,
                out var matchedParagraphs);
            LastFindMatchCount = matchCount;

            // Apply paragraph-level properties to ONLY the paragraphs whose text
            // actually matched the find pattern. R8-fuzz-1 / R8-fuzz-2: re-resolving
            // via ResolveParagraphsForFind here ignores the find filter and
            // mass-rewrites every paragraph under the path (data corruption).
            //
            // When revisionProps is present, capture pPrChange per matched
            // paragraph by routing through BeginTrackChangeIfRequested (same
            // entry point used by `set /body/p[N] --prop align=center --prop
            // revision.author=X`), so find + paragraph-prop + revision matches
            // the non-find behaviour exactly.
            if (paraProps.Count > 0)
            {
                foreach (var para in matchedParagraphs)
                {
                    Action paragraphWrap = _trackChangeNoop;
                    if (revisionProps.Count > 0)
                    {
                        var combined = new Dictionary<string, string>(paraProps, StringComparer.OrdinalIgnoreCase);
                        foreach (var (rk, rv) in revisionProps) combined[rk] = rv;
                        var (_, wrap) = BeginTrackChangeIfRequested(para, combined);
                        paragraphWrap = wrap;
                    }
                    var pProps = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
                    foreach (var (key, value) in paraProps)
                    {
                        ApplyParagraphLevelProperty(pProps, key, value);
                        // CONSISTENCY(rtl-cascade): direction is paragraph-scope
                        // but Word's UI also stamps <w:rtl/> on every run + the
                        // paragraph mark when the user toggles direction. See
                        // WordHandler.I18n.cs.
                        var k = key.ToLowerInvariant();
                        if (k is "direction" or "dir" or "bidi")
                            ApplyDirectionCascade(para, ParseDirectionRtl(value));
                    }
                    paragraphWrap();
                }
            }

            SaveDoc();
            return unsupported;
        }

        // Document-level properties
        if (path == "/" || path == "" || path.Equals("/body", StringComparison.OrdinalIgnoreCase)
            || path.Equals("/document", StringComparison.OrdinalIgnoreCase))
        {
            SetDocumentProperties(properties, unsupported);
            SaveDoc();
            return unsupported;
        }

        // /docDefaults: same routing as /document, but bare keys (font,
        // fontSize, color, bold, …) implicitly target docDefaults. Prepend
        // "docdefaults." to any non-prefixed key so TrySetDocDefaults picks
        // them up. Keys already prefixed pass through unchanged.
        if (path.Equals("/docDefaults", StringComparison.OrdinalIgnoreCase))
        {
            var rewritten = new Dictionary<string, string>(properties.Comparer ?? StringComparer.OrdinalIgnoreCase);
            foreach (var (k, v) in properties)
            {
                if (k.StartsWith("docdefaults.", StringComparison.OrdinalIgnoreCase))
                    rewritten[k] = v;
                else
                    rewritten["docDefaults." + k] = v;
            }
            SetDocumentProperties(rewritten, unsupported);
            SaveDoc();
            return unsupported;
        }

        // Handle /settings path — route to SetDocumentProperties which calls TrySetDocSetting
        if (path.Equals("/settings", StringComparison.OrdinalIgnoreCase))
        {
            SetDocumentProperties(properties, unsupported);
            EnsureSettings().Save();
            return unsupported;
        }

        // Handle /watermark path
        if (path.Equals("/watermark", StringComparison.OrdinalIgnoreCase))
            return SetWatermarkPath(properties);

        // FormField paths: /formfield[N], /formfield[name], or /formfield[@name=NAME]
        // Routed BEFORE ParsePath because the generic predicate validator
        // only accepts positive-integer / last() / [@attr=v] predicates and
        // would reject the documented /formfield[name] form.
        // R14-bug4: schema declared /formfield[@name=NAME] as the stable
        // selector; mirror the [@name=…] arm here so set commands match
        // the same surface as get.
        var ffSetMatchEarly = System.Text.RegularExpressions.Regex.Match(path, @"^/formfield\[(?:@name=)?([^\]]+)\]$");
        if (ffSetMatchEarly.Success)
        {
            var allFormFields = FindFormFields();
            var indexOrName = ffSetMatchEarly.Groups[1].Value;
            bool nameForm = path.Contains("@name=", StringComparison.OrdinalIgnoreCase);
            (FieldInfo Field, FormFieldData FfData) target;
            if (!nameForm && int.TryParse(indexOrName, out var ffIdx))
            {
                if (ffIdx < 1 || ffIdx > allFormFields.Count)
                    throw new ArgumentException($"FormField {ffIdx} not found (total: {allFormFields.Count})");
                target = allFormFields[ffIdx - 1];
            }
            else
            {
                target = allFormFields.FirstOrDefault(ff =>
                    ff.FfData.GetFirstChild<FormFieldName>()?.Val?.Value == indexOrName);
                if (target.Field == null)
                    throw new ArgumentException($"FormField '{indexOrName}' not found");
            }
            return SetFormField(target, properties);
        }

        // Positional aliases /numbering/abstractNum[N] and /numbering/num[N]
        // translate to canonical [@id=K] form (mirrors Get's normalization in
        // commit 0257e8ca). Without this, Set on positional paths fell
        // through to generic Navigation, which has no NumberingInstance
        // branch — and CLI printed "Updated …" while nothing changed on
        // disk. Tagged CONSISTENCY(numbering-positional-normalize).
        var numPosSetMatch = System.Text.RegularExpressions.Regex.Match(
            path, @"^/numbering/(abstractNum|num)\[(\d+)\](.*)$");
        if (numPosSetMatch.Success)
        {
            var kind = numPosSetMatch.Groups[1].Value;
            var posIdx = int.Parse(numPosSetMatch.Groups[2].Value); // 1-based
            var rest = numPosSetMatch.Groups[3].Value;
            var nb = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
            int? canonId = null;
            if (kind == "abstractNum")
            {
                var abs = nb?.Elements<AbstractNum>().ElementAtOrDefault(posIdx - 1);
                canonId = abs?.AbstractNumberId?.Value;
            }
            else
            {
                var inst = nb?.Elements<NumberingInstance>().ElementAtOrDefault(posIdx - 1);
                canonId = inst?.NumberID?.Value;
            }
            if (canonId != null)
                return Set($"/numbering/{kind}[@id={canonId}]{rest}", properties);
        }

        // Numbering paths: /numbering/abstractNum[@id=N] and
        // /numbering/abstractNum[@id=N]/level[L]. Intercept BEFORE the generic
        // ParsePath call below — those paths use [@id=...] / [N starting at 0]
        // predicates that ParsePath's 1-based positional rule rejects.
        // Accept both /level[N] (positional, 0-based ilvl) and /lvl[@ilvl=N]
        // (canonical form returned by Get/Query — see R2 commit 48ee8c8c, R3
        // commit 2a634aeb). Without the @ilvl branch, Set silently no-ops on
        // the canonical path: the CLI prints "Updated" but numbering.Save()
        // never runs because the path falls through to generic Navigation
        // which has no Level branch in SetElement.
        var absNumSetMatchEarly = System.Text.RegularExpressions.Regex.Match(
            path, @"^/numbering/abstractNum\[@id=(\d+)\](?:/(?:level|lvl)\[(?:@ilvl=)?(\d+)\])?$");
        if (absNumSetMatchEarly.Success) return SetAbstractNumPath(absNumSetMatchEarly, properties);

        // /numbering/num[@id=N] — set abstractNumId on a NumberingInstance.
        // Without this intercept, generic Navigation finds the <w:num> element
        // but SetElement has no NumberingInstance branch, so the call returns
        // an empty unsupported list and the CLI prints "Updated …" while
        // nothing changes on disk.
        var numSetMatchEarly = System.Text.RegularExpressions.Regex.Match(
            path, @"^/numbering/num\[@id=(\d+)\]$");
        if (numSetMatchEarly.Success) return SetNumPath(numSetMatchEarly, properties);

        // Handle header/footer paths
        var hfParts = ParsePath(path);
        if (hfParts.Count >= 1)
        {
            var firstName = hfParts[0].Name.ToLowerInvariant();
            if ((firstName == "header" || firstName == "footer") && hfParts.Count == 1)
            {
                // last() resolves to the LAST part by creation order (mirrors
                // get); otherwise set /header[last()] silently wrote into the
                // FIRST header (Index == null → 0), losing content.
                int hfCount = (firstName == "header"
                    ? _doc.MainDocumentPart?.HeaderParts.Count()
                    : _doc.MainDocumentPart?.FooterParts.Count()) ?? 0;
                int hfIdx = hfParts[0].StringIndex == "last()" ? hfCount - 1 : (hfParts[0].Index ?? 1) - 1;
                SetHeaderFooter(firstName, hfIdx, properties, unsupported);
                return unsupported;
            }
        }

        // Chart axis-by-role sub-path: /chart[N]/axis[@role=ROLE].
        var chartAxisSetMatch = System.Text.RegularExpressions.Regex.Match(path,
            @"^/chart\[(\d+)\]/axis\[@role=([a-zA-Z0-9_]+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (chartAxisSetMatch.Success) return SetChartAxisPath(chartAxisSetMatch, properties);

        // Chart paths: /chart[N] or /chart[N]/series[K]
        var chartMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/chart\[(\d+)\](?:/series\[(\d+)\])?$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (chartMatch.Success) return SetChartPath(chartMatch, properties);

        // Field paths: /field[N]
        var fieldSetMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/field\[(\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (fieldSetMatch.Success) return SetFieldPath(fieldSetMatch, properties);

        // TOC paths: /toc[N], /toc (= first), /tableofcontents alias.
        var tocMatch = System.Text.RegularExpressions.Regex.Match(path,
            @"^/(?:toc|tableofcontents)(?:\[(\d+)\])?$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (tocMatch.Success) return SetTocPath(tocMatch, properties);

        // Footnote paths: /footnote[N], /footnote[@footnoteId=N] (incl. -1/0
        // structural ids — separator/continuation/continuationNotice).
        var fnSetMatch = System.Text.RegularExpressions.Regex.Match(
            path, @"/footnote\[(?:@footnoteId=)?(-?\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (fnSetMatch.Success) return SetFootnotePath(fnSetMatch, properties);

        // Endnote paths: same shape as footnote.
        var enSetMatch = System.Text.RegularExpressions.Regex.Match(
            path, @"/endnote\[(?:@endnoteId=)?(-?\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (enSetMatch.Success) return SetEndnotePath(enSetMatch, properties);

        // CONSISTENCY(path-element-case-insensitive): same rule as Query.cs — top-level
        // element paths (/section[N], /body/sectPr[N], /chart[N], /toc[N], …) match
        // case-insensitively so /Section[1] is equivalent to /section[1]. styleSetMatch
        // below remains case-sensitive — style ids are user-defined identifiers.
        var secSetMatch = System.Text.RegularExpressions.Regex.Match(path, @"^(?:/section\[(\d+|last\(\))\]|/body/sectPr(?:\[(\d+)\])?)$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (secSetMatch.Success) return SetSectionPath(secSetMatch, properties);

        // Style paths: /styles/StyleId (set props on the style itself).
        // Restrict to a single segment so deeper paths like /styles/<id>/tab[N]
        // fall through to generic Navigation + SetElement (TabStop branch).
        var styleSetMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/styles/([^/]+)$");
        if (styleSetMatch.Success) return SetStylePath(styleSetMatch, properties);

        // CONSISTENCY(ole-shorthand-set): mirror the /body/ole[N] shorthand
        // already supported in Get (WordHandler.Query.cs) and Remove
        // (WordHandler.Mutations.cs). Without this intercept, Set falls through
        // to NavigateToElement which hits "No ole found at /body" because OLE
        // lives inside a run, not as a direct child of the body.
        var wordOleSetMatch = System.Text.RegularExpressions.Regex.Match(
            path,
            @"^(?<parent>/body|/header\[\d+\]|/footer\[\d+\])?/(?:ole|object|embed)\[(?<idx>\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (wordOleSetMatch.Success) return SetWordOlePath(wordOleSetMatch, properties);

        var parts = ParsePath(path);
        var element = NavigateToElement(parts, out var ctx);
        if (element == null)
            throw new ArgumentException($"Path not found: {path}" + (ctx != null ? $". {ctx}" : ""));

        // Clone element for rollback on failure (atomic: no partial modifications)
        var elementBackup = element.CloneNode(true);
        try
        {
            // Phase 2 trackChange capture: snapshot rPr/pPr BEFORE SetElement
            // mutates it, strip trackChange.* from the dict so downstream
            // helpers don't see them, then append rPrChange/pPrChange AFTER
            // the Set succeeds. See WordHandler.Set.TrackChange.cs.
            var (effectiveProps, wrapTrackChange) = BeginTrackChangeIfRequested(element, properties);
            var result = SetElement(element, effectiveProps);
            wrapTrackChange();
            return result;
        }
        catch
        {
            // Rollback: restore element to pre-modification state
            element.Parent?.ReplaceChild(elementBackup, element);
            throw;
        }
    }

    private List<string> SetElement(OpenXmlElement element, Dictionary<string, string> properties)
    {
        if (element is Comment cmt) return SetElementComment(cmt, properties);
        if (element is BookmarkStart bk) return SetElementBookmark(bk, properties);
        if (element is SdtBlock || element is SdtRun) return SetElementSdt(element, properties);
        if (element is Run run) return SetElementRun(run, properties);
        if (element is Hyperlink hl) return SetElementHyperlink(hl, properties);
        if (element is M.Paragraph mPara) return SetElementMPara(mPara, properties);
        if (element is M.OfficeMath oMathEl) return SetElementOMath(oMathEl, properties);
        if (element is Paragraph para) return SetElementParagraph(para, properties);
        if (element is TableCell cell) return SetElementTableCell(cell, properties);
        if (element is TableRow row) return SetElementTableRow(row, properties);
        if (element is Table tbl) return SetElementTable(tbl, properties);
        if (element is TabStop tabStop) return SetElementTabStop(tabStop, properties);
        if (element is TextBoxContent txbx) return SetElementTextBoxContent(txbx, properties);
        // /body/shape[N] resolves to the wps:wsp element. SetShapeProps models the
        // curated spPr surface (fill/line/width/height/geometry) reusing the Add
        // builders, and forwards out-of-scope keys as unsupported so Add and
        // Set track the same prop surface.
        if (string.Equals(element.LocalName, "wsp", StringComparison.Ordinal)
            && string.Equals(element.NamespaceUri,
                "http://schemas.microsoft.com/office/word/2010/wordprocessingShape", StringComparison.Ordinal))
            return SetShapeProps(element, properties);
        // /body/group[N] resolves to the wpg:wgp group. Resizing scales the whole
        // group (grpSpPr ext + ancestor wp:extent, leaving chExt as the baseline
        // so Word compresses the children) and re-bakes child font sizes.
        if (string.Equals(element.LocalName, "wgp", StringComparison.Ordinal)
            && string.Equals(element.NamespaceUri,
                "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup", StringComparison.Ordinal))
            return SetGroupProps(element, properties);
        // Other shape carriers (Drawing host or v:shape descendant) have no
        // curated Set handler. Without this guard the dispatcher returned an empty
        // unsupported list → CLI reported "Updated" (exit 0) while writing
        // nothing. Surface every key as unsupported + a warning so the call
        // signals the gap clearly, mirroring how Add returns UNSUPPORTED for
        // shape props it doesn't model.
        bool isShape = element is Drawing
            || string.Equals(element.LocalName, "wsp", StringComparison.Ordinal)
            || string.Equals(element.LocalName, "shape", StringComparison.Ordinal);
        if (isShape)
        {
            var dropped = new List<string>();
            foreach (var k in properties.Keys)
            {
                dropped.Add(k);
                LastSetWarnings.Add($"unsupported property on shape: '{k}' (Set does not currently model shape transforms; see Add path for the supported props).");
            }
            return dropped;
        }
        return new List<string>();
    }

    private void SetHeaderFooter(string kind, int index, Dictionary<string, string> properties, List<string> unsupported)
    {
        var mainPart = _doc.MainDocumentPart!;
        OpenXmlCompositeElement? container;
        OpenXmlPart partRef;

        if (kind == "header")
        {
            var part = mainPart.HeaderParts.ElementAtOrDefault(index)
                ?? throw new ArgumentException($"Header not found: /header[{index + 1}]");
            container = part.Header;
            partRef = part;
        }
        else
        {
            var part = mainPart.FooterParts.ElementAtOrDefault(index)
                ?? throw new ArgumentException($"Footer not found: /footer[{index + 1}]");
            container = part.Footer;
            partRef = part;
        }

        if (container == null)
            throw new ArgumentException($"{kind} content not found at index {index + 1}");

        var firstPara = container.Elements<Paragraph>().FirstOrDefault();
        if (firstPara == null)
        {
            firstPara = new Paragraph();
            container.AppendChild(firstPara);
        }
        var pProps = firstPara.ParagraphProperties ?? firstPara.PrependChild(new ParagraphProperties());

        foreach (var (key, value) in properties)
        {
            var k = key.ToLowerInvariant();
            if (ApplyParagraphLevelProperty(pProps, key, value))
            {
                // CONSISTENCY(rtl-cascade): direction=rtl on header/footer
                // must also stamp <w:rtl/> on the paragraph mark and runs.
                // See WordHandler.I18n.cs.
                if (k is "direction" or "dir" or "bidi")
                    ApplyDirectionCascade(firstPara, ParseDirectionRtl(value));
                // handled by paragraph-level helper
            }
            else switch (k)
            {
                case "text":
                {
                    // CONSISTENCY(xml-text-validation): mirror AppendTextWithBreaks —
                    // reject XML 1.0 illegal control chars at input time so the resident
                    // process doesn't accept them into the in-memory DOM only to fail at
                    // close with "save failed — data may be lost" and lose user work.
                    ParseHelpers.ValidateXmlText(value, "text", allowSoftBreakChar: true);
                    // Only replace non-field static text runs. Complex fields are
                    // a multi-run sequence: [Begin][Instr]([Separate][Result])[End].
                    // Runs carrying <w:fldChar>/<w:instrText> AND any run nested
                    // between Separate and End (the field "result" run) must all
                    // survive — otherwise PAGE/DATE/etc. embedded in header/footer
                    // are silently destroyed.
                    var paraRuns = firstPara.Elements<Run>().ToList();
                    var inField = false;
                    var fieldRunSet = new HashSet<Run>();
                    foreach (var r in paraRuns)
                    {
                        var fldChar = r.Elements<FieldChar>().FirstOrDefault();
                        var hasInstr = r.Elements<FieldCode>().Any();
                        if (fldChar != null || hasInstr)
                        {
                            fieldRunSet.Add(r);
                            if (fldChar?.FieldCharType?.Value == FieldCharValues.Begin) inField = true;
                            else if (fldChar?.FieldCharType?.Value == FieldCharValues.End) inField = false;
                        }
                        else if (inField)
                        {
                            fieldRunSet.Add(r);
                        }
                    }
                    RunProperties? existingRProps = null;
                    var firstStaticRun = paraRuns.FirstOrDefault(r => !fieldRunSet.Contains(r));
                    if (firstStaticRun?.RunProperties != null)
                        existingRProps = (RunProperties)firstStaticRun.RunProperties.CloneNode(true);
                    var firstFieldRun = paraRuns.FirstOrDefault(fieldRunSet.Contains);
                    foreach (var r in paraRuns.Where(r => !fieldRunSet.Contains(r))) r.Remove();
                    var newRun = new Run();
                    if (existingRProps != null)
                        newRun.AppendChild(existingRProps);
                    // CONSISTENCY(text-breaks): route through AppendTextWithBreaks
                    // so \n/\t in value become <w:br/>/<w:tab/>, matching Add and
                    // body-paragraph Set behavior (WordHandler.Set.Element.cs).
                    AppendTextWithBreaks(newRun, value);
                    if (firstFieldRun != null)
                        firstPara.InsertBefore(newRun, firstFieldRun);
                    else
                        firstPara.AppendChild(newRun);
                    break;
                }
                // Per-script font slots and CS run flags follow the same dispatch
                // as bare bold/italic/size — ApplyRunFormatting handles the
                // canonical and alias forms, so they are listed here for the
                // header/footer route to reach it (mirrors body-paragraph Set
                // dispatch in Set.Element.cs).
                case "size" or "fontsize" or "font" or "bold" or "italic" or "color" or "highlight" or "underline" or "strike"
                  or "font.latin" or "font.ea" or "font.eastasia" or "font.eastasian"
                  or "font.cs" or "font.complexscript" or "font.complex"
                  or "bold.cs" or "italic.cs" or "size.cs"
                  or "font.bold.cs" or "font.italic.cs" or "font.size.cs"
                  or "boldcs" or "italiccs" or "sizecs":
                    // Apply run-level formatting to all runs in the container
                    foreach (var run in container.Descendants<Run>())
                        ApplyRunFormatting(EnsureRunProperties(run), key, value);
                    // Also update paragraph mark run properties so new runs inherit formatting
                    var markRPr = pProps.ParagraphMarkRunProperties ?? pProps.AppendChild(new ParagraphMarkRunProperties());
                    ApplyRunFormatting(markRPr, key, value);
                    break;
                case "type":
                {
                    // Mutate the HeaderReference/FooterReference Type attribute
                    // pointing at this part. Read side (WordHandler.Query.cs:660-666,
                    // 717-723) only inspects body-level SectionProperties, so the
                    // write side stays scoped to the same set for round-trip parity.
                    var newType = value.ToLowerInvariant() switch
                    {
                        "first" => HeaderFooterValues.First,
                        "even" => HeaderFooterValues.Even,
                        "default" => HeaderFooterValues.Default,
                        _ => throw new ArgumentException(
                            $"Invalid {kind} type: '{value}'. Valid values: default, first, even.")
                    };
                    var partRid = mainPart.GetIdOfPart(partRef);
                    var body = mainPart.Document?.Body
                        ?? throw new InvalidOperationException("Document body not found");
                    bool found = false;
                    foreach (var sp in body.Elements<SectionProperties>())
                    {
                        if (kind == "header")
                        {
                            var ownRef = sp.Elements<HeaderReference>().FirstOrDefault(r => r.Id?.Value == partRid);
                            if (ownRef == null) continue;
                            if (ownRef.Type?.Value == newType) { found = true; continue; }
                            if (sp.Elements<HeaderReference>().Any(r => r != ownRef && r.Type?.Value == newType))
                                throw new ArgumentException(
                                    $"Header of type '{value}' already exists in this section.");
                            ownRef.Type = newType;
                            found = true;
                        }
                        else
                        {
                            var ownRef = sp.Elements<FooterReference>().FirstOrDefault(r => r.Id?.Value == partRid);
                            if (ownRef == null) continue;
                            if (ownRef.Type?.Value == newType) { found = true; continue; }
                            if (sp.Elements<FooterReference>().Any(r => r != ownRef && r.Type?.Value == newType))
                                throw new ArgumentException(
                                    $"Footer of type '{value}' already exists in this section.");
                            ownRef.Type = newType;
                            found = true;
                        }
                        // Mirrors AddHeader: Title-page header requires <w:titlePg/> on the section.
                        if (newType == HeaderFooterValues.First && sp.GetFirstChild<TitlePage>() == null)
                            sp.AddChild(new TitlePage(), throwOnError: false);
                    }
                    if (!found) unsupported.Add(key);
                    break;
                }
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        if (kind == "header")
            mainPart.HeaderParts.ElementAt(index).Header?.Save();
        else
            mainPart.FooterParts.ElementAt(index).Footer?.Save();
    }

    // Border style format: "style" or "style;size" or "style;size;color" or "style;size;color;space"
    // Styles: none, single, thick, double, dotted, dashed, dotDash, dotDotDash, triple,
    //         thinThickSmallGap, thickThinSmallGap, thinThickThinSmallGap,
    //         thinThickMediumGap, thickThinMediumGap, thinThickThinMediumGap,
    //         thinThickLargeGap, thickThinLargeGap, thinThickThinLargeGap, wave, doubleWave, threeDEmboss, threeDEngrave
    /// <summary>Insert StyleParagraphProperties before StyleRunProperties to maintain OOXML schema order.</summary>
    private static StyleParagraphProperties EnsureStyleParagraphProperties(Style style)
    {
        var pPr = new StyleParagraphProperties();
        var rPr = style.StyleRunProperties;
        if (rPr != null)
            style.InsertBefore(pPr, rPr);
        else
            style.AppendChild(pPr);
        return pPr;
    }

    private static BorderValues ParseBorderStyle(string style) => style.ToLowerInvariant() switch
    {
        "none" => BorderValues.None,
        "nil" => BorderValues.Nil,
        "single" or "thin" => BorderValues.Single,
        "thick" or "medium" => BorderValues.Thick,
        "double" => BorderValues.Double,
        "dotted" => BorderValues.Dotted,
        "dashed" => BorderValues.Dashed,
        "dotdash" => BorderValues.DotDash,
        "dotdotdash" => BorderValues.DotDotDash,
        "triple" => BorderValues.Triple,
        "thinthicksmallgap" => BorderValues.ThinThickSmallGap,
        "thickthinsmallgap" => BorderValues.ThickThinSmallGap,
        "thinthickthinsmallgap" => BorderValues.ThinThickThinSmallGap,
        "thinthickmediumgap" => BorderValues.ThinThickMediumGap,
        "thickthinmediumgap" => BorderValues.ThickThinMediumGap,
        "thinthickthinmediumgap" => BorderValues.ThinThickThinMediumGap,
        "thinthicklargegap" => BorderValues.ThinThickLargeGap,
        "thickthinlargegap" => BorderValues.ThickThinLargeGap,
        "thinthickthinlargegap" => BorderValues.ThinThickThinLargeGap,
        // BUG-DUMP23-02: dashSmallGap is a valid ST_Border token (just wasn't
        // listed). wavy is a colloquial alias for wave. wavyDouble / wavyHeavy
        // are not part of ECMA-376 ST_Border (the SDK rejects them at validate
        // time even via the string ctor) — accept them as input aliases and
        // map to the nearest valid token (DoubleWave) so add/set don't reject
        // user-supplied style names that show up in real-world docs.
        "wave" or "wavy" => BorderValues.Wave,
        "doublewave" or "wavydouble" or "wavyheavy" => BorderValues.DoubleWave,
        "dashsmallgap" => BorderValues.DashSmallGap,
        // BUG-R4B(BUG3): outset / inset are valid ST_Border tokens (real-world
        // tables emit them) and dump→replay rejected them, cascading to a
        // whole-table add failure. dashDotStroked was likewise a valid
        // ST_Border value missing from the parser. Add the remaining
        // line-style ST_Border members so they round-trip.
        "outset" => BorderValues.Outset,
        "inset" => BorderValues.Inset,
        "dashdotstroked" => BorderValues.DashDotStroked,
        "threedembed" or "3demboss" => BorderValues.ThreeDEmboss,
        "threedengrave" or "3dengrave" => BorderValues.ThreeDEngrave,
        // BUG-DUMP-ARTBORDER: ST_Border has a large ART family (hearts, apples,
        // balloons3Colors, …) with no named SDK fields. Mapping only line-style
        // tokens above and throwing here dropped art page borders on dump→batch
        // AND collaterally aborted the whole atomic `set /` command — taking the
        // valid `background` prop down with it. Be lenient (project input
        // philosophy): pass any unknown token through the string-backed
        // BorderValues ctor. Valid ST_Border values round-trip; genuinely bad
        // tokens produce XML that Word ignores / validate flags — same outcome
        // as other lenient parses, never a hard reject.
        _ => new BorderValues(style)
    };

    // CONSISTENCY(border-empty-segment): space is uint? rather than uint so the
    // caller can distinguish "not specified" from "explicitly 0" — the OOXML
    // default for w:space is 0, so writing the attribute when the user did not
    // ask for it round-trips into a spurious `border.X.space: 0` readback (and
    // a `STYLE;SZ;;0` artifact in batch dump).
    private static (BorderValues style, uint size, string? color, uint? space) ParseBorderValue(string value)
    {
        var (style, size, color, space, _) = ParseBorderValueDetailed(value);
        return (style, size, color, space);
    }

    // BUG-DUMP-R36-1: the compound border string carries two optional boolean
    // segments AFTER space — STYLE;SIZE;COLOR;SPACE;SHADOW;FRAME. w:shadow draws
    // a drop-shadow, w:frame draws the border on the OUTSIDE edge of text; both
    // change rendering and were previously dropped on round-trip. Segments are
    // "true"/"false"/"" (empty = not specified). Back-compat: a 4-field string
    // returns (null, null) so MakeBorder stamps neither attribute.
    //
    // CONSISTENCY(border-shadow-frame): the positional SHADOW/FRAME segments are
    // skipped if they contain '=' (a theme key=val tail — see
    // ExtractThemeTail/BUG-DUMP-R41-2) so a string carrying only theme tails
    // (e.g. "single;12;;;themeColor=accent1") never mis-reads the tail segment
    // as a boolean shadow/frame value.
    private static (bool? shadow, bool? frame) ParseBorderShadowFrame(string value)
    {
        var parts = value.Split(';');
        bool? Slot(int i) =>
            (parts.Length > i && !string.IsNullOrEmpty(parts[i].Trim())
             && !parts[i].Contains('='))
                ? IsTruthy(parts[i].Trim())
                : (bool?)null;
        return (Slot(4), Slot(5));
    }

    // BUG-DUMP-R41-2 / BUG-DUMP-R41-4: theme-color/theme-fill linkage on a
    // border or shading carries the document-theme slot (accent1, text1, …)
    // plus an optional shade/tint, independent of the resolved hex. The
    // compound border / shading strings are positional (STYLE;SIZE;COLOR;…
    // resp. VAL;FILL[;COLOR]); to round-trip the theme attrs WITHOUT breaking
    // the positional parse the emitter appends backward-compatible `key=val`
    // tail segments (`;themeColor=accent1;themeShade=BF;themeTint=99` for
    // borders, `;themeFill=text1;themeFillShade=..;themeFillTint=..` plus the
    // same themeColor/Shade/Tint trio for shading). This helper splits the raw
    // value into (positionalOnly, themeMap): every `;`-segment containing '='
    // is harvested into themeMap (key lower-cased), and the remaining segments
    // are rejoined with ';' so the legacy positional parser sees an unchanged
    // 4-field (or shorter) string. A value with no theme tail returns the input
    // verbatim and an empty map — plain borders/shading are untouched.
    private static (string positional, Dictionary<string, string> theme) ExtractThemeTail(string value)
    {
        if (!value.Contains('=')) return (value, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase));
        var theme = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var kept = new List<string>();
        foreach (var seg in value.Split(';'))
        {
            var eq = seg.IndexOf('=');
            if (eq > 0)
            {
                var k = seg[..eq].Trim();
                var v = seg[(eq + 1)..].Trim();
                if (k.Length > 0) theme[k] = v;
            }
            else
                kept.Add(seg);
        }
        return (string.Join(';', kept), theme);
    }

    // BUG-DUMP-R41-2: stamp w:themeColor / w:themeShade / w:themeTint on a
    // <w:*Border> from the theme map produced by ExtractThemeTail. Only the
    // keys actually present are written, so a plain border is unaffected.
    private static void ApplyBorderTheme(BorderType border, Dictionary<string, string> theme)
    {
        if (theme.Count == 0) return;
        if (theme.TryGetValue("themeColor", out var tc) && tc.Length > 0)
            border.ThemeColor = new EnumValue<ThemeColorValues>(new ThemeColorValues(tc));
        if (theme.TryGetValue("themeShade", out var tsh) && tsh.Length > 0)
            border.ThemeShade = tsh;
        if (theme.TryGetValue("themeTint", out var tt) && tt.Length > 0)
            border.ThemeTint = tt;
    }

    // BUG-DUMP-R41-4: stamp the theme shading attrs (w:themeFill /
    // w:themeFillShade / w:themeFillTint and the w:themeColor / w:themeShade /
    // w:themeTint trio) on a <w:shd> from the ExtractThemeTail map. Only
    // present keys are written.
    private static void ApplyShadingTheme(Shading shd, Dictionary<string, string> theme)
    {
        if (theme.Count == 0) return;
        if (theme.TryGetValue("themeFill", out var tf) && tf.Length > 0)
            shd.ThemeFill = new EnumValue<ThemeColorValues>(new ThemeColorValues(tf));
        if (theme.TryGetValue("themeFillShade", out var tfs) && tfs.Length > 0)
            shd.ThemeFillShade = tfs;
        if (theme.TryGetValue("themeFillTint", out var tft) && tft.Length > 0)
            shd.ThemeFillTint = tft;
        if (theme.TryGetValue("themeColor", out var tc) && tc.Length > 0)
            shd.ThemeColor = new EnumValue<ThemeColorValues>(new ThemeColorValues(tc));
        if (theme.TryGetValue("themeShade", out var tsh) && tsh.Length > 0)
            shd.ThemeShade = tsh;
        if (theme.TryGetValue("themeTint", out var tt) && tt.Length > 0)
            shd.ThemeTint = tt;
    }

    // BUG-DUMP-R43-3: stamp w:themeColor / w:themeShade / w:themeTint on a
    // <w:color> from the ExtractThemeTail map. Mirrors ApplyBorderTheme /
    // ApplyShadingTheme. Only present keys are written, so a plain color is
    // unaffected. Used by AddStyle for style run-color theme linkage.
    private static void ApplyColorTheme(Color color, Dictionary<string, string> theme)
    {
        if (theme.Count == 0) return;
        if (theme.TryGetValue("themeColor", out var tc) && tc.Length > 0)
            color.ThemeColor = new EnumValue<ThemeColorValues>(new ThemeColorValues(tc));
        if (theme.TryGetValue("themeShade", out var tsh) && tsh.Length > 0)
            color.ThemeShade = tsh;
        if (theme.TryGetValue("themeTint", out var tt) && tt.Length > 0)
            color.ThemeTint = tt;
    }

    // CONSISTENCY(border-size-roundtrip): expose whether the caller provided
    // a SIZE segment so MakeBorder can suppress w:sz on nil borders that
    // never had it in the source. Mirrors the w:space "only when given"
    // pattern. Internal helper — callers that don't care about the flag
    // use the legacy 4-tuple wrapper above.
    private static (BorderValues style, uint size, string? color, uint? space, bool sizeProvided) ParseBorderValueDetailed(string value)
    {
        // BUG-DUMP-R41-2: strip any trailing theme key=val tail
        // (themeColor=…/themeShade=…/themeTint=…) before positional parsing so
        // it never lands in the COLOR/SPACE slots. The theme attrs are applied
        // separately via ParseBorderTheme → ApplyBorderTheme at the call site.
        var (value0, _) = ExtractThemeTail(value);
        value = value0;
        var parts = value.Split(';');
        var style = ParseBorderStyle(parts[0]);
        uint size;
        bool sizeProvided = parts.Length > 1 && !string.IsNullOrEmpty(parts[1].Trim());
        // CONSISTENCY(border-empty-segment): mirror the empty-color tolerance
        // below — WordBatchEmitter's border fold emits "STYLE;;COLOR" whenever a
        // side has color but no explicit sz attribute (very common in real
        // .docx files where w:sz is inherited via the style chain). Treat an
        // empty SIZE segment as "use default" instead of throwing.
        if (parts.Length > 1 && !string.IsNullOrEmpty(parts[1].Trim()))
        {
            // OOXML stores border size in eighth-of-a-point units. Accept bare
            // integer (already in eighths) plus unit-qualified lengths
            // ('1pt', '0.5cm', '0.05in') for parity with other Word length
            // inputs (CONSISTENCY: spacing-units, the project conventions "Spacing
            // input is lenient").
            var sz = parts[1].Trim();
            if (uint.TryParse(sz, out size))
            { /* bare integer = eighths */ }
            else if (sz.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
                     && double.TryParse(sz[..^2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var pts) && pts >= 0)
                size = (uint)Math.Round(pts * 8);
            else if (sz.EndsWith("cm", StringComparison.OrdinalIgnoreCase)
                     && double.TryParse(sz[..^2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var cm) && cm >= 0)
                size = (uint)Math.Round(cm * (72.0 / 2.54) * 8); // cm → pt → eighths
            else if (sz.EndsWith("in", StringComparison.OrdinalIgnoreCase)
                     && double.TryParse(sz[..^2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var inches) && inches >= 0)
                size = (uint)Math.Round(inches * 72 * 8); // in → pt → eighths
            else
                throw new ArgumentException($"Invalid border size '{parts[1]}', expected integer (eighths-of-pt) or unit-qualified length (e.g. '1pt', '0.5cm'). Format: STYLE[;SIZE[;COLOR[;SPACE]]]");
        }
        else
            size = style == BorderValues.Nil ? 0u : style == BorderValues.Thick ? 12u : 4u;
        // BUG-R7-02: dump emits "nil;0;;0" for nil borders (empty color
        // segment). SanitizeHex rejects empty input with "Invalid color
        // value: ''", breaking round-trip on any docx with nil paragraph
        // borders. Treat an empty color segment as "no color" (null) rather
        // than a parse error — this matches dump's emit semantics.
        string? color = (parts.Length > 2 && !string.IsNullOrEmpty(parts[2]))
            ? SanitizeHex(parts[2])
            : null;
        uint? space = null;
        // CONSISTENCY(border-empty-segment): symmetric with the SIZE/COLOR
        // tolerance — empty SPACE segment means "no override".
        if (parts.Length > 3 && !string.IsNullOrEmpty(parts[3]))
        {
            if (!uint.TryParse(parts[3], out var spaceVal))
                throw new ArgumentException($"Invalid border space '{parts[3]}', expected integer. Format: STYLE[;SIZE[;COLOR[;SPACE]]]");
            space = spaceVal;
        }
        return (style, size, color, space, sizeProvided);
    }

    private static T MakeBorder<T>(BorderValues style, uint size, string? color, uint? space, bool? shadow = null, bool? frame = null, Dictionary<string, string>? theme = null) where T : BorderType, new()
    {
        // BUG-R2-P2-7: only emit w:space attribute when the caller actually
        // provided one. Writing space=0 explicitly round-trips into a
        // spurious `border.X.space: 0` readback and a `STYLE;SZ;;0` batch
        // artifact, even though 0 is the OOXML default and the user never
        // asked for it.
        // CONSISTENCY(border-size-roundtrip): a bare `nil` value (no sz
        // segment) round-trips with size=0 default — but emitting w:sz="0"
        // surfaces a `nil;0` readback on the next dump. Callers that have
        // sizeProvided info route through MakeBorderTyped below; the legacy
        // overload stays size-always-stamped for non-nil borders that
        // genuinely default to size 4.
        T b;
        try { b = new T { Val = style }; }
        catch (Exception ex) when (ex is not OutOfMemoryException and not StackOverflowException)
        {
            // ParseBorderStyle is lenient (passes unknown tokens via
            // `new BorderValues(raw)` to support art borders). The SDK's
            // EnumValue<BorderValues> setter validates on assignment and
            // throws an enum-type exception with the opaque "specified
            // value is not valid" message. Convert to a friendly
            // ArgumentException carrying the bad token.
            throw new ArgumentException(
                $"Invalid border style '{style.ToString()}'. " +
                "Valid styles include: single, double, dashed, dotted, thick, thin, " +
                "wave, doubleWave, dashSmallGap, dotDash, dotDotDash, triple, outset, " +
                "inset, dashDotStroked, threeDEmboss, threeDEngrave, none, nil.", ex);
        }
        if (!(style == BorderValues.Nil && size == 0)) b.Size = size;
        if (space.HasValue) b.Space = space.Value;
        if (color != null) b.Color = color;
        // BUG-DUMP-R36-1: stamp w:shadow / w:frame only when the source carried
        // them (bool? null = absent) so a plain border round-trips without a
        // spurious shadow="false"/frame="false" attribute.
        if (shadow.HasValue) b.Shadow = shadow.Value;
        if (frame.HasValue) b.Frame = frame.Value;
        // BUG-DUMP-R41-2: stamp w:themeColor / w:themeShade / w:themeTint from
        // the theme map (parsed out of the value's key=val tail) so the
        // border's theme linkage round-trips. Null/empty map = plain border,
        // untouched.
        if (theme != null) ApplyBorderTheme(b, theme);
        return b;
    }

    /// <summary>
    /// Apply a paragraph-level property. Returns true if handled, false if not recognized.
    /// Handles: style, alignment, indent, spacing, keepNext, keepLines, pageBreakBefore, widowControl, shading, pbdr.
    /// </summary>
    private bool ApplyParagraphLevelProperty(ParagraphProperties pProps, string key, string? value, List<string>? warnings = null)
    {
        if (value is null) return false;
        switch (key.ToLowerInvariant())
        {
            case "style" or "styleid":
                // CONSISTENCY(style-dual-key): Get exposes styleId as a
                // canonical readback key alongside the legacy `style`
                // (Round 2). Round 7+8 wired the alias trio on AddStyle
                // and SetStyle for /styles/X; the paragraph-level
                // Set surface was the missing link.
                // R7 deferred BT-4: warn (advisory, non-fatal) when the
                // style id does not exist in the styles part — opening
                // such a doc in Word shows a "style not found" badge.
                if (warnings != null && !StyleIdExists(value))
                    warnings.Add($"style '{value}' not found in styles part — will be referenced as-is");
                pProps.ParagraphStyleId = new ParagraphStyleId { Val = value };
                return true;
            case "stylename":
                // CONSISTENCY(style-dual-key): paragraph-level Set on
                // styleName resolves the display name through the styles
                // part — mirrors what Get reverses to expose styleName.
                // Falls back to using the value as styleId verbatim if no
                // matching display name is found (preserves the lenient-
                // input pattern used elsewhere).
                var resolved = ResolveStyleIdFromName(value);
                pProps.ParagraphStyleId = new ParagraphStyleId { Val = resolved ?? value };
                return true;
            case "align" or "alignment" or "jc":
                pProps.Justification = new Justification { Val = ParseJustification(value) };
                return true;
            case "firstlineindent":
                var indent = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                // Lenient input: accept "2cm", "0.5in", "18pt", or bare twips.
                // BUG-DUMP-NEGIND: ST_SignedTwipsMeasure — see SpacingConverter.
                indent.FirstLine = SpacingConverter.ParseWordSpacingSigned(value).ToString();
                indent.Hanging = null;
                return true;
            case "leftindent" or "indentleft" or "indent":
                var indentL = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                // CONSISTENCY(lenient-spacing): mirror Add — accept cm/in/pt/twips via SpacingConverter.
                // BUG-DUMP-NEGIND: signed.
                indentL.Left = SpacingConverter.ParseWordSpacingSigned(value).ToString();
                return true;
            case "rightindent" or "indentright":
                var indentR = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                // BUG-DUMP-NEGIND: signed.
                indentR.Right = SpacingConverter.ParseWordSpacingSigned(value).ToString();
                return true;
            case "hangingindent" or "hanging":
                var indentH = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                indentH.Hanging = SpacingConverter.ParseWordSpacing(value).ToString();
                indentH.FirstLine = null;
                return true;
            // CONSISTENCY(ind-char-units): CJK-convention character-unit
            // indents (recomputed by Word when font size changes). Mirror
            // AddStyle's char-unit handlers so dump→batch on Set paragraph
            // preserves the source's firstLineChars / leftChars / rightChars
            // / hangingChars attrs.
            case "firstlinechars":
            {
                var indFlc = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                indFlc.FirstLineChars = ParseHelpers.SafeParseInt(value, "firstLineChars");
                return true;
            }
            case "leftchars" or "startchars":
            {
                var indLc = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                indLc.LeftChars = ParseHelpers.SafeParseInt(value, "leftChars");
                return true;
            }
            case "rightchars" or "endchars":
            {
                var indRc = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                indRc.RightChars = ParseHelpers.SafeParseInt(value, "rightChars");
                return true;
            }
            case "hangingchars":
            {
                var indHc = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                indHc.HangingChars = ParseHelpers.SafeParseInt(value, "hangingChars");
                return true;
            }
            // Toggle props: always replace the element (don't `??=`) so an
            // existing `<w:foo w:val="false"/>` written by a previous Set or
            // by external tooling is correctly overridden when the new value
            // is true. With `??=` the val=false sticks and the toggle never
            // flips back to true (BUG-LT3).
            case "keepnext" or "keepwithnext":
                if (IsTruthy(value)) pProps.KeepNext = new KeepNext();
                else if (IsExplicitFalseAddOverride(value))
                    pProps.KeepNext = new KeepNext { Val = OnOffValue.FromBoolean(false) };
                else pProps.KeepNext = null;
                return true;
            case "keeplines" or "keeptogether":
                if (IsTruthy(value)) pProps.KeepLines = new KeepLines();
                else if (IsExplicitFalseAddOverride(value))
                    pProps.KeepLines = new KeepLines { Val = OnOffValue.FromBoolean(false) };
                else pProps.KeepLines = null;
                return true;
            case "pagebreakbefore":
                if (IsTruthy(value)) pProps.PageBreakBefore = new PageBreakBefore();
                else if (IsExplicitFalseAddOverride(value))
                    pProps.PageBreakBefore = new PageBreakBefore { Val = OnOffValue.FromBoolean(false) };
                else pProps.PageBreakBefore = null;
                return true;
            case "outlinelvl" or "outlinelevel":
            {
                // OOXML w:outlineLvl/@w:val is ST_DecimalNumber 0..9.
                // TypedAttributeFallback accepts the full ByteValue range so
                // val=10 silently produced schema-invalid XML. Mirror Add.
                if (!int.TryParse(value, System.Globalization.NumberStyles.Integer,
                        System.Globalization.CultureInfo.InvariantCulture, out var olvlS)
                    || olvlS < 0 || olvlS > 9)
                    throw new ArgumentException($"Invalid 'outlineLvl' value: '{value}'. Must be 0-9 (OOXML MaxInclusive=9).");
                pProps.OutlineLevel = new OutlineLevel { Val = olvlS };
                return true;
            }
            case "textalignment":
            {
                // OOXML ST_TextAlignment enum (auto/top/center/baseline/bottom).
                // TypedAttributeFallback wrote the raw string so "banana" 422'd
                // the file. Validate canonically.
                pProps.TextAlignment = new TextAlignment
                {
                    Val = value.ToLowerInvariant() switch
                    {
                        "auto"     => VerticalTextAlignmentValues.Auto,
                        "top"      => VerticalTextAlignmentValues.Top,
                        "center"   => VerticalTextAlignmentValues.Center,
                        "baseline" => VerticalTextAlignmentValues.Baseline,
                        "bottom"   => VerticalTextAlignmentValues.Bottom,
                        _ => throw new ArgumentException($"Invalid 'textAlignment' value: '{value}'. Valid: auto, top, center, baseline, bottom."),
                    },
                };
                return true;
            }
            case "wordwrap":
            {
                // OOXML CT_OnOff toggle (true = wrap, false = no wrap). The
                // generic fallback treated unrecognized strings as truthy/falsy
                // silently — reject anything that isn't a recognized boolean.
                var wwLower = value.ToLowerInvariant();
                if (wwLower is "true" or "1" or "yes" or "on")
                    pProps.WordWrap = new WordWrap();
                else if (wwLower is "false" or "0" or "no" or "off")
                    pProps.WordWrap = new WordWrap { Val = OnOffValue.FromBoolean(false) };
                else
                    throw new ArgumentException($"Invalid 'wordWrap' value: '{value}'. Valid: true, false, 1, 0, yes, no, on, off.");
                return true;
            }
            case "textboxtightwrap":
            {
                // OOXML ST_TextboxTightWrap. Curate so the SDK's lenient
                // EnumValue (which stores invalid strings as extension attrs)
                // can't slip past — invalid input must throw.
                pProps.TextBoxTightWrap = new TextBoxTightWrap
                {
                    Val = value.ToLowerInvariant() switch
                    {
                        "none"             => TextBoxTightWrapValues.None,
                        "alllines"         => TextBoxTightWrapValues.AllLines,
                        "firstandlastline" => TextBoxTightWrapValues.FirstAndLastLine,
                        "firstlineonly"    => TextBoxTightWrapValues.FirstLineOnly,
                        "lastlineonly"     => TextBoxTightWrapValues.LastLineOnly,
                        _ => throw new ArgumentException($"Invalid 'textboxTightWrap' value: '{value}'. Valid: none, allLines, firstAndLastLine, firstLineOnly, lastLineOnly."),
                    },
                };
                return true;
            }
            // R15 sweep: CT_OnOff paragraph toggles that previously fell through
            // to GenericXmlQuery.TryCreateTypedChild without strict boolean
            // validation. Same family as wordWrap above — invalid input silently
            // produced <w:kinsoku w:val="banana"/> (Get re-read as false).
            case "kinsoku":
            case "overflowpunct":
            case "toplinepunct":
            case "autospaceDE" or "autospacede":
            case "autospaceDN" or "autospacedn":
            case "adjustrightind":
            case "snaptogrid":
            case "mirrorindents":
            case "suppressoverlap":
            case "suppressautohyphens":
            case "suppresslinenumbers":
            {
                var onOffLower = value.ToLowerInvariant();
                bool? boolVal = onOffLower switch
                {
                    "true" or "1" or "yes" or "on"  => true,
                    "false" or "0" or "no" or "off" => false,
                    _ => null,
                };
                if (boolVal == null)
                    throw new ArgumentException($"Invalid '{key}' value: '{value}'. Valid: true, false, 1, 0, yes, no, on, off.");
                switch (key.ToLowerInvariant())
                {
                    case "kinsoku":
                        pProps.Kinsoku = boolVal.Value
                            ? new Kinsoku()
                            : new Kinsoku { Val = OnOffValue.FromBoolean(false) }; break;
                    case "overflowpunct":
                        pProps.OverflowPunctuation = boolVal.Value
                            ? new OverflowPunctuation()
                            : new OverflowPunctuation { Val = OnOffValue.FromBoolean(false) }; break;
                    case "toplinepunct":
                        pProps.TopLinePunctuation = boolVal.Value
                            ? new TopLinePunctuation()
                            : new TopLinePunctuation { Val = OnOffValue.FromBoolean(false) }; break;
                    case "autospacede":
                        pProps.AutoSpaceDE = boolVal.Value
                            ? new AutoSpaceDE()
                            : new AutoSpaceDE { Val = OnOffValue.FromBoolean(false) }; break;
                    case "autospacedn":
                        pProps.AutoSpaceDN = boolVal.Value
                            ? new AutoSpaceDN()
                            : new AutoSpaceDN { Val = OnOffValue.FromBoolean(false) }; break;
                    case "adjustrightind":
                        pProps.AdjustRightIndent = boolVal.Value
                            ? new AdjustRightIndent()
                            : new AdjustRightIndent { Val = OnOffValue.FromBoolean(false) }; break;
                    case "snaptogrid":
                        pProps.SnapToGrid = boolVal.Value
                            ? new SnapToGrid()
                            : new SnapToGrid { Val = OnOffValue.FromBoolean(false) }; break;
                    case "mirrorindents":
                        pProps.MirrorIndents = boolVal.Value
                            ? new MirrorIndents()
                            : new MirrorIndents { Val = OnOffValue.FromBoolean(false) }; break;
                    case "suppressoverlap":
                        pProps.SuppressOverlap = boolVal.Value
                            ? new SuppressOverlap()
                            : new SuppressOverlap { Val = OnOffValue.FromBoolean(false) }; break;
                    case "suppressautohyphens":
                        pProps.SuppressAutoHyphens = boolVal.Value
                            ? new SuppressAutoHyphens()
                            : new SuppressAutoHyphens { Val = OnOffValue.FromBoolean(false) }; break;
                    case "suppresslinenumbers":
                        pProps.SuppressLineNumbers = boolVal.Value
                            ? new SuppressLineNumbers()
                            : new SuppressLineNumbers { Val = OnOffValue.FromBoolean(false) }; break;
                }
                return true;
            }
            // fuzz-2: 'break=newPage' is the natural paragraph-context spelling
            // (mirrors section-context CONSISTENCY(section-type-alias) in
            // WordHandler.Set.Dispatch.cs:387). For a paragraph this maps to
            // pageBreakBefore=true; bare break=true also accepted.
            case "break":
                bool pbb = value.ToLowerInvariant() switch
                {
                    "newpage" or "page" or "nextpage" or "pagebreak" => true,
                    "none" or "" => false,
                    _ => IsTruthy(value)
                };
                if (pbb) pProps.PageBreakBefore = new PageBreakBefore();
                else pProps.PageBreakBefore = null;
                return true;
            case "widowcontrol" or "widoworphan":
                if (IsTruthy(value)) pProps.WidowControl = new WidowControl();
                else pProps.WidowControl = new WidowControl { Val = false };
                return true;
            case "contextualspacing" or "contextualSpacing":
                if (IsTruthy(value)) pProps.ContextualSpacing = new ContextualSpacing();
                else if (IsExplicitFalseAddOverride(value))
                    pProps.ContextualSpacing = new ContextualSpacing { Val = OnOffValue.FromBoolean(false) };
                else pProps.ContextualSpacing = null;
                return true;
            case "shading" or "shd" or "fill":
                pProps.Shading = ParseShadingValue(value);
                return true;
            case "spacebefore":
                var spacingBefore = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                // CONSISTENCY(ind-chars): "Nlines" suffix routes through the
                // hundredths-of-line attr (w:beforeLines), mirroring the
                // dedicated `spaceBeforeLines=` key. P1-7.
                // BUG-DUMP-SPACING-BOTHUNITS: a paragraph may legitimately carry
                // BOTH w:before (twips) AND w:beforeLines (hundredths-of-line) —
                // Word stores both and the dump emits both as spaceBefore +
                // spaceBeforeLines. The old within-axis clear (spaceBefore wipes
                // BeforeLines) dropped one of the pair on replay: a `set` with
                // both keys applied in sequence kept only the last writer, so a
                // table cell's `<w:spacing w:before="144" w:beforeLines="60"/>`
                // round-tripped as beforeLines-only — the cell shrank, the row
                // got shorter, and the form reflowed. AddParagraph already keeps
                // both (no clear); match it here. The "Nlines" suffix still routes
                // to the lines attr but no longer nulls the twips sibling — both
                // coexist exactly as the source had them.
                if (TryParseLinesSuffix(value, out var sblHundredths))
                    spacingBefore.BeforeLines = int.Parse(sblHundredths, System.Globalization.CultureInfo.InvariantCulture);
                else
                    spacingBefore.Before = SpacingConverter.ParseWordSpacing(value).ToString();
                return true;
            case "spaceafter":
                var spacingAfter = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                if (TryParseLinesSuffix(value, out var salHundredths))
                    spacingAfter.AfterLines = int.Parse(salHundredths, System.Globalization.CultureInfo.InvariantCulture);
                else
                    spacingAfter.After = SpacingConverter.ParseWordSpacing(value).ToString();
                return true;
            case "spacebeforelines":
                var spacingBL = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                spacingBL.BeforeLines = ParseHelpers.SafeParseInt(value, "spaceBeforeLines");
                return true;
            // BUG-DUMP-R44-4: auto-spacing on/off toggles (w:beforeAutospacing /
            // w:afterAutospacing). Mirror AddParagraph; round-trips the bool the
            // readback emits as spaceBeforeAuto / spaceAfterAuto.
            case "spacebeforeauto" or "beforeautospacing":
                var spacingBA = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                spacingBA.BeforeAutoSpacing = OnOffValue.FromBoolean(IsTruthy(value));
                return true;
            case "spaceafterauto" or "afterautospacing":
                var spacingAA = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                spacingAA.AfterAutoSpacing = OnOffValue.FromBoolean(IsTruthy(value));
                return true;
            case "spaceafterlines":
                var spacingAL = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                spacingAL.AfterLines = ParseHelpers.SafeParseInt(value, "spaceAfterLines");
                return true;
            case "linespacing":
                var spacingLine = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                var (lsTwips, lsIsMultiplier) = SpacingConverter.ParseWordLineSpacing(value);
                spacingLine.Line = lsTwips.ToString();
                spacingLine.LineRule = lsIsMultiplier ? LineSpacingRuleValues.Auto : LineSpacingRuleValues.Exact;
                return true;
            case "linerule" or "linespacingrule":
                // BUG-019: explicit override needed to distinguish AtLeast
                // from Exact — both serialize as "Npt" via SpacingConverter.
                var spacingRule = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                spacingRule.LineRule = ParseLineRule(value);
                return true;
            case "numId" or "numid":
                var numPr = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                var numIdVal = ParseHelpers.SafeParseInt(value, "numId");
                // numId=-1 is the OOXML negation marker that overrides inherited
                // numbering back to "no list"; treat it like 0 (skip existence check).
                if (numIdVal < -1)
                    throw new ArgumentException($"numId must be >= -1 (got {numIdVal}). Use numId=0 or numId=-1 to remove numbering.");
                if (numIdVal > 0)
                {
                    var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
                    var numExists = numbering?.Elements<NumberingInstance>()
                        .Any(n => n.NumberID?.Value == numIdVal) ?? false;
                    if (!numExists)
                        throw new ArgumentException(
                            $"numId={numIdVal} not found in /numbering. " +
                            "Create the num first (add /numbering --type num), or use numId=0 to remove numbering.");
                }
                numPr.NumberingId = new NumberingId { Val = numIdVal };
                return true;
            case "numLevel" or "numlevel" or "ilvl" or "listLevel" or "listlevel":
                var numPr2 = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                var ilvlSetVal = ParseHelpers.SafeParseInt(value, "numLevel");
                // BUG-R4B(BUG2): Word tolerates ilvl > 8 (real-world docs carry
                // w:ilvl w:val="12"); the old hard reject dropped the whole
                // paragraph on dump→replay. Clamp to the OOXML-legal range
                // [0,8] with a warning so the paragraph and its text survive.
                // Mirrors the HtmlPreview ilvl clamp (HtmlPreview.cs ~1889).
                if (ilvlSetVal < 0 || ilvlSetVal > 8)
                {
                    var clampedSet = Math.Clamp(ilvlSetVal, 0, 8);
                    warnings?.Add($"ilvl {ilvlSetVal} out of OOXML range 0..8 — clamped to {clampedSet}");
                    ilvlSetVal = clampedSet;
                }
                numPr2.NumberingLevelReference = new NumberingLevelReference { Val = ilvlSetVal };
                return true;
            case "pbdr.top" or "pbdr.bottom" or "pbdr.left" or "pbdr.right" or "pbdr.between" or "pbdr.bar" or "pbdr.all" or "pbdr":
            case "border.all" or "border" or "border.top" or "border.bottom" or "border.left" or "border.right" or "border.between" or "border.bar":
            case "border.color" or "border.sz" or "border.size" or "border.space" or "border.val" or "border.style":
                ApplyParagraphBorders(pProps, key, value);
                return true;
            // Reading direction: "rtl" enables right-to-left layout for Arabic
            // / Hebrew, "ltr" removes the bidi flag. Maps to <w:bidi/> in pPr.
            case "direction" or "dir" or "bidi":
                pProps.BiDi = ParseDirectionRtl(value) ? new BiDi() : null;
                return true;
            default:
                return false;
        }
    }

    // R17-consistency: align direction parsing across Word / PPT / Excel and
    // run-level rtl. Accepts rtl|righttoleft|right-to-left|true|1 (truthy),
    // ltr|lefttoright|left-to-right|false|0|"" (falsy), all case-insensitive.
    // Other values (yes/no/auto/2/...) throw — direction is a 2-value enum,
    // not an open boolean surface.
    private static bool ParseDirectionRtl(string value) => value.ToLowerInvariant() switch
    {
        "rtl" or "righttoleft" or "right-to-left" or "true" or "1" => true,
        "ltr" or "lefttoright" or "left-to-right" or "false" or "0" or "" => false,
        _ => throw new ArgumentException($"Invalid direction value: '{value}'. Valid values: rtl, ltr (also accepts true/false, 1/0, righttoleft/lefttoright, right-to-left/left-to-right; case-insensitive).")
    };

    /// <summary>
    /// Parse a w:numFmt value (page numbering / list numbering). Accepts the
    /// common Latin / Roman / Letter forms plus locale-specific scripts —
    /// notably 'arabicAlpha' / 'arabicAbjad' / 'hindiVowels' / 'hindiNumbers'
    /// for Arabic-script documents, and CJK ideographic forms for Chinese /
    /// Japanese / Korean. Falls through to the OOXML enum constructor so the
    /// full ECMA-376 set (chicago, persian, thaiCounting, etc.) round-trips.
    /// </summary>
    private static EnumValue<NumberFormatValues> ParseNumberFormat(string value)
    {
        var lower = value.ToLowerInvariant();
        return lower switch
        {
            "decimal" => NumberFormatValues.Decimal,
            "lowerroman" => NumberFormatValues.LowerRoman,
            "upperroman" => NumberFormatValues.UpperRoman,
            "lowerletter" or "loweralpha" => NumberFormatValues.LowerLetter,
            "upperletter" or "upperalpha" => NumberFormatValues.UpperLetter,
            "arabicalpha" => NumberFormatValues.ArabicAlpha,
            "arabicabjad" => NumberFormatValues.ArabicAbjad,
            "hindivowels" => NumberFormatValues.HindiVowels,
            "hindiconsonants" => NumberFormatValues.HindiConsonants,
            "hindinumbers" => NumberFormatValues.HindiNumbers,
            "hindicounting" => NumberFormatValues.HindiCounting,
            "thailetters" => NumberFormatValues.ThaiLetters,
            "thainumbers" => NumberFormatValues.ThaiNumbers,
            "thaicounting" => NumberFormatValues.ThaiCounting,
            "chinesecounting" => NumberFormatValues.ChineseCounting,
            "chinesecountingthousand" => NumberFormatValues.ChineseCountingThousand,
            "chineselegalsimplified" => NumberFormatValues.ChineseLegalSimplified,
            "japanesecounting" => NumberFormatValues.JapaneseCounting,
            "japaneselegal" => NumberFormatValues.JapaneseLegal,
            // OOXML SDK exposes only one Japanese-digital enum
            // (JapaneseDigitalTenThousand, ECMA-376 §17.18.59). Accept both
            // the short "ten" alias and the canonical OOXML wire name.
            "japanesedigitalten" or "japanesedigitaltenthousand"
                => NumberFormatValues.JapaneseDigitalTenThousand,
            "koreancounting" => NumberFormatValues.KoreanCounting,
            "koreanlegal" => NumberFormatValues.KoreanLegal,
            "koreandigital" => NumberFormatValues.KoreanDigital,
            "ideographdigital" => NumberFormatValues.IdeographDigital,
            "ideographtraditional" => NumberFormatValues.IdeographTraditional,
            "ideographzodiac" => NumberFormatValues.IdeographZodiac,
            "ideographenclosedcircle" => NumberFormatValues.IdeographEnclosedCircle,
            // Hebrew / Thai / Korean / English text and ordinal forms (ECMA-376
            // §17.18.59 ST_NumberFormat). Previously rejected — required for
            // Hebrew (hebrew1/2), Thai (bahtText), Japanese iroha ordering,
            // Korean ganada ordering, and English-language ordinal lists.
            "hebrew1" => NumberFormatValues.Hebrew1,
            "hebrew2" => NumberFormatValues.Hebrew2,
            "bahttext" => NumberFormatValues.BahtText,
            "dollartext" => NumberFormatValues.DollarText,
            "iroha" => NumberFormatValues.Iroha,
            "irohafullwidth" => NumberFormatValues.IrohaFullWidth,
            "ganada" => NumberFormatValues.Ganada,
            "ordinal" => NumberFormatValues.Ordinal,
            "cardinaltext" => NumberFormatValues.CardinalText,
            "ordinaltext" => NumberFormatValues.OrdinalText,
            "bullet" or "unordered" or "ul" => NumberFormatValues.Bullet,
            "decimalzero" => NumberFormatValues.DecimalZero,
            "decimalfullwidth" => NumberFormatValues.DecimalFullWidth,
            "decimalenclosedcircle" => NumberFormatValues.DecimalEnclosedCircle,
            "decimalenclosedcirclechinese" => NumberFormatValues.DecimalEnclosedCircleChinese,
            "decimalenclosedfullstop" => NumberFormatValues.DecimalEnclosedFullstop,
            "decimalenclosedparen" => NumberFormatValues.DecimalEnclosedParen,
            // Page-number and locale formats the Get readback emits verbatim
            // (pgNumType/@w:fmt, lvl/@w:numFmt) but Add/Set previously rejected,
            // so a dump→batch round-trip of any document using them aborted the
            // section/numbering add. numberInDash is the common one — "-1-"
            // page numbers in government/legal templates.
            "numberindash" => NumberFormatValues.NumberInDash,
            "hex" => NumberFormatValues.Hex,
            "chicago" => NumberFormatValues.Chicago,
            "decimalhalfwidth" => NumberFormatValues.DecimalHalfWidth,
            "decimalfullwidth2" => NumberFormatValues.DecimalFullWidth2,
            "aiueo" => NumberFormatValues.Aiueo,
            "aiueofullwidth" => NumberFormatValues.AiueoFullWidth,
            "chosung" => NumberFormatValues.Chosung,
            "koreandigital2" => NumberFormatValues.KoreanDigital2,
            "ideographzodiactraditional" => NumberFormatValues.IdeographZodiacTraditional,
            "ideographlegaltraditional" => NumberFormatValues.IdeographLegalTraditional,
            "taiwanesecounting" => NumberFormatValues.TaiwaneseCounting,
            "taiwanesecountingthousand" => NumberFormatValues.TaiwaneseCountingThousand,
            "taiwanesedigital" => NumberFormatValues.TaiwaneseDigital,
            "vietnamesecounting" => NumberFormatValues.VietnameseCounting,
            "russianlower" => NumberFormatValues.RussianLower,
            "russianupper" => NumberFormatValues.RussianUpper,
            "custom" => NumberFormatValues.Custom,
            "none" => NumberFormatValues.None,
            _ => throw new ArgumentException(
                $"Unknown numbering format '{value}'. Common values: decimal, lowerRoman, upperRoman, "
                + "lowerLetter, upperLetter, bullet. Locale-specific: hindiNumbers, hindiVowels, arabicAlpha, "
                + "arabicAbjad, thaiCounting, chineseCounting, japaneseCounting, koreanCounting, "
                + "ideographDigital, none.")
        };
    }


    private static void ApplyParagraphBorders(ParagraphProperties pProps, string key, string value)
    {
        var borders = pProps.ParagraphBorders;
        if (borders == null)
        {
            borders = new ParagraphBorders();
            pProps.ParagraphBorders = borders; // typed setter maintains CT_PPr schema order
        }
        var (style, size, color, space) = ParseBorderValue(value);
        var sf = ParseBorderShadowFrame(value);
        // BUG-DUMP-R41-2: harvest the theme key=val tail once; passed to every
        // MakeBorder below so w:themeColor/Shade/Tint round-trip.
        var (_, bTheme) = ExtractThemeTail(value);

        switch (key.ToLowerInvariant())
        {
            case "pbdr.all" or "pbdr" or "border.all" or "border":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                borders.BetweenBorder = MakeBorder<BetweenBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "pbdr.top" or "border.top":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "pbdr.bottom" or "border.bottom":
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "pbdr.left" or "border.left":
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "pbdr.right" or "border.right":
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "pbdr.between" or "border.between":
                borders.BetweenBorder = MakeBorder<BetweenBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "pbdr.bar" or "border.bar":
                borders.BarBorder = MakeBorder<BarBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            default:
                // Companion form: border.color, border.sz, border.space, border.val
                // — update the attribute on every already-stored edge so users can
                // author `border=single border.color=FF border.sz=12` and have all
                // three land. Lazily creates a `single` border on each missing
                // edge so the companion props aren't a no-op when called first.
                ApplyParagraphBorderCompanion(borders, key.ToLowerInvariant(), value);
                break;
        }
    }

    private static void ApplyParagraphBorderCompanion(ParagraphBorders borders, string key, string value)
    {
        // key shape: border.<attr> or pbdr.<attr> — must NOT be an edge name.
        var parts = key.Split('.');
        if (parts.Length != 2) return;
        var attr = parts[1];
        if (attr is "all" or "top" or "bottom" or "left" or "right" or "between" or "bar") return;
        var sides = new BorderType[]
        {
            borders.TopBorder ??= new TopBorder { Val = BorderValues.Single },
            borders.BottomBorder ??= new BottomBorder { Val = BorderValues.Single },
            borders.LeftBorder ??= new LeftBorder { Val = BorderValues.Single },
            borders.RightBorder ??= new RightBorder { Val = BorderValues.Single },
        };
        foreach (var b in sides)
            SetBorderAttr(b, attr, value);
    }

    private static void SetBorderAttr(BorderType b, string attr, string value)
    {
        switch (attr)
        {
            case "sz":
            case "size":
                b.Size = (uint)ParseHelpers.SafeParseInt(value, attr);
                break;
            case "color":
                var (cHex, _) = ParseHelpers.SanitizeColorForOoxml(value);
                b.Color = cHex;
                break;
            case "space":
                b.Space = (uint)ParseHelpers.SafeParseInt(value, attr);
                break;
            case "val":
            case "style":
                b.Val = new EnumValue<BorderValues>(new BorderValues(value));
                break;
        }
    }

    private static void ApplyStyleParagraphBorders(StyleParagraphProperties spPr, string key, string value)
    {
        var borders = spPr.GetFirstChild<ParagraphBorders>();
        if (borders == null)
        {
            // CT_PPr places <w:pBdr> EARLY (after numPr/suppressLineNumbers,
            // before tabs/spacing/ind/jc) — the prior hand-rolled "after
            // Indentation, before Shading" anchor was wrong and landed pBdr
            // after jc, which strict validators reject. Append, then let
            // SchemaOrder hoist it to the SDK-authoritative slot.
            borders = spPr.AppendChild(new ParagraphBorders());
            Core.SchemaOrder.Place(spPr, borders);
        }
        var (style, size, color, space) = ParseBorderValue(value);
        var sf = ParseBorderShadowFrame(value);
        // BUG-DUMP-R41-2: harvest the theme key=val tail once; passed to every
        // MakeBorder below so w:themeColor/Shade/Tint round-trip.
        var (_, bTheme) = ExtractThemeTail(value);

        switch (key.ToLowerInvariant())
        {
            case "pbdr.all" or "pbdr" or "border.all" or "border":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                borders.BetweenBorder = MakeBorder<BetweenBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "pbdr.top" or "border.top":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "pbdr.bottom" or "border.bottom":
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "pbdr.left" or "border.left":
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "pbdr.right" or "border.right":
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "pbdr.between" or "border.between":
                borders.BetweenBorder = MakeBorder<BetweenBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "pbdr.bar" or "border.bar":
                borders.BarBorder = MakeBorder<BarBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            default:
                ApplyParagraphBorderCompanion(borders, key.ToLowerInvariant(), value);
                break;
        }
    }

    private static void ApplyTableBorders(TableProperties tblPr, string key, string value)
    {
        var borders = tblPr.TableBorders ?? tblPr.AppendChild(new TableBorders());
        var (style, size, color, space) = ParseBorderValue(value);
        var sf = ParseBorderShadowFrame(value);
        // BUG-DUMP-R41-2: harvest the theme key=val tail once; passed to every
        // MakeBorder below so w:themeColor/Shade/Tint round-trip.
        var (_, bTheme) = ExtractThemeTail(value);

        switch (key.ToLowerInvariant())
        {
            case "border.all" or "border":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                borders.InsideHorizontalBorder = MakeBorder<InsideHorizontalBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                borders.InsideVerticalBorder = MakeBorder<InsideVerticalBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "border.top":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "border.bottom":
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "border.left":
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "border.right":
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "border.insideh" or "border.horizontal":
                borders.InsideHorizontalBorder = MakeBorder<InsideHorizontalBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "border.insidev" or "border.vertical":
                borders.InsideVerticalBorder = MakeBorder<InsideVerticalBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            default:
                // Sub-property form: border.<edge>.<attr> — set a single
                // attribute on the existing edge border (caller must have
                // set border.<edge> first; we lazily create a default
                // `single` border if absent). Mirrors ApplyCellBorderSubProperty.
                ApplyTableBorderSubProperty(borders, key.ToLowerInvariant(), value);
                break;
        }
    }

    private static void ApplyTableBorderSubProperty(TableBorders borders, string key, string value)
    {
        var parts = key.Split('.');
        if (parts.Length != 3 || parts[0] != "border") return;
        var edge = parts[1];
        var attr = parts[2];
        BorderType? b = edge switch
        {
            "top"    => borders.TopBorder    ??= new TopBorder    { Val = BorderValues.Single },
            "bottom" => borders.BottomBorder ??= new BottomBorder { Val = BorderValues.Single },
            "left" or "start" => borders.LeftBorder  ??= new LeftBorder  { Val = BorderValues.Single },
            "right" or "end"  => borders.RightBorder ??= new RightBorder { Val = BorderValues.Single },
            "insideh" or "horizontal" => borders.InsideHorizontalBorder ??= new InsideHorizontalBorder { Val = BorderValues.Single },
            "insidev" or "vertical"   => borders.InsideVerticalBorder   ??= new InsideVerticalBorder   { Val = BorderValues.Single },
            _ => null,
        };
        if (b == null) return;
        switch (attr)
        {
            case "sz":
            case "size":
                b.Size = (uint)ParseHelpers.SafeParseInt(value, key);
                break;
            case "color":
                var (cHex, _) = ParseHelpers.SanitizeColorForOoxml(value);
                b.Color = cHex;
                break;
            case "space":
                b.Space = (uint)ParseHelpers.SafeParseInt(value, key);
                break;
            case "val":
            case "style":
                b.Val = new EnumValue<BorderValues>(new BorderValues(value));
                break;
        }
    }

    /// <summary>
    /// CT_TcPr child schema order. Used by InsertTcPrChildInOrder to insert
    /// new tcPr children at their schema position rather than the tail.
    /// Children whose type isn't on this list (mc:AlternateContent and
    /// extensions, for instance) are tolerated — they sort to the end via
    /// the IndexOf == -1 sentinel.
    /// </summary>
    private static readonly Type[] s_tcPrChildOrder =
    [
        typeof(ConditionalFormatStyle),
        typeof(TableCellWidth),
        typeof(GridSpan),
        typeof(HorizontalMerge),
        typeof(VerticalMerge),
        typeof(TableCellBorders),
        typeof(Shading),
        typeof(NoWrap),
        typeof(TableCellMargin),
        typeof(TextDirection),
        typeof(TableCellFitText),
        typeof(TableCellVerticalAlignment),
        typeof(HideMark),
        // headers/cellIns/cellDel/cellMerge/tcPrChange follow but are rare
        // enough that we let the SDK's own setters handle them; they get
        // sentinel positions (-1) and end up at the tail, which is correct
        // when nothing else past tcPr has been written.
    ];

    private static void InsertTcPrChildInOrder(TableCellProperties tcPr, OpenXmlElement child)
    {
        var targetIdx = Array.IndexOf(s_tcPrChildOrder, child.GetType());
        if (targetIdx < 0)
        {
            tcPr.AppendChild(child);
            return;
        }
        foreach (var sibling in tcPr.ChildElements)
        {
            var sibIdx = Array.IndexOf(s_tcPrChildOrder, sibling.GetType());
            if (sibIdx > targetIdx)
            {
                tcPr.InsertBefore(child, sibling);
                return;
            }
        }
        tcPr.AppendChild(child);
    }

    private static bool ApplyCellBorders(TableCellProperties tcPr, string key, string value)
    {
        // Validate the key shape before mutating tcPr — an unknown border.*
        // key (e.g. "borderStyle", or a length-2 split that ApplyCellBorderSubProperty
        // would silently drop) must surface as unsupported rather than leaving
        // an empty <w:tcBorders/> element behind.
        var keyLower = key.ToLowerInvariant();
        bool isWhole = keyLower is "border.all" or "border" or "border.top"
            or "border.bottom" or "border.left" or "border.right"
            or "border.tl2br" or "border.tr2bl";
        bool isSubProp = false;
        if (!isWhole)
        {
            var parts = keyLower.Split('.');
            if (parts.Length == 3 && parts[0] == "border")
            {
                var edge = parts[1];
                var attr = parts[2];
                bool edgeOk = edge is "top" or "bottom" or "left" or "start"
                    or "right" or "end" or "tl2br" or "tr2bl";
                bool attrOk = attr is "sz" or "size" or "color" or "space" or "val" or "style";
                isSubProp = edgeOk && attrOk;
            }
        }
        if (!isWhole && !isSubProp) return false;

        // CT_TcPr child sequence is strict: cnfStyle → tcW → gridSpan →
        // hMerge → vMerge → tcBorders → shd → noWrap → tcMar →
        // textDirection → tcFitText → vAlign → hideMark → ... → tcPrChange.
        // Plain AppendChild lands tcBorders at the tail, after shd/vAlign/
        // tcMar that earlier setter calls already wrote, producing
        // Sch_UnexpectedElementContentExpectingComplex on tcBorders. Insert
        // before the first existing sibling that should come after tcBorders.
        var borders = tcPr.TableCellBorders;
        if (borders == null)
        {
            borders = new TableCellBorders();
            InsertTcPrChildInOrder(tcPr, borders);
        }
        var (style, size, color, space) = ParseBorderValue(value);
        var sf = ParseBorderShadowFrame(value);
        // BUG-DUMP-R41-2: harvest the theme key=val tail once; passed to every
        // MakeBorder below so w:themeColor/Shade/Tint round-trip.
        var (_, bTheme) = ExtractThemeTail(value);

        switch (keyLower)
        {
            case "border.all" or "border":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "border.top":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "border.bottom":
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "border.left":
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "border.right":
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "border.tl2br":
                borders.TopLeftToBottomRightCellBorder = MakeBorder<TopLeftToBottomRightCellBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            case "border.tr2bl":
                borders.TopRightToBottomLeftCellBorder = MakeBorder<TopRightToBottomLeftCellBorder>(style, size, color, space, sf.shadow, sf.frame, theme: bTheme);
                break;
            default:
                // Sub-property form: border.<edge>.<attr> — set an attribute on
                // the existing edge border (caller must have set border.<edge>
                // first; we lazily create a default `single` border if absent).
                ApplyCellBorderSubProperty(borders, keyLower, value);
                break;
        }
        return true;
    }

    private static void ApplyCellBorderSubProperty(TableCellBorders borders, string key, string value)
    {
        // key shape: border.<edge>.<attr> — e.g. border.top.sz, border.left.color
        var parts = key.Split('.');
        if (parts.Length != 3) return;
        var edge = parts[1];
        var attr = parts[2];
        BorderType? b = edge switch
        {
            "top" => borders.TopBorder ??= new TopBorder { Val = BorderValues.Single },
            "bottom" => borders.BottomBorder ??= new BottomBorder { Val = BorderValues.Single },
            "left" or "start" => borders.LeftBorder ??= new LeftBorder { Val = BorderValues.Single },
            "right" or "end" => borders.RightBorder ??= new RightBorder { Val = BorderValues.Single },
            "tl2br" => borders.TopLeftToBottomRightCellBorder ??= new TopLeftToBottomRightCellBorder { Val = BorderValues.Single },
            "tr2bl" => borders.TopRightToBottomLeftCellBorder ??= new TopRightToBottomLeftCellBorder { Val = BorderValues.Single },
            _ => null,
        };
        if (b == null) return;
        switch (attr)
        {
            case "sz":
            case "size":
                b.Size = (uint)ParseHelpers.SafeParseInt(value, key);
                break;
            case "color":
                var (cHex, _) = ParseHelpers.SanitizeColorForOoxml(value);
                b.Color = cHex;
                break;
            case "space":
                b.Space = (uint)ParseHelpers.SafeParseInt(value, key);
                break;
            case "val":
            case "style":
                b.Val = new EnumValue<BorderValues>(new BorderValues(value));
                break;
        }
    }

    /// <summary>
    /// Apply gradient fill to a Word table cell using mc:AlternativeContent with w14:gradFill.
    /// Fallback is a solid shading with the start color.
    /// </summary>
    private static void ApplyCellGradient(TableCellProperties tcPr, string startColor, string endColor, int angleDeg)
    {
        // Sanitize colors: strip 8-char RRGGBBAA to 6-char RGB (w14:srgbClr requires 6 chars)
        var (startRgb, _) = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(startColor);
        var (endRgb, _) = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(endColor);

        // Remove existing shading/gradient
        RemoveCellGradient(tcPr);
        tcPr.Shading?.Remove();

        // Set fallback solid fill
        tcPr.Shading = new Shading { Val = ShadingPatternValues.Clear, Fill = startRgb };

        // Build w14:gradFill XML via raw OpenXml
        var w14Ns = "http://schemas.microsoft.com/office/word/2010/wordml";
        var mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        // Convert angle to OOXML 60000ths of a degree
        var angleOoxml = angleDeg * 60000;

        var acElement = new OpenXmlUnknownElement("mc", "AlternateContent", mcNs);
        acElement.InnerXml = $@"<mc:Choice xmlns:mc=""{mcNs}"" xmlns:w14=""{w14Ns}"" Requires=""w14"">
    <w14:gradFill>
      <w14:gsLst>
        <w14:gs w14:pos=""0"">
          <w14:srgbClr w14:val=""{startRgb}""/>
        </w14:gs>
        <w14:gs w14:pos=""100000"">
          <w14:srgbClr w14:val=""{endRgb}""/>
        </w14:gs>
      </w14:gsLst>
      <w14:lin w14:ang=""{angleOoxml}"" w14:scaled=""1""/>
    </w14:gradFill>
  </mc:Choice>";

        tcPr.AppendChild(acElement);
    }

    /// <summary>
    /// Remove any existing gradient mc:AlternateContent from table cell properties.
    /// </summary>
    private static void RemoveCellGradient(TableCellProperties tcPr)
    {
        var mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        var existing = tcPr.ChildElements
            .Where(e => e.LocalName == "AlternateContent" && e.NamespaceUri == mcNs)
            .ToList();
        foreach (var e in existing) e.Remove();
    }

    /// <summary>
    /// Parse twips from a string with optional unit suffix: "1.5cm", "0.5in", "36pt", or raw twips.
    /// 1 inch = 1440 twips, 1 cm = 567 twips, 1 pt = 20 twips.
    /// </summary>
    private static TablePositionProperties EnsureTablePositionProperties(TableProperties tblPr)
    {
        var tpp = tblPr.GetFirstChild<TablePositionProperties>();
        if (tpp == null)
        {
            // BUG-DUMP-TBLPPR-HORZANCHOR: do NOT bake in default
            // vertAnchor/horzAnchor=page. Both attributes are OPTIONAL in
            // ST_TblPPr (Word's default is text/column-relative). Hard-defaulting
            // them meant a dump→batch of a floating table whose source set only
            // some tblp.* attrs (e.g. vertAnchor=text + leftFromText, no
            // horzAnchor) gained a spurious horzAnchor="page" — which re-anchored
            // the table to the page's left edge, so the following paragraph (a
            // table's "Notes" caption) wrapped BESIDE the half-width table instead
            // of flowing below it (visible layout corruption). Create the element
            // empty; each anchor is set only when the caller explicitly provides
            // tblp.horzAnchor / tblp.vertAnchor, matching the source faithfully.
            tpp = new TablePositionProperties();
            // CONSISTENCY(tblpr-schema-order): tblpPr is rank 1.
            InsertTblPrChildInOrder(tblPr, tpp);
        }
        return tpp;
    }

    internal static uint ParseTwips(string value)
    {
        value = value.Trim();
        // Twips back OOXML length attributes that are uint in the schema (pgSz/@w:w,
        // pgMar/@w:top, etc.). Negative inputs would wrap silently on the (uint) cast
        // below — reject them up front in every unit branch with a uniform message.
        // The integer branch already rejects negatives via SafeParseUint.
        if (value.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseHelpers.SafeParseDouble(value[..^2], "twips (cm)");
            if (num < 0)
                throw new ArgumentException($"length must be non-negative, got {num}cm.");
            return (uint)Math.Round(num * 1440.0 / 2.54);
        }
        if (value.EndsWith("in", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseHelpers.SafeParseDouble(value[..^2], "twips (in)");
            if (num < 0)
                throw new ArgumentException($"length must be non-negative, got {num}in.");
            return (uint)Math.Round(num * 1440);
        }
        if (value.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseHelpers.SafeParseDouble(value[..^2], "twips (pt)");
            if (num < 0)
                throw new ArgumentException($"length must be non-negative, got {num}pt.");
            return (uint)Math.Round(num * 20);
        }
        // BUG-R2b: accept the "dxa" suffix that Get now emits for colWidths/width
        // so a dxa-qualified value round-trips back through Set. dxa is the OOXML
        // twip unit, so it strips to a bare integer (no scaling).
        // BUG-DUMP-FRACTWIPS: a source <w:col w:w="4521.5"> (an authoring-tool
        // artifact — twips are integer in the schema, but Word tolerates and
        // rounds fractional widths) made `dump --format batch` emit "4521.5",
        // which Set's integer-only parse then REJECTED — officecli could not
        // round-trip a value it itself produced, dropping the multi-column
        // layout. Accept fractional twips and round, matching the cm/in/pt
        // branches above (Word rounds the same way).
        if (value.EndsWith("dxa", StringComparison.OrdinalIgnoreCase))
            return RoundNonNegativeTwips(value[..^3]);
        return RoundNonNegativeTwips(value);
    }

    // Parse a bare twip value, tolerating a fractional component (rounded to the
    // nearest integer twip) so a source's fractional length round-trips. A plain
    // integer string still parses exactly; only a non-integer takes the rounding
    // path. Negative values are rejected, matching the unit branches.
    private static uint RoundNonNegativeTwips(string s)
    {
        s = s.Trim();
        if (uint.TryParse(s, System.Globalization.CultureInfo.InvariantCulture, out var exact))
            return exact;
        var num = ParseHelpers.SafeParseDouble(s, "twips");
        if (num < 0)
            throw new ArgumentException($"length must be non-negative, got {num}.");
        return (uint)Math.Round(num);
    }
}
