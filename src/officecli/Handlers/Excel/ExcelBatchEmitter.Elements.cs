// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

// PR2 element emits: tables, conditional formats, data validations, cell
// comments, charts, sparklines. Each transcribes the indexed Get node
// (/Sheet/table[N], /Sheet/cf[N], ...) into the matching `add` vocabulary.
// All run AFTER the value baseline + style layer so referenced ranges hold
// real data by the time the element replays.
public static partial class ExcelBatchEmitter
{
    private static void EmitSheetElements(ExcelHandler xl, string sheetName,
        List<BatchItem> items, List<UnsupportedWarning> warnings)
    {
        var sheetPath = "/" + sheetName;
        var counts = xl.GetDumpElementCounts(sheetName);

        EmitTables(xl, sheetPath, counts.Tables, items, warnings);
        EmitConditionalFormats(xl, sheetPath, counts.Cfs, items, warnings);
        EmitValidations(xl, sheetPath, counts.Validations, items, warnings);
        EmitComments(xl, sheetPath, counts.Comments, items, warnings);
        // Charts replay LAST (see EmitChartsPass) — a series can reference a
        // range on a sheet emitted later in the workbook; adding the chart
        // before that sheet's data exists leaves numRef with an empty numCache
        // (re-dump shows an empty series). Deferred to the final pass.
        EmitSparklines(xl, sheetPath, counts.Sparklines, items, warnings);
        var drawingCounts = xl.GetDumpDrawingCounts(sheetName);
        EmitPictures(xl, sheetName, sheetPath, drawingCounts.Pictures, items, warnings);
        EmitShapes(xl, sheetName, sheetPath, items, warnings);
        EmitOles(xl, sheetName, sheetPath, items, warnings);
        EmitAutoFilterCriteria(xl, sheetPath, items, warnings);
    }

    // ==================== AutoFilter criteria ====================

    // The sheet-settings pass already emits `set autoFilter=<range>` (the bare
    // ref). Per-column filter criteria live in <filterColumn> children that
    // `set` cannot express, so when any exist we replay them through
    // `add --type autofilter` (which accepts range + criteriaN.OP and rebuilds
    // the filterColumns idempotently over the range the set row just created).
    private static void EmitAutoFilterCriteria(ExcelHandler xl, string sheetPath,
        List<BatchItem> items, List<UnsupportedWarning> warnings)
    {
        DocumentNode af;
        try { af = xl.Get($"{sheetPath}/autofilter"); }
        catch { return; }
        if (!af.Format.TryGetValue("range", out var rv) || rv is not string range || range.Length == 0)
            return;
        var criteria = af.Format
            .Where(kv => kv.Key.StartsWith("criteria", StringComparison.OrdinalIgnoreCase)
                && kv.Value is string s && s.Length > 0)
            .ToList();
        if (criteria.Count == 0) return;
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["range"] = range };
        foreach (var (k, v) in criteria) props[k] = (string)v!;
        items.Add(new BatchItem { Command = "add", Parent = sheetPath, Type = "autofilter", Props = props });
    }

    /// <summary>
    /// Emit all charts across every sheet AFTER every sheet's data baseline
    /// has been imported. Cross-sheet series references (Sheet2!$B$2:$B$5 on a
    /// chart living on Sheet1) need the referenced sheet's cells to already
    /// exist so the numCache is populated on replay.
    /// </summary>
    internal static void EmitChartsPass(ExcelHandler xl, string sheetName,
        List<BatchItem> items, List<UnsupportedWarning> warnings)
    {
        var sheetPath = "/" + sheetName;
        var counts = xl.GetDumpElementCounts(sheetName);
        EmitCharts(xl, sheetPath, counts.Charts, items, warnings);
        EmitChartExCharts(xl, sheetName, sheetPath, items, warnings);
    }

    // ==================== Tables ====================

    private static void EmitTables(ExcelHandler xl, string sheetPath, int count,
        List<BatchItem> items, List<UnsupportedWarning> warnings)
    {
        for (int i = 1; i <= count; i++)
        {
            DocumentNode t;
            try { t = xl.Get($"{sheetPath}/table[{i}]"); }
            catch (Exception ex)
            {
                warnings.Add(new UnsupportedWarning("table", $"{sheetPath}/table[{i}]", ex.Message));
                continue;
            }
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            CopyString(t, "ref", props, "ref");
            CopyString(t, "name", props, "name");
            CopyString(t, "displayName", props, "displayName");
            CopyString(t, "style", props, "style");
            // A table with no tableStyleInfo renders plain; AddTable defaults
            // to TableStyleMedium2, so an omitted style key restyled the
            // replayed table (visible SSIM drift on externally-authored
            // files). Emit the explicit `none` alias to preserve plainness.
            if (!props.ContainsKey("style"))
                props["style"] = "none";
            // headerRow defaults true on Add; totalRow defaults false — emit
            // only the non-default direction to keep round-trip idempotent.
            if (t.Format.TryGetValue("headerRow", out var hr) && hr is bool hrB && !hrB)
                props["headerRow"] = "false";
            if (t.Format.TryGetValue("totalRow", out var tr) && tr is bool trB && trB)
            {
                props["totalRow"] = "true";
                // The OOXML ref INCLUDES the totals row, but AddTable expands
                // its ref by one row when totalRow=true — emitting the raw ref
                // grew the table a row per replay cycle. Emit the data-only ref.
                if (props.TryGetValue("ref", out var fullRef) && fullRef.Contains(':'))
                {
                    var rp = fullRef.Split(':');
                    var m2 = System.Text.RegularExpressions.Regex.Match(rp[1], @"^([A-Za-z]+)(\d+)$");
                    if (m2.Success && int.TryParse(m2.Groups[2].Value, out var lastRow) && lastRow > 1)
                        props["ref"] = $"{rp[0]}:{m2.Groups[1].Value}{lastRow - 1}";
                }
                // Per-column aggregation tokens; without them replay falls
                // back to AddTable's label+SUM default (average became sum).
                var tokens = xl.GetDumpTableTotalsTokens(sheetPath.TrimStart('/'), i);
                if (!string.IsNullOrEmpty(tokens))
                    props["totalsRowFunction"] = tokens!;
            }
            // AddTable defaults ShowRowStripes=true, ShowColumnStripes/
            // ShowFirstColumn/ShowLastColumn=false. bandedRows is thus the one
            // with an ON default: emit it only when FALSE (the non-default).
            // The other three keep the emit-when-true form.
            if (t.Format.TryGetValue("bandedRows", out var brv) && brv is bool brB && !brB)
                props["bandedRows"] = "false";
            CopyBool(t, "bandedCols", props, "bandedCols");
            CopyBool(t, "firstCol", props, "firstCol");
            CopyBool(t, "lastCol", props, "lastCol");
            if (!props.ContainsKey("ref"))
            {
                warnings.Add(new UnsupportedWarning("table", t.Path ?? sheetPath, "table has no ref; skipped"));
                continue;
            }
            items.Add(new BatchItem { Command = "add", Parent = sheetPath, Type = "table", Props = props });
        }
    }

    // ==================== Conditional formats ====================

    // Get canonical type → Add element type. Every entry's prop mapping is
    // verified against ExcelHandler.Add.Cf.cs consumption sites.
    private static void EmitConditionalFormats(ExcelHandler xl, string sheetPath, int count,
        List<BatchItem> items, List<UnsupportedWarning> warnings)
    {
        // Iterate cf RULES, not cf elements. A <conditionalFormatting> element
        // can hold several <cfRule> children (Excel stacks rules on one range);
        // the cf[N] Get index space and `count` are per-element and only expose
        // rule[0], so emitting per-element dropped rule[1..N] silently. The
        // per-rule enumeration surfaces each rule as its own add-command.
        List<DocumentNode> cfNodes;
        try { cfNodes = xl.GetDumpCfRuleNodes(sheetPath.TrimStart('/')); }
        catch (Exception ex)
        {
            warnings.Add(new UnsupportedWarning("conditionalformatting", sheetPath, ex.Message));
            return;
        }
        foreach (var cf in cfNodes)
        {
            var type = cf.Format.TryGetValue("type", out var tv) ? tv as string ?? "" : "";
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            CopyString(cf, "ref", props, "ref");
            if (!props.ContainsKey("ref"))
            {
                warnings.Add(new UnsupportedWarning("conditionalformatting", cf.Path ?? sheetPath, "cf rule has no ref; skipped"));
                continue;
            }
            // stopIfTrue applies to every CF rule type; the Add path honors it
            // uniformly via ApplyStopIfTrue, so carry it for all branches.
            CopyBool(cf, "stopIfTrue", props, "stopIfTrue");

            string addType;
            switch (type)
            {
                case "dataBar":
                    addType = "databar";
                    CopyString(cf, "color", props, "color");
                    if (cf.Format.TryGetValue("showValue", out var sv) && sv is bool svB && !svB)
                        props["showValue"] = "false";
                    CopyValue(cf, "minLength", props, "minLength");
                    CopyValue(cf, "maxLength", props, "maxLength");
                    CopyString(cf, "direction", props, "direction");
                    CopyString(cf, "negativeColor", props, "negativeColor");
                    CopyString(cf, "axisColor", props, "axisColor");
                    CopyString(cf, "axisPosition", props, "axisPosition");
                    // Explicit numeric min/max bounds; omitting them lets the
                    // Add path fall back to autoMin/autoMax.
                    CopyString(cf, "min", props, "min");
                    CopyString(cf, "max", props, "max");
                    break;
                case "colorScale":
                    addType = "colorscale";
                    CopyString(cf, "minColor", props, "mincolor");
                    CopyString(cf, "maxColor", props, "maxcolor");
                    CopyString(cf, "midColor", props, "midcolor");
                    CopyString(cf, "midpoint", props, "midpoint");
                    break;
                case "iconSet":
                    addType = "iconset";
                    CopyString(cf, "iconset", props, "iconset");
                    CopyBool(cf, "reverse", props, "reverse");
                    if (cf.Format.TryGetValue("showValue", out var isv) && isv is bool isvB && !isvB)
                        props["showvalue"] = "false";
                    break;
                case "formula":
                    addType = "formulacf";
                    CopyString(cf, "formula", props, "formula");
                    CopyDxfStyle(cf, props);
                    break;
                case "cellIs":
                    addType = "cellis";
                    CopyString(cf, "operator", props, "operator");
                    CopyString(cf, "value", props, "value");
                    CopyString(cf, "value2", props, "value2");
                    CopyDxfStyle(cf, props);
                    break;
                case "topN":
                    addType = "topn";
                    CopyValue(cf, "rank", props, "rank");
                    CopyBool(cf, "percent", props, "percent");
                    CopyBool(cf, "bottom", props, "bottom");
                    CopyDxfStyle(cf, props);
                    break;
                case "aboveAverage":
                    addType = "aboveaverage";
                    if (cf.Format.TryGetValue("aboveAverage", out var aa) && aa is bool aaB && !aaB)
                        props["above"] = "false";
                    CopyValue(cf, "stdDev", props, "stdDev");
                    CopyBool(cf, "equalAverage", props, "equalAverage");
                    CopyDxfStyle(cf, props);
                    break;
                case "uniqueValues":
                    addType = "uniquevalues";
                    CopyDxfStyle(cf, props);
                    break;
                case "duplicateValues":
                    addType = "duplicatevalues";
                    CopyDxfStyle(cf, props);
                    break;
                case "containsText":
                    addType = "containstext";
                    CopyString(cf, "text", props, "text");
                    CopyDxfStyle(cf, props);
                    break;
                // Blank/error/text-operator family — the add vocabulary routes
                // these via `--type cf --prop type=<name>` (AddCf → cfextended).
                case "containsBlanks":
                case "notContainsBlanks":
                case "containsErrors":
                case "notContainsErrors":
                    addType = "cf";
                    props["type"] = type;
                    CopyDxfStyle(cf, props);
                    break;
                case "beginsWith":
                case "endsWith":
                    addType = "cf";
                    props["type"] = type;
                    CopyString(cf, "text", props, "text");
                    CopyDxfStyle(cf, props);
                    break;
                case "notContainsText":
                    addType = "cf";
                    props["type"] = "notcontains";
                    CopyString(cf, "text", props, "text");
                    CopyDxfStyle(cf, props);
                    break;
                case "timePeriod":
                    addType = "dateoccurring";
                    CopyString(cf, "period", props, "period");
                    CopyDxfStyle(cf, props);
                    break;
                default:
                    warnings.Add(new UnsupportedWarning("conditionalformatting", cf.Path ?? sheetPath,
                        $"cf rule type '{type}' has no add vocabulary; skipped"));
                    continue;
            }
            items.Add(new BatchItem { Command = "add", Parent = sheetPath, Type = addType, Props = props });
        }
    }

    // dxf-resolved style facets PopulateCfNodeFromDxf surfaces; the dxf-backed
    // Add paths (formulacf/cellis/topn/...) consume the same fill/font.* keys.
    private static void CopyDxfStyle(DocumentNode cf, Dictionary<string, string> props)
    {
        CopyString(cf, "fill", props, "fill");
        CopyBool(cf, "font.bold", props, "font.bold");
        CopyBool(cf, "font.italic", props, "font.italic");
        CopyBool(cf, "font.strike", props, "font.strike");
        CopyString(cf, "font.underline", props, "font.underline");
        CopyString(cf, "font.size", props, "font.size");
        CopyString(cf, "font.name", props, "font.name");
        CopyString(cf, "font.color", props, "font.color");
    }

    // ==================== Data validations ====================

    private static void EmitValidations(ExcelHandler xl, string sheetPath, int count,
        List<BatchItem> items, List<UnsupportedWarning> warnings)
    {
        for (int i = 1; i <= count; i++)
        {
            DocumentNode dv;
            try { dv = xl.Get($"{sheetPath}/validation[{i}]"); }
            catch (Exception ex)
            {
                warnings.Add(new UnsupportedWarning("validation", $"{sheetPath}/validation[{i}]", ex.Message));
                continue;
            }
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var key in new[]
            {
                "ref", "type", "operator", "formula1", "formula2",
                "errorTitle", "error", "promptTitle", "prompt",
                // errorStyle: warning/information popups silently became the
                // stop default on replay when this was omitted.
                "errorStyle",
            })
                CopyString(dv, key, props, key);
            foreach (var key in new[] { "allowBlank", "showError", "showInput", "inCellDropdown" })
            {
                // Add defaults each of these to true; carry explicit values
                // either way so the OOXML attribute round-trips exactly.
                if (dv.Format.TryGetValue(key, out var v) && v is bool b)
                    props[key] = b ? "true" : "false";
            }
            if (!props.ContainsKey("ref")) continue;
            items.Add(new BatchItem { Command = "add", Parent = sheetPath, Type = "validation", Props = props });
        }
    }

    // ==================== Comments ====================

    private static void EmitComments(ExcelHandler xl, string sheetPath, int count,
        List<BatchItem> items, List<UnsupportedWarning> warnings)
    {
        for (int i = 1; i <= count; i++)
        {
            DocumentNode c;
            try { c = xl.Get($"{sheetPath}/comment[{i}]"); }
            catch (Exception ex)
            {
                warnings.Add(new UnsupportedWarning("comment", $"{sheetPath}/comment[{i}]", ex.Message));
                continue;
            }
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            CopyString(c, "ref", props, "ref");
            CopyString(c, "author", props, "author");
            // Multi-run comments replay via runs=<json>; single-run comments
            // keep the text=/font.* path. Only one of the two carries content.
            if (c.Format.TryGetValue("runs", out var cRuns) && cRuns is string cRunsJson && cRunsJson.Length > 2)
            {
                props["runs"] = cRunsJson;
            }
            else
            {
                if (!string.IsNullOrEmpty(c.Text)) props["text"] = c.Text!;
                CopyBool(c, "font.bold", props, "font.bold");
                CopyBool(c, "font.italic", props, "font.italic");
                CopyString(c, "font.underline", props, "font.underline");
                CopyString(c, "font.color", props, "font.color");
                CopyString(c, "font.size", props, "font.size");
                CopyString(c, "font.name", props, "font.name");
            }
            if (!props.ContainsKey("ref")) continue;
            items.Add(new BatchItem { Command = "add", Parent = sheetPath, Type = "comment", Props = props });
        }
    }

    // ==================== Charts ====================

    private static void EmitCharts(ExcelHandler xl, string sheetPath, int count,
        List<BatchItem> items, List<UnsupportedWarning> warnings)
    {
        // Extended (cx:) charts ride the verbatim chartex carrier
        // (EmitChartExCharts); emitting them here too would double them.
        var extendedFlags = xl.GetDumpChartExtendedFlags(sheetPath.TrimStart('/'));
        for (int i = 1; i <= count; i++)
        {
            if (i - 1 < extendedFlags.Count && extendedFlags[i - 1]) continue;
            DocumentNode chart;
            try { chart = xl.Get($"{sheetPath}/chart[{i}]"); }
            catch (Exception ex)
            {
                warnings.Add(new UnsupportedWarning("chart", $"{sheetPath}/chart[{i}]", ex.Message));
                continue;
            }
            // Reuse the battle-tested docx chart transcription (ChartHelper is
            // the shared reader/builder across all three formats).
            var spec = new WordBatchEmitter.ChartSpec(
                chart.Format, chart.InternalFormat, chart.Children ?? new List<DocumentNode>());
            Dictionary<string, string> props;
            try { props = WordBatchEmitter.BuildChartProps(spec); }
            catch (Exception ex)
            {
                warnings.Add(new UnsupportedWarning("chart", chart.Path ?? sheetPath,
                    $"chart transcription failed: {ex.Message}"));
                continue;
            }
            props.Remove("anchor"); // re-added below in canonical form

            // Excel-specific: when every series carries a range reference
            // (Sheet1!$B$2:$B$5), replay via dotted seriesN.values/categories
            // refs instead of literal `data=` — the values already live in the
            // replayed sheet (import baseline), and refs keep the chart LIVE
            // (editing a cell updates the chart, matching the source).
            var series = (chart.Children ?? new List<DocumentNode>())
                .Where(s => s.Type == "series"
                    && !(s.Format.TryGetValue("refLine", out var rl) && rl?.ToString() == "true"))
                .ToList();
            if (series.Count > 0 && series.All(s =>
                    s.Format.TryGetValue("valuesRef", out var vr) && vr is string vrS && vrS.Length > 0))
            {
                props.Remove("data");
                props.Remove("categories");
                // BuildChartProps also emits combined literal `series{N}` keys
                // ("Name:1,2,3"). For externally-authored files whose series
                // carry no numCache the values half is empty, producing a
                // malformed dangling-colon spec ("'Sheet'!B1:") — drop those;
                // the dotted series{N}.name/.values refs below replay the
                // series and repopulate the cache. Well-formed combined keys
                // are kept: they carry the resolved cache literals.
                foreach (var combinedKey in props.Keys
                    .Where(k => System.Text.RegularExpressions.Regex.IsMatch(k, @"^series\d+$")
                        && (props[k].Length == 0 || props[k].EndsWith(":")))
                    .ToList())
                    props.Remove(combinedKey);
                for (int n = 0; n < series.Count; n++)
                {
                    var s = series[n];
                    var idx = n + 1;
                    if (s.Format.TryGetValue("name", out var nm) && nm is string nmS && nmS.Length > 0)
                        props[$"series{idx}.name"] = nmS;
                    props[$"series{idx}.values"] = (string)s.Format["valuesRef"]!;
                    if (s.Format.TryGetValue("categoriesRef", out var cr) && cr is string crS && crS.Length > 0)
                        props[$"series{idx}.categories"] = crS;
                    // Bubble series carry a third dimension in <c:bubbleSize>.
                    // Prefer the cell ref (keeps the size live-linked); fall back
                    // to the cached literal so the pixel geometry survives even
                    // when the source used a literal numLit. Without this the
                    // Builder's BuildBubbleChart defaults size = y-values.
                    if (s.Format.TryGetValue("bubbleSizeRef", out var bsr) && bsr is string bsrS && bsrS.Length > 0)
                        props[$"series{idx}.bubbleSizeRef"] = bsrS;
                    else if (s.Format.TryGetValue("bubbleSize", out var bs) && bs is string bsS && bsS.Length > 0)
                        props[$"series{idx}.bubbleSize"] = bsS;
                }
            }

            if (chart.Format.TryGetValue("anchor", out var anch) && anch is string anchS && anchS.Length > 0)
            {
                props["anchor"] = anchS;
                // anchor defines the full rectangle on replay; the x/y/width/
                // height BuildChartProps also emitted would be ignored with a
                // warning — drop them so the replay is warning-free.
                props.Remove("x");
                props.Remove("y");
                props.Remove("width");
                props.Remove("height");
            }

            // A chart with zero series and no cached literals transcribes to
            // props with neither `data` nor dotted seriesN.values refs — and
            // `add chart` hard-requires data, so emitting the row would make
            // the whole dump unreplayable (atomic batch: one bad item fails
            // the file). Skip the empty frame with an explicit warning
            // instead; every other carrier degrades the same way.
            if (!props.ContainsKey("data")
                && !props.Keys.Any(k => System.Text.RegularExpressions.Regex.IsMatch(k, @"^series\d+\.values$")))
            {
                warnings.Add(new UnsupportedWarning("chart", chart.Path ?? sheetPath,
                    "chart has no series data (empty plot area) — frame dropped on dump; `add chart` requires data"));
                continue;
            }

            items.Add(new BatchItem { Command = "add", Parent = sheetPath, Type = "chart", Props = props });
        }
    }

    // ==================== Sparklines ====================

    private static void EmitSparklines(ExcelHandler xl, string sheetPath, int count,
        List<BatchItem> items, List<UnsupportedWarning> warnings)
    {
        for (int i = 1; i <= count; i++)
        {
            DocumentNode spk;
            try { spk = xl.Get($"{sheetPath}/sparkline[{i}]"); }
            catch (Exception ex)
            {
                warnings.Add(new UnsupportedWarning("sparkline", $"{sheetPath}/sparkline[{i}]", ex.Message));
                continue;
            }
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            CopyString(spk, "location", props, "location");
            CopyString(spk, "dataRange", props, "dataRange");
            CopyString(spk, "type", props, "type");
            CopyString(spk, "color", props, "color");
            CopyString(spk, "negativeColor", props, "negativeColor");
            CopyBool(spk, "markers", props, "markers");
            CopyBool(spk, "highPoint", props, "highPoint");
            CopyBool(spk, "lowPoint", props, "lowPoint");
            CopyBool(spk, "firstPoint", props, "firstPoint");
            CopyBool(spk, "lastPoint", props, "lastPoint");
            CopyBool(spk, "negative", props, "negative");
            CopyString(spk, "highMarkerColor", props, "highMarkerColor");
            CopyString(spk, "lowMarkerColor", props, "lowMarkerColor");
            CopyString(spk, "firstMarkerColor", props, "firstMarkerColor");
            CopyString(spk, "lastMarkerColor", props, "lastMarkerColor");
            CopyString(spk, "markersColor", props, "markersColor");
            CopyValue(spk, "lineWeight", props, "lineWeight");
            CopyString(spk, "displayEmptyCellsAs", props, "displayEmptyCellsAs");
            CopyBool(spk, "displayXAxis", props, "displayXAxis");
            CopyBool(spk, "rightToLeft", props, "rightToLeft");
            CopyBool(spk, "dateAxis", props, "dateAxis");
            if (!props.ContainsKey("location") || !props.ContainsKey("dataRange"))
            {
                warnings.Add(new UnsupportedWarning("sparkline", spk.Path ?? sheetPath,
                    "sparkline missing location/dataRange; skipped"));
                continue;
            }
            items.Add(new BatchItem { Command = "add", Parent = sheetPath, Type = "sparkline", Props = props });
        }
    }

    // ==================== Pictures ====================

    private static void EmitPictures(ExcelHandler xl, string sheetName, string sheetPath, int count,
        List<BatchItem> items, List<UnsupportedWarning> warnings)
    {
        for (int i = 1; i <= count; i++)
        {
            DocumentNode pic;
            try { pic = xl.Get($"{sheetPath}/picture[{i}]"); }
            catch (Exception ex)
            {
                warnings.Add(new UnsupportedWarning("picture", $"{sheetPath}/picture[{i}]", ex.Message));
                continue;
            }
            var dataUri = xl.GetDumpPictureDataUri(sheetName, i);
            if (dataUri == null)
            {
                warnings.Add(new UnsupportedWarning("picture", pic.Path ?? sheetPath,
                    "picture image part could not be read; skipped"));
                continue;
            }
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["src"] = dataUri,
            };
            // Exact anchor rectangle (x/y cell origin + EMU dimensions), not
            // the Get surface's whole-cell width/height deltas — those drop
            // sub-cell To offsets, so replay would snap the picture to cell
            // boundaries (visible drift in real Excel rendering).
            var picAnchor = xl.GetDumpPictureAnchorEmu(sheetName, i);
            if (picAnchor != null)
            {
                if (picAnchor.Mode == "absolute")
                {
                    // Absolute anchors position by EMU <xdr:pos>, and the add
                    // vocabulary takes x/y as EMU for anchorMode=absolute.
                    props["anchorMode"] = "absolute";
                    props["x"] = $"{picAnchor.XEmu}emu";
                    props["y"] = $"{picAnchor.YEmu}emu";
                }
                else
                {
                    if (picAnchor.Mode == "oneCell") props["anchorMode"] = "oneCell";
                    props["x"] = picAnchor.X.ToString();
                    props["y"] = picAnchor.Y.ToString();
                }
                props["width"] = $"{picAnchor.WidthEmu}emu";
                props["height"] = $"{picAnchor.HeightEmu}emu";
                if (picAnchor.HasFromOffset)
                    warnings.Add(new UnsupportedWarning("picture", pic.Path ?? sheetPath,
                        "picture From marker has a sub-cell offset; replay anchors at the cell origin (add vocabulary cannot express from-offsets)"));
            }
            else
            {
                CopyValue(pic, "x", props, "x");
                CopyValue(pic, "y", props, "y");
                CopyValue(pic, "width", props, "width");
                CopyValue(pic, "height", props, "height");
            }
            CopyString(pic, "alt", props, "alt");
            CopyString(pic, "name", props, "name");
            CopyString(pic, "title", props, "title");
            CopyValue(pic, "rotation", props, "rotation");
            CopyString(pic, "flip", props, "flip");
            CopyString(pic, "crop", props, "crop");
            CopyString(pic, "hyperlink", props, "hyperlink");
            // Add-only accessibility/visual props now surfaced by Get; without
            // these the dump silently dropped alt-text-exclusion (decorative)
            // and transparency (opacity), which real Excel does not regenerate.
            CopyValue(pic, "opacity", props, "opacity");
            CopyValue(pic, "decorative", props, "decorative");
            items.Add(new BatchItem { Command = "add", Parent = sheetPath, Type = "picture", Props = props });
        }
    }

    // ==================== Shapes ====================

    private static void EmitShapes(ExcelHandler xl, string sheetName, string sheetPath,
        List<BatchItem> items, List<UnsupportedWarning> warnings)
    {
        // Skip mc:Fallback placeholder shapes (slicer down-level rectangles):
        // the owning feature regenerates them on replay; re-adding them as
        // real shapes duplicates content and breaks dump idempotency.
        var fallbackFlags = xl.GetDumpShapeFallbackFlags(sheetPath.TrimStart('/'));
        var replayItems = xl.GetDumpShapeReplayItems(sheetName);
        foreach (var replay in replayItems)
        {
            if (replay.GroupAnchorXml != null)
            {
                var groupProps = new Dictionary<string, string>(StringComparer.Ordinal)
                {
                    ["anchor-xml"] = replay.GroupAnchorXml,
                    ["hyperlinks"] = ExcelHandler.EncodeDumpDrawingHyperlinks(replay.GroupHyperlinks),
                };
                items.Add(new BatchItem
                {
                    Command = "add-part",
                    Parent = sheetPath,
                    Type = "drawing-group",
                    Props = groupProps,
                });
                continue;
            }

            if (!replay.ShapeIndex.HasValue) continue;
            var i = replay.ShapeIndex.Value;
            if (!string.IsNullOrEmpty(replay.Warning))
                warnings.Add(new UnsupportedWarning("group", $"{sheetPath}/shape[{i}]", replay.Warning));
            if (i - 1 < fallbackFlags.Count && fallbackFlags[i - 1]) continue;
            DocumentNode shp;
            try { shp = xl.Get($"{sheetPath}/shape[{i}]"); }
            catch (Exception ex)
            {
                warnings.Add(new UnsupportedWarning("shape", $"{sheetPath}/shape[{i}]", ex.Message));
                continue;
            }
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrEmpty(shp.Text)) props["text"] = shp.Text!;
            CopyString(shp, "name", props, "name");
            CopyString(shp, "geometry", props, "preset");
            CopyString(shp, "fill", props, "fill");
            CopyString(shp, "gradientFill", props, "gradientFill");
            CopyValue(shp, "x", props, "x");
            CopyValue(shp, "y", props, "y");
            CopyValue(shp, "width", props, "width");
            CopyValue(shp, "height", props, "height");
            CopyString(shp, "size", props, "size");
            CopyBool(shp, "bold", props, "bold");
            CopyBool(shp, "italic", props, "italic");
            CopyString(shp, "underline", props, "font.underline");
            CopyString(shp, "color", props, "color");
            CopyString(shp, "font", props, "font");
            CopyString(shp, "align", props, "align");
            CopyString(shp, "valign", props, "valign");
            CopyValue(shp, "rotation", props, "rotation");
            CopyString(shp, "flip", props, "flip");
            CopyString(shp, "line", props, "line");
            CopyString(shp, "hyperlink", props, "hyperlink");
            // Effect + text-inset props. Get surfaces them (shadow=#000000,
            // glow=#4472C4-8, softEdge=5pt, margin=7.2pt) and Add re-consumes
            // those exact forms, but the copy-list omitted them so dump dropped
            // the shape's whole <a:effectLst> and text inset on round-trip.
            CopyString(shp, "shadow", props, "shadow");
            CopyString(shp, "glow", props, "glow");
            CopyString(shp, "softEdge", props, "softEdge");
            CopyString(shp, "reflection", props, "reflection");
            CopyString(shp, "margin", props, "margin");
            items.Add(new BatchItem { Command = "add", Parent = sheetPath, Type = "shape", Props = props });
        }
    }

    // ==================== Pivot tables ====================

    // Called from EmitExcel AFTER all sheets (a pivot's source range can live
    // on a later sheet); the value baseline already excluded each pivot's
    // location rectangle so the rebuilt pivot lands on clean cells.
    internal static void EmitPivotTables(ExcelHandler xl, string sheetPath, int count,
        List<BatchItem> items, List<UnsupportedWarning> warnings)
    {
        for (int i = 1; i <= count; i++)
        {
            DocumentNode pt;
            try { pt = xl.Get($"{sheetPath}/pivottable[{i}]"); }
            catch (Exception ex)
            {
                warnings.Add(new UnsupportedWarning("pivottable", $"{sheetPath}/pivottable[{i}]", ex.Message));
                continue;
            }
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            CopyString(pt, "source", props, "source");
            if (!props.ContainsKey("source"))
            {
                warnings.Add(new UnsupportedWarning("pivottable", pt.Path ?? sheetPath,
                    "pivot cache source range could not be read; skipped"));
                continue;
            }
            // Readback surfaces the full location rectangle; Add wants the
            // top-left position cell.
            if (pt.Format.TryGetValue("location", out var loc) && loc is string locS && locS.Length > 0)
                props["position"] = locS.Split(':')[0];
            CopyString(pt, "name", props, "name");
            CopyString(pt, "style", props, "style");
            CopyString(pt, "rows", props, "rows");
            CopyString(pt, "cols", props, "cols");
            CopyString(pt, "filters", props, "filters");
            CopyString(pt, "layout", props, "layout");
            CopyString(pt, "grandTotalCaption", props, "grandtotalcaption");
            CopyString(pt, "subtotals", props, "subtotals");
            if (pt.Format.TryGetValue("blankRows", out var br) && br?.ToString() == "true")
                props["blankrows"] = "true";
            if (pt.Format.TryGetValue("repeatLabels", out var rl2) && rl2?.ToString() == "true")
                props["repeatlabels"] = "true";
            // Grand totals default to ON — carry only the off direction.
            if (pt.Format.TryGetValue("rowGrandTotals", out var rgt) && rgt?.ToString() == "false")
                props["rowgrandtotals"] = "false";
            if (pt.Format.TryGetValue("colGrandTotals", out var cgt) && cgt?.ToString() == "false")
                props["colgrandtotals"] = "false";
            // Style-info toggles: emit only non-default values (headers on,
            // stripes off, last column on).
            EmitPivotToggle(pt, props, "showRowHeaders", "showrowheaders", defaultValue: true);
            EmitPivotToggle(pt, props, "showColHeaders", "showcolheaders", defaultValue: true);
            EmitPivotToggle(pt, props, "showRowStripes", "showrowstripes", defaultValue: false);
            EmitPivotToggle(pt, props, "showColStripes", "showcolstripes", defaultValue: false);
            EmitPivotToggle(pt, props, "showLastColumn", "showlastcolumn", defaultValue: true);

            // dataFieldN = "DisplayName:func:Field" (+ optional
            // dataFieldN.showAs) → values="Field:func[:showAs]:name=DisplayName".
            var valueSpecs = new List<string>();
            var dfCount = pt.Format.TryGetValue("dataFieldCount", out var dfc)
                && int.TryParse(dfc?.ToString(), out var dfcN) ? dfcN : 0;
            var dfOk = true;
            for (int d = 1; d <= dfCount; d++)
            {
                if (!pt.Format.TryGetValue($"dataField{d}", out var dfv) || dfv is not string dfS)
                { dfOk = false; break; }
                var seg = dfS.Split(':');
                if (seg.Length < 3) { dfOk = false; break; }
                // Right-anchored: last two segments are func and FIELD-index;
                // the display name may itself contain ':'.
                var func = seg[^2];
                var displayName = string.Join(":", seg[..^2]);
                // Prefer the resolved source field NAME (dataField{N}.srcField)
                // over the raw index in seg[^1]: a bare index breaks replay when
                // the rebuilt cache has a different field layout.
                var field = (pt.Format.TryGetValue($"dataField{d}.srcField", out var sf)
                    && sf is string sfS && sfS.Length > 0) ? sfS : seg[^1];
                if (field.Contains(',') || displayName.Contains(','))
                {
                    warnings.Add(new UnsupportedWarning("pivottable", pt.Path ?? sheetPath,
                        $"data field '{displayName}' contains a comma; value spec cannot express it"));
                    dfOk = false; break;
                }
                var spec = $"{field}:{func}";
                if (pt.Format.TryGetValue($"dataField{d}.showAs", out var sa)
                    && sa is string saS && saS.Length > 0 && saS != "normal")
                    spec += $":{saS}";
                spec += $":name={displayName}";
                valueSpecs.Add(spec);
            }
            if (dfOk && valueSpecs.Count > 0)
                props["values"] = string.Join(",", valueSpecs);

            // Readback-only state with no Add vocabulary — surface the loss.
            foreach (var lossy in new[] { "sortByField", "collapsedFields", "axisAsDataField" })
                if (pt.Format.ContainsKey(lossy))
                    warnings.Add(new UnsupportedWarning("pivottable", pt.Path ?? sheetPath,
                        $"pivot {lossy} state is not round-tripped (no add vocabulary)"));

            items.Add(new BatchItem { Command = "add", Parent = sheetPath, Type = "pivottable", Props = props });
        }
    }

    // ==================== chartEx (extended charts) ====================

    // Waterfall / funnel / sunburst / treemap etc. have no semantic add
    // vocabulary; carry the ExtendedChartPart (+ colors/style sub-parts)
    // verbatim with pinned rIds, then raw-append the hosting TwoCellAnchor
    // into the sheet's drawing. Mirrors the pptx SmartArt passthrough.
    private static void EmitChartExCharts(ExcelHandler xl, string sheetName, string sheetPath,
        List<BatchItem> items, List<UnsupportedWarning> warnings)
    {
        List<ExcelHandler.DumpChartExSlice> slices;
        try { slices = xl.GetDumpChartExSlices(sheetName); }
        catch (Exception ex)
        {
            warnings.Add(new UnsupportedWarning("chartex", sheetPath, $"chartEx scan failed: {ex.Message}"));
            return;
        }
        foreach (var s in slices)
        {
            var props = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["rid"] = s.RelId,
                ["xml"] = s.XmlBase64,
            };
            if (s.ColorsRelId != null && s.ColorsXmlBase64 != null)
            {
                props["colors-rid"] = s.ColorsRelId;
                props["colors-xml"] = s.ColorsXmlBase64;
            }
            if (s.StyleRelId != null && s.StyleXmlBase64 != null)
            {
                props["style-rid"] = s.StyleRelId;
                props["style-xml"] = s.StyleXmlBase64;
            }
            items.Add(new BatchItem { Command = "add-part", Parent = sheetPath, Type = "chartex", Props = props });
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = $"{sheetPath}/drawing",
                Xpath = "/xdr:wsDr",
                Action = "append",
                Xml = s.AnchorXml,
            });
        }
    }

    // ==================== OLE embedded objects ====================

    // Verbatim carrier, one all-in-one add-part per object (see the Excel
    // AddPart "ole" case for the wiring it performs). Mirrors the pptx
    // EmitOleForSlide contract.
    private static void EmitOles(ExcelHandler xl, string sheetName, string sheetPath,
        List<BatchItem> items, List<UnsupportedWarning> warnings)
    {
        List<ExcelHandler.DumpOleSlice> slices;
        try { slices = xl.GetDumpOleSlices(sheetName); }
        catch (Exception ex)
        {
            warnings.Add(new UnsupportedWarning("ole", sheetPath, $"OLE scan failed: {ex.Message}"));
            return;
        }
        foreach (var s in slices)
        {
            var props = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["rid"] = s.RelId,
                ["data"] = s.DataBase64,
                ["ole-kind"] = s.Kind,
                ["content-type"] = s.ContentType,
                ["extension"] = s.Extension,
                ["object-xml"] = s.ObjectXml,
            };
            if (s.IconRelId != null && s.IconBase64 != null)
            {
                props["icon-rid"] = s.IconRelId;
                props["icon-data"] = s.IconBase64;
                if (s.IconContentType != null) props["icon-content-type"] = s.IconContentType;
            }
            if (s.VmlShapeXml != null) props["vml-shape"] = s.VmlShapeXml;
            else
                warnings.Add(new UnsupportedWarning("ole", sheetPath,
                    "OLE VML anchor shape not found; the replayed object may lose its on-sheet placement"));
            items.Add(new BatchItem { Command = "add-part", Parent = sheetPath, Type = "ole", Props = props });
        }
    }

    // ==================== Slicers ====================

    // Runs after EmitPivotTables (a slicer binds to a pivot by name). The
    // slicer's drawing anchor has no readback surface, so replay lands the
    // slicer at AddSlicer's default position — surfaced as a warning.
    internal static void EmitSlicers(ExcelHandler xl, string sheetPath, int count,
        List<BatchItem> items, List<UnsupportedWarning> warnings)
    {
        for (int i = 1; i <= count; i++)
        {
            DocumentNode sl;
            try { sl = xl.Get($"{sheetPath}/slicer[{i}]"); }
            catch (Exception ex)
            {
                warnings.Add(new UnsupportedWarning("slicer", $"{sheetPath}/slicer[{i}]", ex.Message));
                continue;
            }
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            CopyString(sl, "pivotTable", props, "pivotTable");
            CopyString(sl, "field", props, "field");
            if (!props.ContainsKey("pivotTable") || !props.ContainsKey("field"))
            {
                warnings.Add(new UnsupportedWarning("slicer", sl.Path ?? sheetPath,
                    "slicer pivot binding could not be read; skipped"));
                continue;
            }
            CopyString(sl, "name", props, "name");
            CopyString(sl, "caption", props, "caption");
            CopyString(sl, "style", props, "style");
            CopyValue(sl, "columnCount", props, "columnCount");
            CopyValue(sl, "rowHeight", props, "rowHeight");
            warnings.Add(new UnsupportedWarning("slicer", sl.Path ?? sheetPath,
                "slicer position and item selection state are not round-tripped (no readback surface); replay uses defaults"));
            items.Add(new BatchItem { Command = "add", Parent = sheetPath, Type = "slicer", Props = props });
        }
    }

    private static void EmitPivotToggle(DocumentNode pt, Dictionary<string, string> props,
        string getKey, string setKey, bool defaultValue)
    {
        if (pt.Format.TryGetValue(getKey, out var v) && v is string s
            && bool.TryParse(s, out var b) && b != defaultValue)
            props[setKey] = b ? "true" : "false";
    }

    // ==================== Copy helpers ====================

    private static void CopyString(DocumentNode node, string getKey,
        Dictionary<string, string> props, string setKey)
    {
        if (node.Format.TryGetValue(getKey, out var v) && v is string s && s.Length > 0)
            props[setKey] = s;
    }

    private static void CopyBool(DocumentNode node, string getKey,
        Dictionary<string, string> props, string setKey)
    {
        if (node.Format.TryGetValue(getKey, out var v) && v is bool b && b)
            props[setKey] = "true";
    }

    private static void CopyValue(DocumentNode node, string getKey,
        Dictionary<string, string> props, string setKey)
    {
        if (node.Format.TryGetValue(getKey, out var v) && v != null)
        {
            var s = FormatValue(v);
            if (s.Length > 0) props[setKey] = s;
        }
    }
}
