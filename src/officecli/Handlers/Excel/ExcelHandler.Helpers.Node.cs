// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{

    internal static DocumentNode SparklineGroupToNode(string sheetName, X14.SparklineGroup spkGroup, int index)
    {
        var node = new DocumentNode
        {
            Path = $"/{sheetName}/sparkline[{index}]",
            Type = "sparkline"
        };

        // Type: default is line when attribute is absent. The OOXML enum
        // calls win-loss sparklines "Stacked", which collides with bar-chart
        // stacked grouping in user vocabulary; surface it as "winLoss" on
        // readback to match Excel's UI label. Set still accepts both
        // "stacked" and "winLoss" / "winloss" / "win-loss" via the input
        // alias map (ExcelHandler.Add.Drawings.cs:881).
        string spkType;
        if (spkGroup.Type?.HasValue == true)
        {
            var tv = spkGroup.Type.Value;
            spkType = tv == X14.SparklineTypeValues.Column ? "column"
                : tv == X14.SparklineTypeValues.Stacked ? "winLoss"
                : "line";
        }
        else
        {
            spkType = "line";
        }
        node.Format["type"] = spkType;

        // Color
        var colorRgb = spkGroup.SeriesColor?.Rgb?.Value;
        node.Format["color"] = colorRgb != null
            ? ParseHelpers.FormatHexColor(colorRgb)
            : "#4472C4";

        // Negative color
        var negColorRgb = spkGroup.NegativeColor?.Rgb?.Value;
        if (negColorRgb != null)
            node.Format["negativeColor"] = ParseHelpers.FormatHexColor(negColorRgb);

        // Boolean flags
        if (spkGroup.Markers?.Value == true) node.Format["markers"] = true;
        if (spkGroup.High?.Value == true) node.Format["highPoint"] = true;
        if (spkGroup.Low?.Value == true) node.Format["lowPoint"] = true;
        if (spkGroup.First?.Value == true) node.Format["firstPoint"] = true;
        if (spkGroup.Last?.Value == true) node.Format["lastPoint"] = true;
        if (spkGroup.Negative?.Value == true) node.Format["negative"] = true;

        // Marker colors — Add persists these into the x14 group; without the
        // readback they were invisible to Get and dropped by dump round-trips.
        if (spkGroup.HighMarkerColor?.Rgb?.Value is string spkHiMc)
            node.Format["highMarkerColor"] = ParseHelpers.FormatHexColor(spkHiMc);
        if (spkGroup.LowMarkerColor?.Rgb?.Value is string spkLoMc)
            node.Format["lowMarkerColor"] = ParseHelpers.FormatHexColor(spkLoMc);
        if (spkGroup.FirstMarkerColor?.Rgb?.Value is string spkFiMc)
            node.Format["firstMarkerColor"] = ParseHelpers.FormatHexColor(spkFiMc);
        if (spkGroup.LastMarkerColor?.Rgb?.Value is string spkLaMc)
            node.Format["lastMarkerColor"] = ParseHelpers.FormatHexColor(spkLaMc);
        if (spkGroup.MarkersColor?.Rgb?.Value is string spkMkMc)
            node.Format["markersColor"] = ParseHelpers.FormatHexColor(spkMkMc);

        // Line weight
        if (spkGroup.LineWeight?.HasValue == true)
            node.Format["lineWeight"] = spkGroup.LineWeight.Value;

        // Group-level axis / empty-cell / RTL attributes. Add/Set persist these
        // into the x14 group; without the readback they were invisible to Get
        // and silently dropped by dump round-trips.
        if (spkGroup.DisplayEmptyCellsAs?.HasValue == true)
        {
            var de = spkGroup.DisplayEmptyCellsAs.Value;
            node.Format["displayEmptyCellsAs"] =
                de == X14.DisplayBlanksAsValues.Gap ? "gap"
                : de == X14.DisplayBlanksAsValues.Span ? "span"
                : "zero";
        }
        if (spkGroup.DisplayXAxis?.Value == true) node.Format["displayXAxis"] = true;
        if (spkGroup.RightToLeft?.Value == true) node.Format["rightToLeft"] = true;
        if (spkGroup.DateAxis?.Value == true) node.Format["dateAxis"] = true;

        // Cell / range from first sparkline element
        var firstSparkline = spkGroup.GetFirstChild<X14.Sparklines>()?.GetFirstChild<X14.Sparkline>();
        if (firstSparkline != null)
        {
            // CONSISTENCY(canonical-key): schema canonical keys are 'location'
            // (target cell) and 'dataRange' (source range). 'cell'/'range' are
            // legacy aliases retained on input.
            var cell = firstSparkline.ReferenceSequence?.Text ?? "";
            node.Format["location"] = cell;

            // Strip the sheet prefix only when it names this sparkline's OWN
            // sheet (Add auto-qualifies a bare 'A1:E1' to 'Sheet1!A1:E1'), so
            // same-sheet ranges read back in the bare form the user supplied.
            // A prefix naming a DIFFERENT sheet is a genuine cross-sheet
            // qualifier and must be preserved (R3-8) so the range stays
            // unambiguous and round-trips.
            var formulaText = firstSparkline.Formula?.Text ?? "";
            var excl = formulaText.IndexOf('!');
            if (excl >= 0)
            {
                var prefix = formulaText[..excl].Trim('\'');
                node.Format["dataRange"] = prefix.Equals(sheetName, StringComparison.OrdinalIgnoreCase)
                    ? formulaText[(excl + 1)..]
                    : formulaText;
            }
            else
            {
                node.Format["dataRange"] = formulaText;
            }
        }

        return node;
    }

    private List<DocumentNode> GetSheetChildNodes(string sheetName, SheetData sheetData, int depth, WorksheetPart? worksheetPart = null)
    {
        var children = new List<DocumentNode>();
        var eval = depth > 0 && worksheetPart != null ? new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart) : null;
        // R6-5: dedupe by RowIndex. When a sheet contains both source data
        // rows and pivot-rendered rows (possible when a pivot is placed on
        // its own source sheet), the renderer appends additional <row> nodes
        // that can collide with existing RowIndex values. Children should
        // expose each logical row once.
        var seenRowIndices = new HashSet<uint>();
        foreach (var row in sheetData.Elements<Row>())
        {
            var ridx = row.RowIndex?.Value ?? 0;
            if (ridx != 0 && !seenRowIndices.Add(ridx))
                continue;

            // Bulk enumeration omits value-less, formula-less empty cells —
            // and rows that thereby become empty and carry no row-level
            // formatting — so `get /` and `get /Sheet` stay bounded on sheets
            // that declare millions of empty cells (issue #149). The load-time
            // WorksheetBloatFilter already strips BARE empties (no style); a
            // STYLED empty cell (<c s="1"/>) must stay in the DOM (common
            // spreadsheet apps keep it — its xf may hold real formatting), so the
            // only consistent place to suppress it is here, at output. This
            // makes bare vs styled empties behave identically in bulk listings.
            // Targeted reads (`get /Sheet/A1`, Query.cs empty-cell path) still
            // return the cell, so no formatting becomes unreachable.
            var contentCells = row.Elements<Cell>().Where(CellHasContent).ToList();
            if (contentCells.Count == 0 && !RowHasMeaningfulAttrs(row))
                continue;

            var rowIdx = row.RowIndex?.Value ?? 0;
            var rowNode = new DocumentNode
            {
                Path = $"/{sheetName}/row[{rowIdx}]",
                Type = "row",
                ChildCount = contentCells.Count
            };
            // CONSISTENCY(unit-qualified-readback): pt-suffix row height
            // (Query.cs:433/1367 mirror). Stored value is already points.
            if (row.Height?.Value != null)
                rowNode.Format["height"] = $"{row.Height.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}pt";
            if (row.Hidden?.Value == true)
                rowNode.Format["hidden"] = true;
            // CONSISTENCY(nested-get-parity): the same node fetched directly
            // (Query.cs row path) exposes outlineLevel/collapsed — the nested
            // sheet-children view must return the identical Format dict.
            if (row.OutlineLevel?.HasValue == true && row.OutlineLevel.Value > 0)
                rowNode.Format["outlineLevel"] = (int)row.OutlineLevel.Value;
            if (row.Collapsed?.Value == true)
                rowNode.Format["collapsed"] = true;

            if (depth > 0)
            {
                foreach (var cell in contentCells)
                {
                    rowNode.Children.Add(CellToNode(sheetName, cell, worksheetPart, eval));
                }
            }

            children.Add(rowNode);
        }

        // Add chart children from DrawingsPart
        if (worksheetPart?.DrawingsPart != null)
        {
            var chartParts = worksheetPart.DrawingsPart.ChartParts.ToList();
            for (int i = 0; i < chartParts.Count; i++)
            {
                var chart = chartParts[i].ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
                var chartNode = new DocumentNode
                {
                    Path = $"/{sheetName}/chart[{i + 1}]",
                    Type = "chart"
                };
                if (chart != null)
                    ChartHelper.ReadChartProperties(chart, chartNode, 0);
                children.Add(chartNode);
            }
        }

        // R16-1: expose pivottable children so Get /Sheet1 lists them.
        // CONSISTENCY(sheet-children): same pattern as chart children above.
        if (worksheetPart != null)
        {
            var pivotParts = worksheetPart.PivotTableParts.ToList();
            for (int i = 0; i < pivotParts.Count; i++)
            {
                var ptNode = new DocumentNode
                {
                    Path = $"/{sheetName}/pivottable[{i + 1}]",
                    Type = "pivottable"
                };
                var pivotDef = pivotParts[i].PivotTableDefinition;
                if (pivotDef != null)
                    Core.PivotTableHelper.ReadPivotTableProperties(pivotDef, ptNode, pivotParts[i]);
                children.Add(ptNode);
            }

            // Slicers are addressable (/Sheet1/slicer[N] get/set/remove all
            // work) but were invisible in the parent sheet's child listing —
            // the only sheet-level element missing from enumeration.
            // CONSISTENCY(sheet-children): same pattern as pivottable above.
            var slicersPart = worksheetPart.GetPartsOfType<SlicersPart>().FirstOrDefault();
            var slicerElems = slicersPart?.Slicers?
                .Elements<DocumentFormat.OpenXml.Office2010.Excel.Slicer>().ToList();
            if (slicerElems != null)
            {
                for (int i = 0; i < slicerElems.Count; i++)
                {
                    var slNode = new DocumentNode
                    {
                        Path = $"/{sheetName}/slicer[{i + 1}]",
                        Type = "slicer"
                    };
                    if (TryFindSlicerByIndex(worksheetPart, i + 1, out var slElem, out var slCache) && slElem != null)
                        ReadSlicerProperties(slElem, slCache, slNode);
                    children.Add(slNode);
                }
            }
        }

        return children;
    }

    /// <summary>
    /// A cell carries content when it has a formula, a non-empty value, or an
    /// inline string. Style alone (<c>&lt;c s="1"/&gt;</c>) is NOT content —
    /// such cells are omitted from bulk enumeration (see GetSheetChildNodes)
    /// but stay in the DOM and remain reachable via a targeted get.
    /// </summary>
    private static bool CellHasContent(Cell cell)
    {
        if (cell.CellFormula != null) return true;
        if (!string.IsNullOrEmpty(cell.CellValue?.Text)) return true;
        if (cell.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.InlineString>() != null) return true;
        // Explicit empty-string cell (<c t="str"><v/></c>): a stored value
        // node with a string DataType is real content — COUNTA counts it,
        // ISBLANK is FALSE — even when the text is empty. Excluding it hid
        // the cell from bulk enumeration AND from dump, which silently
        // dropped it on replay. A typeless <c s="1"/> (style only, no value
        // node) remains non-content.
        if (cell.DataType?.HasValue == true && cell.CellValue != null) return true;
        return false;
    }

    /// <summary>
    /// Mirrors WorksheetBloatFilter.ProcessRow's row-keep rule: a row is
    /// meaningful (kept even with no content cells) when it carries any
    /// attribute beyond <c>r</c>/<c>spans</c>, or any namespaced attribute
    /// (height, hidden, style, outline, x14ac:dyDescent, …).
    /// </summary>
    private static bool RowHasMeaningfulAttrs(Row row)
    {
        foreach (var attr in row.GetAttributes())
        {
            if (!string.IsNullOrEmpty(attr.NamespaceUri)) return true;
            if (attr.LocalName is "r" or "spans") continue;
            return true;
        }
        return false;
    }

    private DocumentNode CellToNode(string sheetName, Cell cell, WorksheetPart? part = null, Core.FormulaEvaluator? evaluator = null)
    {
        var cellRef = cell.CellReference?.Value ?? "?";
        var formula = cell.CellFormula?.Text is { } fText
            ? Core.ModernFunctionQualifier.Unqualify(fText)
            : null;
        string type;
        if (cell.DataType?.HasValue != true)
        {
            // R12-F2: a formula whose cached value is a non-numeric string
            // should report type=String, not the Number default. Excel itself
            // writes t="str" on such cells; external tools or our own writer
            // occasionally leave the attribute off, so infer from the cached
            // value content.
            var raw = cell.CellValue?.Text;
            if (formula != null
                && !string.IsNullOrEmpty(raw)
                && !double.TryParse(raw, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out _))
                type = "String";
            else
                type = "Number";
        }
        else if (cell.DataType.Value == CellValues.String)
            type = "String";
        else if (cell.DataType.Value == CellValues.SharedString)
            type = "SharedString";
        else if (cell.DataType.Value == CellValues.Boolean)
            type = "Boolean";
        else if (cell.DataType.Value == CellValues.Error)
            type = "Error";
        else if (cell.DataType.Value == CellValues.InlineString)
            type = "InlineString";
        else if (cell.DataType.Value == CellValues.Date)
            type = "Date";
        else
            type = "Number";

        // Lazy-create evaluator whenever a formula is present. Used both to
        // back-fill displayText when no cachedValue exists and to surface
        // Format["computedValue"] alongside cachedValue so callers can diff
        // stale imports — see the formula branch below for the diff contract.
        if (evaluator == null && formula != null && part != null)
        {
            var sheetData = GetSheet(part).GetFirstChild<SheetData>();
            if (sheetData != null)
                evaluator = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);
        }

        var displayText = GetCellDisplayValue(cell, evaluator);

        // In-cell image ("Place in Cell" richValue): stored as t="e" #VALUE!
        // with a vm pointer into xl/richData. Surface it as an Image cell
        // instead of a bare error — the #VALUE! is only the down-level
        // fallback, not the cell's meaning.
        InCellImageInfo? inCellImage = null;
        if (type == "Error" && cell.ValueMetaIndex != null && TryGetInCellImage(cell, out var imgInfo))
        {
            inCellImage = imgInfo;
            type = "Image";
            displayText = "[image]";
        }

        var node = new DocumentNode
        {
            Path = $"/{sheetName}/{cellRef}",
            Type = "cell",
            Text = displayText,
            Preview = cellRef
        };

        node.Format["type"] = type;
        if (inCellImage is { } ici)
        {
            node.Format["image.contentType"] = ici.ContentType;
            node.Format["image.fileSize"] = ici.FileSize;
            // CONSISTENCY(picture-alt): bare `alt` is the project-wide
            // canonical readback key for image alt text (picture nodes in all
            // three handlers emit Format["alt"]).
            if (!string.IsNullOrEmpty(ici.Alt))
                node.Format["alt"] = ici.Alt!;
        }
        if (formula != null)
        {
            node.Format["formula"] = formula;
            // Three Format keys for formula cells:
            //   cachedValue   : raw <x:v> the producer persisted. Trusted as
            //                   XML state, not verified against re-evaluation.
            //                   Absent when <x:v> is missing (e.g. Add path
            //                   before the file has been opened in Excel).
            //   computedValue : what officecli's evaluator produces for the
            //                   same formula right now, in the current
            //                   workbook state. Emitted whenever the
            //                   evaluator succeeds, regardless of cachedValue
            //                   presence. Suppressed for formulas whose ref
            //                   points at a sheet that no longer exists —
            //                   ResolveSheetCellResult silently returns 0
            //                   in that branch, and a fake would pollute the
            //                   diff. view issues formula_not_evaluated
            //                   handles that case separately.
            //   evaluated     : cross-handler protocol verdict — true iff
            //                   the cell has a trustworthy value available
            //                   (cachedValue present OR computedValue
            //                   produced). False iff neither is available.
            //                   Agents can use this single bool to decide
            //                   whether to trust the displayed value;
            //                   diff cachedValue vs computedValue to detect
            //                   stale caches (see view issues
            //                   formula_cache_stale).
            var rawCached = cell.CellValue?.Text;
            string? computedValue = null;
            if (!string.IsNullOrEmpty(rawCached))
                node.Format["cachedValue"] = rawCached;
            if (evaluator != null && !FormulaReferencesMissingSheet(formula))
            {
                var report = evaluator.EvaluateForReport(formula, cellRef);
                if (report.Status == Core.EvalReportStatus.Evaluated)
                    computedValue = report.Result!.ToCellValueText();
                else if (report.Status == Core.EvalReportStatus.Error)
                    computedValue = report.Result!.ErrorValue!;
                if (computedValue != null)
                    node.Format["computedValue"] = computedValue;
            }
            // R10-1: a freshly-added formula cell has no DataType and no cached
            // <x:v>, so the type heuristic above defaulted to Number. When the
            // evaluator produced a non-numeric result (e.g. IF(...,"yes","no"),
            // CONCATENATE(...)), classify as String — the same verdict Excel
            // writes once it caches the value (t="str"). Only override the
            // null-CellValue Number default; never touch an explicit DataType.
            if (cell.DataType?.HasValue != true
                && string.IsNullOrEmpty(rawCached)
                && !string.IsNullOrEmpty(computedValue)
                && !double.TryParse(computedValue, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out _))
            {
                node.Format["type"] = "String";
            }
            node.Format["evaluated"] = !string.IsNullOrEmpty(rawCached) || computedValue != null;
        }
        // Array formula readback — keys match Set input
        if (cell.CellFormula?.FormulaType?.Value == CellFormulaValues.Array)
        {
            node.Format["arrayformula"] = true;
            if (cell.CellFormula.Reference?.Value != null)
                node.Format["arrayref"] = cell.CellFormula.Reference.Value;
        }
        if (string.IsNullOrEmpty(displayText) && formula == null) node.Format["empty"] = true;

        // R8-3: phonetic guide readback. Surface the first <rPh>'s text so
        // CJK / Japanese authors writing furigana through `add cell --prop
        // phonetic=…` can verify the value round-trips.
        if (cell.DataType?.Value == CellValues.SharedString
            && int.TryParse(cell.CellValue?.Text, out var phSstIdx))
        {
            var phSst = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            var phSsi = phSst?.SharedStringTable?
                .Elements<SharedStringItem>().ElementAtOrDefault(phSstIdx);
            var firstRPh = phSsi?.Elements<PhoneticRun>().FirstOrDefault();
            if (firstRPh?.Text?.Text is { Length: > 0 } phText)
                node.Format["phonetic"] = phText;
        }

        // Hyperlink readback
        if (part != null)
        {
            var hyperlink = GetSheet(part).GetFirstChild<Hyperlinks>()?.Elements<Hyperlink>()
                .FirstOrDefault(h => h.Reference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true);
            if (hyperlink?.Id?.Value != null)
            {
                try
                {
                    var rel = part.HyperlinkRelationships.FirstOrDefault(r => r.Id == hyperlink.Id.Value);
                    if (rel != null)
                    {
                        var linkStr = rel.Uri.OriginalString;
                        // Strip trailing slash added by Uri normalization for bare authority URLs
                        if (linkStr.EndsWith("/") && rel.Uri.IsAbsoluteUri && rel.Uri.AbsolutePath == "/")
                            linkStr = linkStr.TrimEnd('/');
                        node.Format["link"] = linkStr;
                    }
                }
                catch { }
            }
            // Internal-location hyperlinks (Sheet1!B5, defined names) have no
            // external relationship — they live entirely in the @location
            // attribute. Without this branch, internal links round-trip
            // through Set but vanish from Get.
            // CONSISTENCY(internal-link-hash): OOXML stores the bare location
            // (no '#') but Set accepts (and documents) the '#'-prefixed form
            // — re-prepend '#' on readback so the value is directly usable as
            // `set link=...` input. External URL branch above stays unchanged.
            else if (hyperlink?.Location?.Value is { Length: > 0 } loc)
            {
                node.Format["link"] = "#" + loc;
            }
            // Tooltip readback — Set/Add persist tooltip= into the hyperlink's
            // @tooltip attribute, but without this emit the value was invisible
            // to Get and dropped by dump round-trips.
            if (hyperlink?.Tooltip?.Value is { Length: > 0 } hlTooltip)
                node.Format["tooltip"] = hlTooltip;

            // Display readback — the @display attribute is the friendly text
            // Excel shows for the link; without this emit it was invisible to
            // Get and silently dropped by dump round-trips.
            if (hyperlink?.Display?.Value is { Length: > 0 } hlDisplay)
                node.Format["display"] = hlDisplay;

            // Border readback from stylesheet
            var styleIndex = cell.StyleIndex?.Value ?? 0;
            var wbStylesPart = _doc.WorkbookPart?.WorkbookStylesPart;
            if (wbStylesPart?.Stylesheet != null && styleIndex > 0)
            {
                var cellFormats = wbStylesPart.Stylesheet.CellFormats;
                if (cellFormats != null && styleIndex < (uint)cellFormats.Elements<CellFormat>().Count())
                {
                    var xf = cellFormats.Elements<CellFormat>().ElementAt((int)styleIndex);
                    // Font readback
                    var fontId = xf.FontId?.Value ?? 0;
                    if (fontId > 0)
                    {
                        var fonts = wbStylesPart.Stylesheet.Fonts;
                        if (fonts != null && fontId < (uint)fonts.Elements<Font>().Count())
                        {
                            var font = fonts.Elements<Font>().ElementAt((int)fontId);
                            if (font.Bold != null) { node.Format["font.bold"] = true; }
                            if (font.Italic != null)
                            {
                                node.Format["font.italic"] = true;
                            }
                            // Canonical cell keys are `strike`/`underline` (schema cell.json);
                            // `font.strike`/`font.underline` are declared aliases. Get must
                            // normalize to the canonical key so set->get round-trips.
                            if (font.Strike != null) node.Format["strike"] = true;
                            if (font.Underline != null)
                            {
                                var uCanon = OfficeCli.Core.ExcelStyleManager.NormalizeStoredUnderline(font.Underline.Val?.InnerText);
                                if (uCanon != null) node.Format["underline"] = uCanon;
                            }
                            if (font.Color?.Rgb?.Value != null)
                                node.Format["font.color"] = ParseHelpers.FormatHexColor(font.Color.Rgb.Value);
                            else if (font.Color?.Theme?.Value != null)
                            {
                                var themeName = ParseHelpers.ExcelThemeIndexToName(font.Color.Theme.Value);
                                if (themeName != null) node.Format["font.color"] = themeName;
                            }
                            // vertAlign (superscript/subscript) readback. Canonical cell
                            // keys are `superscript`/`subscript` (schema cell.json + the
                            // Set cases); `font.superscript`/`font.subscript` are legacy
                            // aliases. Emit the canonical so set->get round-trips, and so
                            // this matches the rich-text run readback below.
                            var vertAlign = font.GetFirstChild<VerticalTextAlignment>();
                            if (vertAlign?.Val?.Value == VerticalAlignmentRunValues.Superscript)
                            {
                                node.Format["superscript"] = true;
                            }
                            else if (vertAlign?.Val?.Value == VerticalAlignmentRunValues.Subscript)
                            {
                                node.Format["subscript"] = true;
                            }
                            if (font.FontSize?.Val?.Value != null)
                                node.Format["font.size"] = $"{font.FontSize.Val.Value:0.##}pt";
                            if (font.FontName?.Val?.Value != null) node.Format["font.name"] = font.FontName.Val.Value;
                            // Long-tail Font children (charset, family, outline,
                            // shadow, condense, extend, scheme, ...). Emit as
                            // `font.<localName>` symmetric with the Set-side
                            // GetOrCreateFont longTailFontProps path.
                            FillUnknownDottedProps(font, node, "font.", CuratedFontChildren);
                        }
                    }

                    // Fill readback
                    var fillId = xf.FillId?.Value ?? 0;
                    if (fillId > 0)
                    {
                        var fills = wbStylesPart.Stylesheet.Fills;
                        if (fills != null && fillId < (uint)fills.Elements<Fill>().Count())
                        {
                            var fill = fills.Elements<Fill>().ElementAt((int)fillId);
                            // Check gradient fill first
                            var gf = fill.GetFirstChild<GradientFill>();
                            if (gf != null)
                            {
                                var stops = gf.Elements<GradientStop>().ToList();
                                if (stops.Count >= 2)
                                {
                                    var validColors = stops
                                        .Select(s => s.Color?.Rgb?.Value)
                                        .Where(v => !string.IsNullOrEmpty(v))
                                        .Select(v => ParseHelpers.FormatHexColor(v!))
                                        .ToList();
                                    if (validColors.Count >= 2)
                                    {
                                        var colorParts = string.Join(";", validColors);
                                        int deg = (int)(gf.Degree?.Value ?? 0);
                                        node.Format["fill"] = $"gradient;{colorParts};{deg}";
                                    }
                                }
                            }
                            else
                            {
                                var pf = fill.PatternFill;
                                // Non-solid pattern fills (lightGray, darkGrid, …)
                                // carry a pattern type plus fg/bg colors that must
                                // all survive round-trip. Solid/none keep the
                                // legacy single-color `fill` readback so existing
                                // behavior is unchanged.
                                var patType = pf?.PatternType?.Value;
                                bool nonSolidPattern = patType != null
                                    && patType != PatternValues.Solid
                                    && patType != PatternValues.None;
                                if (nonSolidPattern)
                                {
                                    // Emit the OOXML wire name (e.g. "lightGray"),
                                    // not the enum struct's ToString.
                                    node.Format["fillPattern"] = pf!.PatternType!.InnerText;
                                    // CONSISTENCY(scheme-color): theme fg/bg read
                                    // back as scheme names, mirroring the solid
                                    // path below — Set accepts them, so a pattern
                                    // fill with theme colors must round-trip
                                    // through dump→batch instead of silently
                                    // dropping both colors.
                                    if (pf!.ForegroundColor?.Rgb?.Value != null)
                                        node.Format["fill"] = ParseHelpers.FormatHexColor(pf.ForegroundColor.Rgb.Value);
                                    else if (pf.ForegroundColor?.Theme?.Value != null
                                        && ParseHelpers.ExcelThemeIndexToName(pf.ForegroundColor.Theme.Value) is { } fgTheme)
                                        node.Format["fill"] = fgTheme;
                                    if (pf.BackgroundColor?.Rgb?.Value != null)
                                        node.Format["fillBg"] = ParseHelpers.FormatHexColor(pf.BackgroundColor.Rgb.Value);
                                    else if (pf.BackgroundColor?.Theme?.Value != null
                                        && ParseHelpers.ExcelThemeIndexToName(pf.BackgroundColor.Theme.Value) is { } bgTheme)
                                        node.Format["fillBg"] = bgTheme;
                                }
                                else if (pf?.ForegroundColor?.Rgb?.Value != null)
                                    node.Format["fill"] = ParseHelpers.FormatHexColor(pf.ForegroundColor.Rgb.Value);
                                else if (pf?.ForegroundColor?.Theme?.Value != null)
                                {
                                    var themeName = ParseHelpers.ExcelThemeIndexToName(pf.ForegroundColor.Theme.Value);
                                    if (themeName != null) node.Format["fill"] = themeName;
                                }
                            }
                        }
                    }

                    var borderId = xf.BorderId?.Value ?? 0;
                    if (borderId > 0)
                    {
                        var borders = wbStylesPart.Stylesheet.Borders;
                        if (borders != null && borderId < (uint)borders.Elements<Border>().Count())
                        {
                            var border = borders.Elements<Border>().ElementAt((int)borderId);
                            var sides = new (string name, BorderPropertiesType? bp)[] {
                                ("left", border.LeftBorder), ("right", border.RightBorder),
                                ("top", border.TopBorder), ("bottom", border.BottomBorder)
                            };
                            foreach (var (side, b) in sides)
                            {
                                if (b?.Style?.Value != null && b.Style.Value != BorderStyleValues.None)
                                {
                                    node.Format[$"border.{side}"] = b.Style.InnerText;
                                    if (b.Color?.Rgb?.Value != null)
                                        node.Format[$"border.{side}.color"] = ParseHelpers.FormatHexColor(b.Color.Rgb.Value!);
                                }
                            }
                            // Diagonal border readback
                            var diag = border.DiagonalBorder;
                            if (diag?.Style?.Value != null && diag.Style.Value != BorderStyleValues.None)
                            {
                                node.Format["border.diagonal"] = diag.Style.InnerText;
                                if (diag.Color?.Rgb?.Value != null)
                                    node.Format["border.diagonal.color"] = ParseHelpers.FormatHexColor(diag.Color.Rgb.Value!);
                            }
                            if (border.DiagonalUp?.Value == true)
                                node.Format["border.diagonalUp"] = true;
                            if (border.DiagonalDown?.Value == true)
                                node.Format["border.diagonalDown"] = true;
                        }
                    }

                    // Alignment + wrap readback
                    var alignment = xf.Alignment;
                    if (alignment != null)
                    {
                        if (alignment.WrapText?.Value == true)
                            node.Format["alignment.wrapText"] = true;
                        if (alignment.Horizontal?.HasValue == true)
                            node.Format["alignment.horizontal"] = alignment.Horizontal.InnerText;
                        if (alignment.Vertical?.HasValue == true)
                        {
                            node.Format["alignment.vertical"] = alignment.Vertical.InnerText;
                        }
                        if (alignment.TextRotation?.HasValue == true && alignment.TextRotation.Value != 0)
                            node.Format["alignment.textRotation"] = alignment.TextRotation.Value.ToString();
                        if (alignment.Indent?.HasValue == true && alignment.Indent.Value > 0)
                            node.Format["alignment.indent"] = alignment.Indent.Value.ToString();
                        if (alignment.ShrinkToFit?.Value == true)
                            node.Format["alignment.shrinkToFit"] = true;
                        // DEFERRED(xlsx/cell-reading-order) CE10 — canonical
                        // readback as string form (context/ltr/rtl).
                        if (alignment.ReadingOrder?.HasValue == true && alignment.ReadingOrder.Value != 0)
                        {
                            node.Format["alignment.readingOrder"] = alignment.ReadingOrder.Value switch
                            {
                                1u => "ltr",
                                2u => "rtl",
                                _ => "context"
                            };
                        }
                        // Long-tail Alignment attributes (justifyLastLine,
                        // relativeIndent, ...). Symmetric with Set's default
                        // branch in ExcelStyleManager.ApplyStyle alignment loop.
                        FillUnknownAttrProps(alignment, node, "alignment.", CuratedAlignmentAttrs);
                    }

                    // Protection readback — both curated locked/hidden and any
                    // long-tail Protection attribute symmetric with Set.
                    var xfProt = xf.Protection;
                    if (xfProt != null)
                    {
                        if (xfProt.Locked?.Value != null) node.Format["protection.locked"] = xfProt.Locked.Value;
                        if (xfProt.Hidden?.Value != null) node.Format["protection.hidden"] = xfProt.Hidden.Value;
                        FillUnknownAttrProps(xfProt, node, "protection.", CuratedProtectionAttrs);
                    }

                    // R29: quotePrefix readback (set by leading apostrophe text mode)
                    if (xf.QuotePrefix?.Value == true)
                        node.Format["quotePrefix"] = true;

                    // Number format readback
                    var numFmtId = xf.NumberFormatId?.Value ?? 0;
                    if (numFmtId > 0)
                    {
                        node.Format["numFmtId"] = (int)numFmtId;
                        var numFmts = wbStylesPart.Stylesheet.NumberingFormats;
                        var customFmt = numFmts?.Elements<NumberingFormat>()
                            .FirstOrDefault(nf => nf.NumberFormatId?.Value == numFmtId);
                        object fmtVal;
                        if (customFmt?.FormatCode?.Value != null)
                            fmtVal = customFmt.FormatCode.Value;
                        else
                        {
                            // Resolve built-in number format IDs to their format strings
                            // See ECMA-376 Part 1, 18.8.30 (numFmt) for built-in IDs
                            fmtVal = numFmtId switch
                            {
                                1 => "0",
                                2 => "0.00",
                                3 => "#,##0",
                                4 => "#,##0.00",
                                9 => "0%",
                                10 => "0.00%",
                                11 => "0.00E+00",
                                12 => "# ?/?",
                                13 => "# ??/??",
                                14 => "m/d/yy",
                                15 => "d-mmm-yy",
                                16 => "d-mmm",
                                17 => "mmm-yy",
                                18 => "h:mm AM/PM",
                                19 => "h:mm:ss AM/PM",
                                20 => "h:mm",
                                21 => "h:mm:ss",
                                22 => "m/d/yy h:mm",
                                37 => "#,##0 ;(#,##0)",
                                38 => "#,##0 ;[Red](#,##0)",
                                39 => "#,##0.00;(#,##0.00)",
                                40 => "#,##0.00;[Red](#,##0.00)",
                                45 => "mm:ss",
                                46 => "[h]:mm:ss",
                                47 => "mmss.0",
                                48 => "##0.0E+0",
                                49 => "@",
                                _ => (object)(int)numFmtId // fallback to ID for truly unknown formats
                            };
                        }
                        node.Format["numberformat"] = fmtVal;
                    }

                    // Protection readback handled above via the dotted
                    // canonical form (`protection.locked` / `protection.hidden`)
                    // — see CONSISTENCY(canonical-keys) in the project conventions. Flat
                    // `locked` / `formulahidden` Get emission was removed to
                    // avoid double-emission alongside the dotted form. The
                    // Set side still accepts both flat shorthand and dotted
                    // input via IsStyleKey routing.
                }
            }

            // Merge cell readback
            var mergeCells = GetSheet(part).GetFirstChild<MergeCells>();
            if (mergeCells != null)
            {
                var mergeCell = mergeCells.Elements<MergeCell>()
                    .FirstOrDefault(m => IsCellInMergeRange(cellRef, m.Reference?.Value));
                if (mergeCell != null)
                {
                    var mergeRef = mergeCell.Reference?.Value ?? "";
                    node.Format["merge"] = mergeRef;
                    // Indicate if this cell is the top-left anchor of the merged range
                    if (mergeRef.Split(':')[0].Equals(cellRef, StringComparison.OrdinalIgnoreCase))
                        node.Format["mergeAnchor"] = true;
                }
            }
        }

        // Rich text (SST runs) readback
        if (cell.DataType?.Value == CellValues.SharedString &&
            int.TryParse(cell.CellValue?.Text, out var sstIdx2))
        {
            var sst2 = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            var ssi2 = sst2?.SharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(sstIdx2);
            if (ssi2 != null)
            {
                var runs = ssi2.Elements<Run>().ToList();
                if (runs.Count > 0)
                {
                    node.Format["richtext"] = true;
                    node.ChildCount = runs.Count;
                    int runI = 1;
                    foreach (var run in runs)
                    {
                        node.Children.Add(RunToNode(run, $"/{sheetName}/{cellRef}/run[{runI}]"));
                        runI++;
                    }
                }
            }
        }

        return node;
    }

    private static DocumentNode RunToNode(Run run, string path)
    {
        var runNode = new DocumentNode { Path = path, Type = "run", Text = run.Text?.Text ?? "" };
        var rp = run.RunProperties;
        if (rp != null)
        {
            if (rp.GetFirstChild<Bold>() != null) runNode.Format["bold"] = true;
            if (rp.GetFirstChild<Italic>() != null) runNode.Format["italic"] = true;
            if (rp.GetFirstChild<Strike>() != null) runNode.Format["strike"] = true;
            var ul = rp.GetFirstChild<Underline>();
            if (ul != null) runNode.Format["underline"] = ul.Val?.InnerText == "double" ? "double" : "single";
            var va = rp.GetFirstChild<VerticalTextAlignment>();
            if (va?.Val?.Value == VerticalAlignmentRunValues.Superscript) runNode.Format["superscript"] = true;
            if (va?.Val?.Value == VerticalAlignmentRunValues.Subscript) runNode.Format["subscript"] = true;
            if (rp.GetFirstChild<FontSize>()?.Val?.Value != null)
                runNode.Format["size"] = $"{rp.GetFirstChild<FontSize>()!.Val!.Value:0.##}pt";
            if (rp.GetFirstChild<Color>()?.Rgb?.Value != null)
                runNode.Format["color"] = ParseHelpers.FormatHexColor(rp.GetFirstChild<Color>()!.Rgb!.Value!);
            if (rp.GetFirstChild<RunFont>()?.Val?.Value != null)
                runNode.Format["font"] = rp.GetFirstChild<RunFont>()!.Val!.Value!;
        }
        return runNode;
    }

    // CONSISTENCY(xlsx/comment-font): C8 — build the <x:rPr> for comment runs.
    // When no font.* properties are supplied, keep the legacy Tahoma 9 /
    // indexed-81 default for back-compat. When any font.* is present, honor
    // them and fall back to the defaults only for unspecified facets.
    // Input vocabulary mirrors the cell-level font handling: font.bold,
    // font.italic, font.underline (single|double), font.size (pt-qualified
    // or bare), font.color (#FF0000 / FF0000 / rgb() / named), font.name.
    internal static RunProperties BuildCommentRunProperties(Dictionary<string, string> properties)
    {
        // CONSISTENCY(xlsx/comment-rtl): R9-3 — direction/dir/font.rtl propagate
        // <x:rtl/> on CT_RPrElt. We accept either a top-level direction key
        // (mirrors the rest of the i18n surface) or the explicit font.rtl
        // boolean. The flag is independent of font.* defaults — a comment
        // with only direction=rtl must still keep the legacy Tahoma 9 default
        // for the font facets, just with an additional <x:rtl/> child.
        bool wantsRtl = false;
        bool hasExplicitRtl = false;
        if (properties.TryGetValue("direction", out var dirRaw)
            || properties.TryGetValue("dir", out dirRaw))
        {
            hasExplicitRtl = true;
            wantsRtl = string.Equals(dirRaw, "rtl", StringComparison.OrdinalIgnoreCase);
        }
        if (properties.TryGetValue("font.rtl", out var fRtl))
        {
            hasExplicitRtl = true;
            wantsRtl = IsTruthy(fRtl);
        }

        bool hasAnyFont = properties.Keys.Any(k =>
            k.StartsWith("font.", StringComparison.OrdinalIgnoreCase));
        if (!hasAnyFont && !hasExplicitRtl)
        {
            return new RunProperties(
                new FontSize { Val = 9 },
                new Color { Indexed = 81 },
                new RunFont { Val = "Tahoma" });
        }
        // CT_RPrElt has no schema-level <rtl> child; we synthesize one as an
        // unknown element using the Spreadsheet namespace so consumers that
        // honor the i18n extension (Excel for Mac / RTL locales) pick it up.
        // The element is a leaf empty marker; absence means LTR (default).
        const string xlsNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        OpenXmlElement BuildRtlMarker() =>
            new DocumentFormat.OpenXml.OpenXmlUnknownElement("rtl", xlsNs);

        if (!hasAnyFont && hasExplicitRtl)
        {
            var rPrDefault = new RunProperties();
            if (wantsRtl) rPrDefault.AppendChild(BuildRtlMarker());
            rPrDefault.AppendChild(new FontSize { Val = 9 });
            rPrDefault.AppendChild(new Color { Indexed = 81 });
            rPrDefault.AppendChild(new RunFont { Val = "Tahoma" });
            return rPrDefault;
        }

        var rPr = new RunProperties();
        if (wantsRtl) rPr.AppendChild(BuildRtlMarker());
        if (properties.TryGetValue("font.bold", out var fb) && IsTruthy(fb))
            rPr.AppendChild(new Bold());
        if (properties.TryGetValue("font.italic", out var fi) && IsTruthy(fi))
            rPr.AppendChild(new Italic());
        if (properties.TryGetValue("font.underline", out var fu) && !string.IsNullOrEmpty(fu)
            && !string.Equals(fu, "none", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(fu, "false", StringComparison.OrdinalIgnoreCase))
        {
            var uVal = string.Equals(fu, "double", StringComparison.OrdinalIgnoreCase)
                ? UnderlineValues.Double : UnderlineValues.Single;
            rPr.AppendChild(new Underline { Val = uVal });
        }
        // Size default 9pt
        var sizePt = properties.TryGetValue("font.size", out var fs)
            ? ParseHelpers.ParseFontSize(fs) : 9.0;
        rPr.AppendChild(new FontSize { Val = sizePt });
        // Color: explicit overrides default indexed=81
        if (properties.TryGetValue("font.color", out var fc) && !string.IsNullOrWhiteSpace(fc))
            rPr.AppendChild(new Color { Rgb = ParseHelpers.NormalizeArgbColor(fc) });
        else
            rPr.AppendChild(new Color { Indexed = 81 });
        // Name default Tahoma
        var fontName = properties.TryGetValue("font.name", out var fn) && !string.IsNullOrWhiteSpace(fn)
            ? fn : "Tahoma";
        rPr.AppendChild(new RunFont { Val = fontName });
        return rPr;
    }

    // Serialize a multi-run comment's runs into the `runs=<json>` array
    // BuildCommentTextFromRuns consumes. Vocabulary mirrors SerializeRichTextRuns.
    internal static string SerializeCommentRuns(List<Run> runs)
    {
        var arr = new System.Text.Json.Nodes.JsonArray();
        foreach (var run in runs)
        {
            var o = new System.Text.Json.Nodes.JsonObject { ["text"] = run.Text?.Text ?? "" };
            var rp = run.RunProperties;
            if (rp != null)
            {
                if (rp.GetFirstChild<Bold>() != null) o["bold"] = true;
                if (rp.GetFirstChild<Italic>() != null) o["italic"] = true;
                if (rp.GetFirstChild<Strike>() != null) o["strike"] = true;
                var ul = rp.GetFirstChild<Underline>();
                if (ul != null) o["underline"] = ul.Val?.InnerText == "double" ? "double" : "single";
                var va = rp.GetFirstChild<VerticalTextAlignment>();
                if (va?.Val?.Value == VerticalAlignmentRunValues.Superscript) o["superscript"] = true;
                if (va?.Val?.Value == VerticalAlignmentRunValues.Subscript) o["subscript"] = true;
                var sz = rp.GetFirstChild<FontSize>();
                if (sz?.Val?.Value != null) o["size"] = $"{sz.Val.Value:0.##}pt";
                var clr = rp.GetFirstChild<Color>();
                if (clr?.Rgb?.Value != null) o["color"] = ParseHelpers.FormatHexColor(clr.Rgb.Value!);
                var rf = rp.GetFirstChild<RunFont>();
                if (rf?.Val?.Value != null) o["font"] = rf.Val.Value!;
            }
            // Non-generic Add(JsonNode) — the generic Add<T> overload carries
            // RequiresUnreferencedCode (IL2026) though it never serializes a JsonNode.
            arr.Add((System.Text.Json.Nodes.JsonNode)o);
        }
        return arr.ToJsonString();
    }

    // Build a CommentText from a `runs=<json array>` value, one <r> per run
    // carrying its own <rPr>. Run vocabulary mirrors rich-text cells
    // (bold/italic/strike/underline/superscript/subscript/size/color/font) so
    // the two paths share one input contract. Comment runs keep the Tahoma-9
    // indexed-81 default for facets a run leaves unspecified.
    internal static CommentText BuildCommentTextFromRuns(string runsJson)
    {
        System.Text.Json.Nodes.JsonArray? arr;
        try { arr = System.Text.Json.Nodes.JsonNode.Parse(runsJson) as System.Text.Json.Nodes.JsonArray; }
        catch (System.Text.Json.JsonException ex)
        {
            throw new ArgumentException($"comment runs: invalid JSON array — {ex.Message}");
        }
        if (arr == null)
            throw new ArgumentException("comment runs: value must be a JSON array");

        var ct = new CommentText();
        foreach (var item in arr)
        {
            if (item is not System.Text.Json.Nodes.JsonObject o) continue;
            var text = o["text"]?.GetValue<string>() ?? "";
            OfficeCli.Core.ParseHelpers.ValidateXmlText(text, "comment run text");

            var rPr = new RunProperties();
            bool RunBool(string key) => o[key] is { } n
                && (n.GetValueKind() == System.Text.Json.JsonValueKind.True
                    || (n.GetValueKind() == System.Text.Json.JsonValueKind.String && IsTruthy(n.GetValue<string>())));

            if (RunBool("bold")) rPr.AppendChild(new Bold());
            if (RunBool("italic")) rPr.AppendChild(new Italic());
            if (RunBool("strike")) rPr.AppendChild(new Strike());
            var uStr = o["underline"]?.GetValue<string>();
            if (!string.IsNullOrEmpty(uStr) && !string.Equals(uStr, "none", StringComparison.OrdinalIgnoreCase))
            {
                rPr.AppendChild(new Underline
                {
                    Val = string.Equals(uStr, "double", StringComparison.OrdinalIgnoreCase)
                        ? UnderlineValues.Double : UnderlineValues.Single
                });
            }
            if (RunBool("superscript"))
                rPr.AppendChild(new VerticalTextAlignment { Val = VerticalAlignmentRunValues.Superscript });
            else if (RunBool("subscript"))
                rPr.AppendChild(new VerticalTextAlignment { Val = VerticalAlignmentRunValues.Subscript });

            var szStr = o["size"]?.GetValue<string>();
            rPr.AppendChild(new FontSize { Val = szStr != null ? ParseHelpers.ParseFontSize(szStr) : 9.0 });

            var colStr = o["color"]?.GetValue<string>();
            if (!string.IsNullOrWhiteSpace(colStr))
                rPr.AppendChild(new Color { Rgb = ParseHelpers.NormalizeArgbColor(colStr) });
            else
                rPr.AppendChild(new Color { Indexed = 81 });

            var fontStr = o["font"]?.GetValue<string>();
            rPr.AppendChild(new RunFont { Val = string.IsNullOrWhiteSpace(fontStr) ? "Tahoma" : fontStr });

            ct.AppendChild(new Run(rPr,
                new Text(text.Replace("\r\n", "\n")) { Space = SpaceProcessingModeValues.Preserve }));
        }
        return ct;
    }

    // ==================== Data Validation Helpers ====================

    private DocumentNode TableToNode(string sheetName, WorksheetPart worksheetPart, int tableIndex, int depth)
    {
        var tableParts = worksheetPart.TableDefinitionParts.ToList();
        if (tableIndex < 1 || tableIndex > tableParts.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range (1..{tableParts.Count})");

        var tbl = tableParts[tableIndex - 1].Table
            ?? throw new ArgumentException($"Table {tableIndex} has no definition");

        var node = new DocumentNode
        {
            Path = $"/{sheetName}/table[{tableIndex}]",
            Type = "table",
            Text = tbl.DisplayName?.Value ?? tbl.Name?.Value ?? $"Table{tableIndex}",
            Preview = $"{tbl.Name?.Value} ({tbl.Reference?.Value})"
        };

        node.Format["name"] = tbl.Name?.Value ?? "";
        node.Format["displayName"] = tbl.DisplayName?.Value ?? "";
        node.Format["ref"] = tbl.Reference?.Value ?? "";

        var styleInfo = tbl.GetFirstChild<TableStyleInfo>();
        if (styleInfo?.Name?.Value != null)
            node.Format["style"] = styleInfo.Name.Value;
        if (styleInfo != null)
        {
            // BUG-R4-03/04: cross-format canonical key alignment with docx/pptx.
            // Get emits camelCase canonical (bandedRows/bandedCols/firstCol/lastCol).
            // Set still accepts the OOXML-internal aliases (showRowStripes etc).
            if (styleInfo.ShowRowStripes is not null) node.Format["bandedRows"] = styleInfo.ShowRowStripes.Value;
            if (styleInfo.ShowColumnStripes is not null) node.Format["bandedCols"] = styleInfo.ShowColumnStripes.Value;
            if (styleInfo.ShowFirstColumn is not null) node.Format["firstCol"] = styleInfo.ShowFirstColumn.Value;
            if (styleInfo.ShowLastColumn is not null) node.Format["lastCol"] = styleInfo.ShowLastColumn.Value;
        }

        node.Format["headerRow"] = (tbl.HeaderRowCount?.Value ?? 1) != 0;
        node.Format["totalRow"] = (tbl.TotalsRowCount?.Value ?? 0) > 0 || (tbl.TotalsRowShown?.Value ?? false);

        var tableColumns = tbl.GetFirstChild<TableColumns>();
        if (tableColumns != null)
        {
            var colNames = tableColumns.Elements<TableColumn>()
                .Select(c => c.Name?.Value ?? "").ToArray();
            node.Format["columns"] = string.Join(",", colNames);
            node.ChildCount = colNames.Length;
        }

        return node;
    }

    private DocumentNode CommentToNode(string sheetName, Comment comment, Comments comments, int index)
    {
        var reference = comment.Reference?.Value ?? "?";
        var text = comment.CommentText?.InnerText ?? "";
        var authorId = comment.AuthorId?.Value ?? 0;

        var authors = comments.GetFirstChild<Authors>();
        var authorName = authors?.Elements<Author>().ElementAtOrDefault((int)authorId)?.Text ?? "Unknown";

        var node = new DocumentNode
        {
            Path = $"/{sheetName}/comment[{index}]",
            Type = "comment",
            Text = text,
            Preview = $"{reference}: {text}"
        };

        node.Format["ref"] = reference;
        node.Format["author"] = authorName;
        node.Format["anchoredTo"] = $"/{sheetName}/{reference}";

        // CONSISTENCY(xlsx/comment-font): C8 — surface font.* from first run's
        // rPr so Query/Get round-trips the Add-time formatting. Only report
        // non-default facets so Tahoma-9-indexed-81 comments stay unadorned.
        var allRuns = comment.CommentText?.Elements<Run>().ToList() ?? new List<Run>();
        // Per-run rich formatting: a comment with 2+ runs carries distinct
        // per-run <rPr>. Surfacing only the first run's font.* (legacy path)
        // collapsed the formatting on round-trip. Emit a `runs=<json>` carrier
        // (same vocabulary as rich-text cells) so Add can rebuild every run.
        // Single-run comments keep the legacy font.* readback unchanged.
        if (allRuns.Count > 1)
        {
            node.Format["runs"] = SerializeCommentRuns(allRuns);
        }
        var firstRun = allRuns.FirstOrDefault();
        var rProps = allRuns.Count <= 1 ? firstRun?.RunProperties : null;
        if (rProps != null)
        {
            if (rProps.Elements<Bold>().Any()) node.Format["font.bold"] = true;
            if (rProps.Elements<Italic>().Any()) node.Format["font.italic"] = true;
            var u = rProps.Elements<Underline>().FirstOrDefault();
            if (u != null)
                node.Format["font.underline"] = u.Val?.InnerText == "double" ? "double" : "single";
            var clr = rProps.Elements<Color>().FirstOrDefault();
            if (clr?.Rgb?.HasValue == true)
                node.Format["font.color"] = ParseHelpers.FormatHexColor(clr.Rgb.Value!);
            var sz = rProps.Elements<FontSize>().FirstOrDefault();
            if (sz?.Val?.HasValue == true && sz.Val.Value != 9.0)
                node.Format["font.size"] = $"{sz.Val.Value:0.##}pt";
            var rf = rProps.Elements<RunFont>().FirstOrDefault();
            if (rf?.Val?.HasValue == true && rf.Val.Value != "Tahoma")
                node.Format["font.name"] = rf.Val.Value;
        }

        return node;
    }

    private static DocumentNode DataValidationToNode(string sheetName, DataValidation dv, int index)
    {
        var sqref = dv.SequenceOfReferences?.InnerText ?? "";
        var node = new DocumentNode
        {
            Path = $"/{sheetName}/dataValidation[{index}]",
            Type = "dataValidation",
            Text = sqref,
            Preview = $"dataValidation[{index}] ({sqref})"
        };

        // CONSISTENCY(canonical-key): schema canonical key is 'ref', not 'sqref'.
        node.Format["ref"] = sqref;

        if (dv.Type?.HasValue == true)
            node.Format["type"] = dv.Type.InnerText;
        if (dv.Operator?.HasValue == true)
            node.Format["operator"] = dv.Operator.InnerText;

        if (dv.Formula1 != null)
        {
            // Preserve formula1 exactly as stored in XML so query→set round-trips:
            // list-type validations wrap literal options in "..." at Add time, and
            // stripping those quotes here made set(formula1=<stripped>) treat the
            // whole list as a single item. See DEFERRED(xlsx/validation-list-formula-roundtrip).
            node.Format["formula1"] = dv.Formula1.Text ?? "";
        }

        if (dv.Formula2 != null)
            node.Format["formula2"] = dv.Formula2.Text ?? "";

        if (dv.AllowBlank?.HasValue == true)
            node.Format["allowBlank"] = dv.AllowBlank.Value;
        if (dv.ShowErrorMessage?.HasValue == true)
            node.Format["showError"] = dv.ShowErrorMessage.Value;
        if (dv.ShowInputMessage?.HasValue == true)
            node.Format["showInput"] = dv.ShowInputMessage.Value;

        if (!string.IsNullOrEmpty(dv.ErrorTitle?.Value))
            node.Format["errorTitle"] = dv.ErrorTitle!.Value!;
        if (!string.IsNullOrEmpty(dv.Error?.Value))
            node.Format["error"] = dv.Error!.Value!;
        if (!string.IsNullOrEmpty(dv.PromptTitle?.Value))
            node.Format["promptTitle"] = dv.PromptTitle!.Value!;
        if (!string.IsNullOrEmpty(dv.Prompt?.Value))
            node.Format["prompt"] = dv.Prompt!.Value!;

        if (dv.ErrorStyle?.HasValue == true)
            node.Format["errorStyle"] = dv.ErrorStyle.InnerText;

        // CONSISTENCY(validation-incelldropdown): Add accepts inCellDropdown
        // (user-friendly sense; OOXML stores the inverse showDropDown).
        // Get must surface the same key so help-doc [add/get] is honored.
        // OOXML default: showDropDown attribute absent => dropdown is shown
        // (inCellDropdown=true). showDropDown=true means hide arrow
        // (inCellDropdown=false). Always emit so round-trip is symmetric.
        node.Format["inCellDropdown"] = !(dv.ShowDropDown?.Value ?? false);

        return node;
    }

    /// <summary>
    /// Reorder RunProperties children to match CT_RPrElt schema order:
    /// b, i, strike, condense, extend, outline, shadow, u, vertAlign, sz, color, rFont, family, charset, scheme
    /// </summary>
    private static void ReorderRunProperties(RunProperties rpr)
    {
        if (rpr == null || !rpr.HasChildren) return;
        var children = rpr.ChildElements.ToList();
        var ordered = children.OrderBy(c => GetRunPropertyOrder(c)).ToList();
        rpr.RemoveAllChildren();
        foreach (var child in ordered) rpr.AppendChild(child);
    }

    private static int GetRunPropertyOrder(DocumentFormat.OpenXml.OpenXmlElement element) => element switch
    {
        Bold => 0,
        Italic => 1,
        Strike => 2,
        Condense => 3,
        Extend => 4,
        Outline => 5,
        Shadow => 6,
        Underline => 7,
        VerticalTextAlignment => 8,
        FontSize => 9,
        Color => 10,
        RunFont => 11,
        FontFamily => 12,
        RunPropertyCharSet => 13,
        FontScheme => 14,
        _ => 99
    };
}
