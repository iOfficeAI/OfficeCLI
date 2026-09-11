// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using ThreadedCmt = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;


namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    private static void InsertSheetViewsInSchemaOrder(Worksheet ws, SheetViews sheetViews)
    {
        OpenXmlElement? anchor = ws.GetFirstChild<SheetDimension>()
            ?? (OpenXmlElement?)ws.GetFirstChild<SheetProperties>();
        if (anchor != null)
            ws.InsertAfter(sheetViews, anchor);
        else
            ws.InsertAt(sheetViews, 0);
    }

    // ==================== Sheet-level Set (freeze panes) ====================

    private List<string> SetSheetLevel(WorksheetPart worksheet, string sheetName, Dictionary<string, string> properties)
    {
        // Find & Replace at sheet level
        if (properties.TryGetValue("find", out var findText) && properties.TryGetValue("replace", out var replaceText))
        {
            var count = FindAndReplace(findText, replaceText, worksheet);
            LastFindMatchCount = count;
            var remaining = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
            remaining.Remove("find");
            remaining.Remove("replace");
            if (remaining.Count > 0)
                return SetSheetLevel(worksheet, sheetName, remaining);
            return [];
        }

        var unsupported = new List<string>();
        var ws = GetSheet(worksheet);

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "name":
                {
                    // Validate sheet name up-front so Excel doesn't reject the file
                    // on open. Rules per Excel:
                    //   - cannot be empty / blank
                    //   - max 31 chars
                    //   - cannot contain  \  /  ?  *  :  [  ]
                    //   - cannot start or end with apostrophe '
                    //   - cannot equal reserved name "History"
                    ValidateSheetName(value);

                    // Rename the sheet
                    var workbook = GetWorkbook();
                    var sheets = workbook.Sheets?.Elements<Sheet>().ToList();
                    var sheet = sheets?.FirstOrDefault(s =>
                        s.Name?.Value?.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true);
                    if (sheet != null)
                    {
                        var oldName = sheet.Name!.Value!;
                        // R35-1: Excel sheet names are case-insensitive and must be
                        // unique. Match the Add path's duplicate-name check
                        // (ExcelHandler.Add.Cells.cs) so renaming Sheet1→Data when a
                        // "Data" sheet already exists fails up-front rather than
                        // writing two <sheet name="Data"> entries.
                        // CONSISTENCY(sheet-name-unique)
                        if (!oldName.Equals(value, StringComparison.OrdinalIgnoreCase) &&
                            sheets!.Any(s => s != sheet &&
                                s.Name?.Value?.Equals(value, StringComparison.OrdinalIgnoreCase) == true))
                        {
                            throw new ArgumentException(
                                $"A sheet named '{value}' already exists. Sheet names must be unique.");
                        }
                        sheet.Name = value;

                        // Excel stores sheet references in formulas as either:
                        //   SimpleSheetName!A1      (no spaces/special chars)
                        //   'Sheet With Spaces'!A1  (name with spaces or special chars)
                        static bool NeedsQuoting(string n) =>
                            n.Any(c => char.IsWhiteSpace(c) || c is '\'' or '[' or ']' or ':' or '*' or '?' or '/' or '\\');
                        // BUG R4-B: ECMA-376 §18.17 requires inner apostrophes to be
                        // doubled inside a quoted sheet identifier — e.g. "Bob's Sheet"
                        // serializes as 'Bob''s Sheet'!A1. Without escaping, the
                        // resulting formula text is parser-ambiguous (Excel can read
                        // it but a strict tokenizer treats the lone apostrophe as the
                        // closing quote and corrupts the reference).
                        static string FormulaRef(string n) => NeedsQuoting(n) ? $"'{n.Replace("'", "''")}'" : n;

                        var oldRef = FormulaRef(oldName) + "!";
                        var newRef = FormulaRef(value) + "!";

                        // Update named range references
                        var definedNames = workbook.GetFirstChild<DefinedNames>();
                        if (definedNames != null)
                        {
                            foreach (var dn in definedNames.Elements<DefinedName>())
                            {
                                if (dn.Text != null && dn.Text.Contains(oldRef, StringComparison.OrdinalIgnoreCase))
                                    dn.Text = dn.Text.Replace(oldRef, newRef, StringComparison.OrdinalIgnoreCase);
                            }
                        }
                        // Update formula references in all cells across all sheets
                        foreach (var (_, wsPart) in GetWorksheets())
                        {
                            var sd = GetSheet(wsPart).GetFirstChild<SheetData>();
                            if (sd == null) continue;
                            foreach (var cell in sd.Descendants<Cell>())
                            {
                                if (cell.CellFormula?.Text != null &&
                                    cell.CellFormula.Text.Contains(oldRef, StringComparison.OrdinalIgnoreCase))
                                {
                                    // R3 BUG-2: must skip string literals — INDIRECT("Sheet1!A1")
                                    // is a user-typed string, not a reference, and Excel preserves
                                    // it verbatim across renames.
                                    cell.CellFormula.Text = Core.FormulaRefShifter.RenameSheetRef(
                                        cell.CellFormula.Text, oldRef, newRef);
                                }
                            }
                            GetSheet(wsPart).Save();
                        }

                        // Update any pivot cache definitions whose WorksheetSource
                        // references the old sheet name. Without this the pivot
                        // cache's stale sheet ref breaks Excel refresh.
                        // CONSISTENCY(sheet-rename-refs)
                        var workbookPart = _doc.WorkbookPart!;
                        foreach (var cacheDefPart in workbookPart.GetPartsOfType<PivotTableCacheDefinitionPart>())
                        {
                            var wsSource = cacheDefPart.PivotCacheDefinition?.CacheSource?.WorksheetSource;
                            if (wsSource?.Sheet?.Value != null &&
                                wsSource.Sheet.Value.Equals(oldName, StringComparison.OrdinalIgnoreCase))
                            {
                                wsSource.Sheet = value;
                                cacheDefPart.PivotCacheDefinition!.Save();
                            }
                        }

                        // CONSISTENCY(sheet-rename-refs): chart series formulas
                        // (<c:f>SheetName!$A$1:$B$2</c:f>) must follow the
                        // rename or Excel reopens the file with an "external
                        // links" warning, treating the orphan SheetName!
                        // prefix as a pointer to a separate workbook. Walk
                        // every WorksheetPart's drawing → chart parts and
                        // rewrite the formula text in-place. Both quoted
                        // ('Sheet With Spaces'!) and bare (Sheet1!) forms
                        // are handled because oldRef/newRef already include
                        // the trailing '!' and quoting decision.
                        foreach (var anyWsPart in workbookPart.WorksheetParts)
                        {
                            if (anyWsPart.DrawingsPart == null) continue;
                            foreach (var chartPart in anyWsPart.DrawingsPart.ChartParts)
                            {
                                if (chartPart.ChartSpace == null) continue;
                                bool changed = false;
                                foreach (var f in chartPart.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>())
                                {
                                    if (f.Text != null && f.Text.Contains(oldRef, StringComparison.OrdinalIgnoreCase))
                                    {
                                        f.Text = f.Text.Replace(oldRef, newRef, StringComparison.OrdinalIgnoreCase);
                                        changed = true;
                                    }
                                }
                                if (changed) chartPart.ChartSpace.Save();
                            }
                            // CONSISTENCY(sheet-rename-refs): chartEx series
                            // formulas (<cx:f>SheetName!$A$1:$B$2</cx:f>) carry
                            // the same sheet-qualified text as classic <c:f>
                            // and hit the same "external links" failure when
                            // left pointing at the old name.
                            foreach (var extChartPart in anyWsPart.DrawingsPart.ExtendedChartParts)
                            {
                                if (extChartPart.ChartSpace == null) continue;
                                bool changed = false;
                                foreach (var f in extChartPart.ChartSpace.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.Formula>())
                                {
                                    if (f.Text != null && f.Text.Contains(oldRef, StringComparison.OrdinalIgnoreCase))
                                    {
                                        f.Text = f.Text.Replace(oldRef, newRef, StringComparison.OrdinalIgnoreCase);
                                        changed = true;
                                    }
                                }
                                if (changed) extChartPart.ChartSpace.Save();
                            }
                        }

                        // CONSISTENCY(sheet-rename-refs): three more places
                        // carry sheet-qualified formula text in worksheet
                        // XML and need the rename cascaded:
                        //   - sparkline data range  (<xne:f>Sheet1!A1:A4</xne:f>)
                        //   - data validation list  (<x:formula1>Sheet1!A1:A3</x:formula1>)
                        //   - conditional formatting (<x:formula>Sheet1!$A$1</x:formula>)
                        // Walk each worksheet's typed descendants so we
                        // don't accidentally rewrite cell text that happens
                        // to contain the literal substring "Sheet1!".
                        foreach (var anyWsPart2 in workbookPart.WorksheetParts)
                        {
                            var wsRoot = anyWsPart2.Worksheet;
                            if (wsRoot == null) continue;
                            bool wsChanged = false;
                            foreach (var f in wsRoot.Descendants<DocumentFormat.OpenXml.Office.Excel.Formula>())
                            {
                                if (f.Text != null && f.Text.Contains(oldRef, StringComparison.OrdinalIgnoreCase))
                                {
                                    f.Text = f.Text.Replace(oldRef, newRef, StringComparison.OrdinalIgnoreCase);
                                    wsChanged = true;
                                }
                            }
                            // CONSISTENCY(sheet-rename-refs): sparkline location
                            // (<xne:sqref>Sheet1!D1</xne:sqref>) carries the same
                            // sheet-qualified ref text and must follow the rename.
                            // Without this, <xne:f> points at the new sheet but
                            // <xne:sqref> still names the old one — Excel loses
                            // the anchor on render.
                            foreach (var s in wsRoot.Descendants<DocumentFormat.OpenXml.Office.Excel.ReferenceSequence>())
                            {
                                if (s.Text != null && s.Text.Contains(oldRef, StringComparison.OrdinalIgnoreCase))
                                {
                                    s.Text = s.Text.Replace(oldRef, newRef, StringComparison.OrdinalIgnoreCase);
                                    wsChanged = true;
                                }
                            }
                            foreach (var f in wsRoot.Descendants<Formula1>())
                            {
                                if (f.Text != null && f.Text.Contains(oldRef, StringComparison.OrdinalIgnoreCase))
                                {
                                    f.Text = f.Text.Replace(oldRef, newRef, StringComparison.OrdinalIgnoreCase);
                                    wsChanged = true;
                                }
                            }
                            foreach (var f in wsRoot.Descendants<Formula2>())
                            {
                                if (f.Text != null && f.Text.Contains(oldRef, StringComparison.OrdinalIgnoreCase))
                                {
                                    f.Text = f.Text.Replace(oldRef, newRef, StringComparison.OrdinalIgnoreCase);
                                    wsChanged = true;
                                }
                            }
                            foreach (var f in wsRoot.Descendants<Formula>())
                            {
                                if (f.Text != null && f.Text.Contains(oldRef, StringComparison.OrdinalIgnoreCase))
                                {
                                    f.Text = f.Text.Replace(oldRef, newRef, StringComparison.OrdinalIgnoreCase);
                                    wsChanged = true;
                                }
                            }
                            // Internal hyperlinks: <x:hyperlink ref="A1"
                            // location="SheetName!A1"/>. Update the
                            // location attribute when it points at the
                            // renamed sheet.
                            foreach (var hl in wsRoot.Descendants<Hyperlink>())
                            {
                                var loc = hl.Location?.Value;
                                if (loc != null && loc.Contains(oldRef, StringComparison.OrdinalIgnoreCase))
                                {
                                    hl.Location = loc.Replace(oldRef, newRef, StringComparison.OrdinalIgnoreCase);
                                    wsChanged = true;
                                }
                            }
                            if (wsChanged) wsRoot.Save();
                        }

                        workbook.Save();
                        // The formula-trace reverse index is keyed by sheet
                        // NAME (both the formula-cell side and the reference
                        // side) — a rename rewrites exactly that key space.
                        InvalidateDependentsIndex();
                    }
                    break;
                }
                case "freeze":
                {
                    var sheetViews = ws.GetFirstChild<SheetViews>();
                    if (sheetViews == null)
                    {
                        sheetViews = new SheetViews();
                        InsertSheetViewsInSchemaOrder(ws, sheetViews);
                    }
                    var sheetView = sheetViews.GetFirstChild<SheetView>();
                    if (sheetView == null)
                    {
                        sheetView = new SheetView { WorkbookViewId = 0 };
                        sheetViews.AppendChild(sheetView);
                    }

                    if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("false", StringComparison.OrdinalIgnoreCase))
                    {
                        // Remove freeze
                        var existingPane = sheetView.GetFirstChild<Pane>();
                        existingPane?.Remove();
                    }
                    else
                    {
                        // Parse cell reference for freeze position
                        // "A2" = freeze row 1, "B1" = freeze col A, "B2" = freeze row 1 + col A
                        var (col, row) = ParseCellReference(value.ToUpperInvariant());
                        var colSplit = ColumnNameToIndex(col) - 1; // 0-based: B=1 means split at 1
                        var rowSplit = row - 1; // 0-based: 2 means split at 1

                        // Remove existing pane
                        var existingPane = sheetView.GetFirstChild<Pane>();
                        existingPane?.Remove();

                        // R18-B3: freeze=A1 means "no freeze". Emitting a <pane> with
                        // no xSplit/ySplit produces invalid OOXML (Excel repairs on
                        // open). Treat A1 as a no-op after clearing the existing pane.
                        if (colSplit <= 0 && rowSplit <= 0)
                            break;

                        var activePane = (colSplit > 0 && rowSplit > 0) ? PaneValues.BottomRight
                            : (rowSplit > 0) ? PaneValues.BottomLeft
                            : PaneValues.TopRight;

                        var pane = new Pane
                        {
                            TopLeftCell = value.ToUpperInvariant(),
                            State = PaneStateValues.Frozen,
                            ActivePane = activePane
                        };
                        if (rowSplit > 0) pane.VerticalSplit = rowSplit;
                        if (colSplit > 0) pane.HorizontalSplit = colSplit;

                        sheetView.InsertAt(pane, 0);
                    }
                    break;
                }
                case "merge":
                {
                    // Sheet-level merge: value is the range(s) to merge (e.g., "A1:A3" or
                    // "A1:D1,B3:B5" for multiple ranges).
                    // R2-1: Split comma-separated ranges into separate <mergeCell> elements;
                    // Excel rejects a single <mergeCell ref="A1:D1,B3:B5"/>.
                    var refParts = value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                    // CONSISTENCY(merge-empty-container): pre-validate before container
                    // creation — see ExcelHandler.Helpers.ValidateMergeRefLiteral.
                    foreach (var r in refParts) ValidateMergeRefLiteral(r);
                    var mergeCells = ws.GetFirstChild<MergeCells>();
                    if (mergeCells == null)
                    {
                        mergeCells = new MergeCells();
                        ws.AppendChild(mergeCells);
                    }
                    foreach (var part in refParts)
                        InsertMergeCellChecked(mergeCells, part.ToUpperInvariant(), worksheet);
                    mergeCells.Count = (uint)mergeCells.Elements<MergeCell>().Count();
                    break;
                }
                case "autofilter":
                {
                    // Set or remove AutoFilter
                    var existingAf = ws.GetFirstChild<AutoFilter>();
                    var trimmed = (value ?? "").Trim();
                    var lower = trimmed.ToLowerInvariant();
                    if (string.IsNullOrEmpty(trimmed) || lower is "none" or "false" or "0" or "no" or "off")
                    {
                        existingAf?.Remove();
                    }
                    else if (lower is "true" or "1" or "yes" or "on")
                    {
                        // Reject bare bool — autoFilter requires an explicit range. Otherwise
                        // we'd write Reference="TRUE" as raw text and Get would return "TRUE",
                        // which is invalid OOXML and confuses round-trip. Mirrors Add's
                        // "AutoFilter requires 'range' property" rule.
                        throw new ArgumentException(
                            "autoFilter requires an explicit range (e.g. 'A1:F100'). " +
                            "Use 'false'/'none' to remove an existing autoFilter.");
                    }
                    else
                    {
                        // Validate as a real A1 cell/range before writing —
                        // an arbitrary string ("badrange") landed verbatim in
                        // ref=, silently producing a dead filter. Same rule
                        // as printarea/freeze.
                        foreach (var afCell in trimmed.Replace("$", "").Split(':'))
                            ParseCellReference(afCell.Trim());
                        // Canonicalize inverted input (D5:A1) like the rest of
                        // the range family.
                        trimmed = NormalizeA1Range(trimmed);
                        // CONSISTENCY(autofilter-table-dup): a Table already owns
                        // its own <autoFilter>; layering a sheet-level filter over
                        // the same range validates green but real Excel refuses to
                        // open the file (0x800A03EC). Mirror AddAutoFilter's
                        // overlap rejection so Set can't create that state.
                        var afUpper = trimmed.ToUpperInvariant();
                        foreach (var setAfTdp in worksheet.TableDefinitionParts)
                        {
                            if (setAfTdp.Table?.Reference?.Value is string setAfTblRef
                                && RangesOverlap(afUpper, setAfTblRef.ToUpperInvariant()))
                                throw new ArgumentException(
                                    $"AutoFilter range '{afUpper}' overlaps existing table " +
                                    $"'{setAfTdp.Table.Name?.Value ?? setAfTdp.Table.DisplayName?.Value}' " +
                                    $"({setAfTblRef}); tables already include their own autoFilter.");
                        }
                        if (existingAf != null)
                        {
                            existingAf.Reference = trimmed.ToUpperInvariant();
                        }
                        else
                        {
                            var af = new AutoFilter { Reference = trimmed.ToUpperInvariant() };
                            var sheetData = ws.GetFirstChild<SheetData>();
                            if (sheetData != null)
                                sheetData.InsertAfterSelf(af);
                            else
                                ws.AppendChild(af);
                        }
                    }
                    break;
                }
                case "autofit":
                {
                    if (ParseHelpers.IsTruthy(value))
                        AutoFitAllColumns(worksheet);
                    break;
                }
                case "zoom" or "zoomscale":
                {
                    var sheetViews = ws.GetFirstChild<SheetViews>();
                    if (sheetViews == null)
                    {
                        sheetViews = new SheetViews();
                        InsertSheetViewsInSchemaOrder(ws, sheetViews);
                    }
                    var sheetView = sheetViews.GetFirstChild<SheetView>();
                    if (sheetView == null)
                    {
                        sheetView = new SheetView { WorkbookViewId = 0 };
                        sheetViews.AppendChild(sheetView);
                    }
                    var zoomVal = ParseHelpers.SafeParseUint(value, "zoom");
                    if (zoomVal < 10 || zoomVal > 400)
                        throw new ArgumentException($"zoom must be between 10 and 400 (got {zoomVal})");
                    sheetView.ZoomScale = zoomVal;
                    sheetView.ZoomScaleNormal = sheetView.ZoomScale;
                    break;
                }
                case "showgridlines" or "gridlines":
                {
                    var sheetViews = ws.GetFirstChild<SheetViews>();
                    if (sheetViews == null)
                    {
                        sheetViews = new SheetViews();
                        InsertSheetViewsInSchemaOrder(ws, sheetViews);
                    }
                    var sheetView = sheetViews.GetFirstChild<SheetView>();
                    if (sheetView == null)
                    {
                        sheetView = new SheetView { WorkbookViewId = 0 };
                        sheetViews.AppendChild(sheetView);
                    }
                    sheetView.ShowGridLines = ParseHelpers.IsTruthy(value);
                    break;
                }
                case "showrowcolheaders" or "showheaders" or "rowcolheaders" or "headings":
                {
                    var sheetViews = ws.GetFirstChild<SheetViews>();
                    if (sheetViews == null)
                    {
                        sheetViews = new SheetViews();
                        InsertSheetViewsInSchemaOrder(ws, sheetViews);
                    }
                    var sheetView = sheetViews.GetFirstChild<SheetView>();
                    if (sheetView == null)
                    {
                        sheetView = new SheetView { WorkbookViewId = 0 };
                        sheetViews.AppendChild(sheetView);
                    }
                    sheetView.ShowRowColHeaders = ParseHelpers.IsTruthy(value);
                    break;
                }
                case "righttoleft" or "rtl" or "direction" or "sheet.direction":
                {
                    // RTL sheet view (Arabic / Hebrew layouts) — column A renders
                    // on the right, column scroll direction inverts.
                    var sheetViews = ws.GetFirstChild<SheetViews>();
                    if (sheetViews == null)
                    {
                        sheetViews = new SheetViews();
                        InsertSheetViewsInSchemaOrder(ws, sheetViews);
                    }
                    var sheetView = sheetViews.GetFirstChild<SheetView>();
                    if (sheetView == null)
                    {
                        sheetView = new SheetView { WorkbookViewId = 0 };
                        sheetViews.AppendChild(sheetView);
                    }
                    bool rtlOn = key.ToLowerInvariant() switch
                    {
                        "direction" or "sheet.direction" => value.ToLowerInvariant() switch
                        {
                            "rtl" or "righttoleft" or "right-to-left" or "true" or "1" => true,
                            "ltr" or "lefttoright" or "left-to-right" or "false" or "0" or "" => false,
                            _ => throw new ArgumentException($"Invalid direction value: '{value}'. Valid values: rtl, ltr (also accepts true/false, 1/0, righttoleft/lefttoright, right-to-left/left-to-right; case-insensitive).")
                        },
                        _ => ParseHelpers.IsTruthy(value),
                    };
                    // CONSISTENCY(canonical): on default-LTR (Excel sheets have
                    // no inheritance source above them), explicit ltr clears the
                    // attribute rather than writing rightToLeft="0". Mirrors
                    // Word `direction=ltr` clear semantics on default-LTR
                    // contexts. Get already only emits direction=rtl, so this
                    // restores Add/Set/Get symmetry.
                    if (rtlOn) sheetView.RightToLeft = true;
                    else sheetView.RightToLeft = null;
                    break;
                }

                case "tabcolor" or "tab_color":
                {
                    var sheetPr = ws.GetFirstChild<SheetProperties>();
                    if (sheetPr == null)
                    {
                        sheetPr = new SheetProperties();
                        ws.InsertAt(sheetPr, 0);
                    }
                    sheetPr.RemoveAllChildren<TabColor>();
                    if (!value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        // CONSISTENCY(scheme-color): accept scheme-color names
                        // ("accent1"-"accent6", "lt1", "dk1", ...) by mapping
                        // them to TabColor.Theme index. Otherwise fall back to
                        // the numeric color parser for hex/named/rgb() inputs.
                        // A `+tintNN` suffix rides along for parity with cell
                        // fill / font.color ("accent1+tint40").
                        var (tabBase, tabTint) = OfficeCli.Core.ParseHelpers.SplitExcelColorTint(value);
                        var themeIndex = ExcelSchemeColorNameToThemeIndex(tabBase);
                        TabColor tabColor;
                        if (themeIndex.HasValue)
                        {
                            tabColor = new TabColor { Theme = (UInt32Value)themeIndex.Value };
                        }
                        else
                        {
                            var colorHex = OfficeCli.Core.ParseHelpers.NormalizeArgbColor(tabBase);
                            tabColor = new TabColor { Rgb = new HexBinaryValue(colorHex) };
                        }
                        if (tabTint is { } tt
                            && Math.Abs(tt) >= OfficeCli.Core.ParseHelpers.ExcelTintEpsilon)
                            tabColor.Tint = tt;
                        sheetPr.AppendChild(tabColor);
                    }
                    break;
                }

                case "hidden":
                case "visibility":
                {
                    // Sheet visibility lives on the workbook-level <sheet> element,
                    // not on the worksheet. Three-state: visible / hidden / veryHidden.
                    var wbSheets = GetWorkbook().GetFirstChild<Sheets>();
                    var wbSheet = wbSheets?.Elements<Sheet>()
                        .FirstOrDefault(s => s.Name?.Value?.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true);
                    if (wbSheet != null)
                    {
                        var v = (value ?? "").Trim();
                        var keyLower = key.ToLowerInvariant();
                        if (v.Equals("veryHidden", StringComparison.OrdinalIgnoreCase)
                            || v.Equals("very", StringComparison.OrdinalIgnoreCase)
                            || v.Equals("veryhidden", StringComparison.OrdinalIgnoreCase))
                        {
                            wbSheet.State = SheetStateValues.VeryHidden;
                        }
                        else if (v.Equals("hidden", StringComparison.OrdinalIgnoreCase)
                            || (keyLower == "hidden" && ParseHelpers.IsTruthy(v)))
                        {
                            wbSheet.State = SheetStateValues.Hidden;
                        }
                        else if (v.Equals("visible", StringComparison.OrdinalIgnoreCase)
                            || (keyLower == "hidden" && !ParseHelpers.IsTruthy(v))
                            || (keyLower == "visibility" && (string.IsNullOrEmpty(v) || v.Equals("none", StringComparison.OrdinalIgnoreCase))))
                        {
                            wbSheet.State = null;
                        }
                        else
                        {
                            // Unknown value — fall back to truthiness on hidden semantics
                            wbSheet.State = ParseHelpers.IsTruthy(v) ? SheetStateValues.Hidden : null;
                        }
                        // Excel requires at least one visible sheet; a workbook
                        // with none passes schema validation but real Excel
                        // refuses the file (0x800A03EC). Reject the transition
                        // that would hide the last visible sheet.
                        static bool IsHiddenState(Sheet s) =>
                            s.State?.Value == SheetStateValues.Hidden || s.State?.Value == SheetStateValues.VeryHidden;
                        if (IsHiddenState(wbSheet) && wbSheets!.Elements<Sheet>().All(IsHiddenState))
                        {
                            wbSheet.State = null;
                            throw new ArgumentException(
                                $"Cannot hide sheet '{sheetName}': a workbook must keep at least one visible sheet, " +
                                "or Excel refuses to open the file. Unhide another sheet first.");
                        }
                        GetWorkbook().Save();
                    }
                    break;
                }

                // ==================== Sheet Protection ====================
                case "protect":
                {
                    var existingSp = ws.GetFirstChild<SheetProtection>();
                    if (ParseHelpers.IsTruthy(value))
                    {
                        if (existingSp == null)
                        {
                            existingSp = new SheetProtection();
                            InsertSheetProtectionInOrder(ws, existingSp);
                        }
                        existingSp.Sheet = true;
                        existingSp.Objects = true;
                        existingSp.Scenarios = true;
                    }
                    else
                    {
                        existingSp?.Remove();
                    }
                    break;
                }
                case "password":
                {
                    var sp = ws.GetFirstChild<SheetProtection>();
                    if (sp == null)
                    {
                        sp = new SheetProtection { Sheet = true, Objects = true, Scenarios = true };
                        InsertSheetProtectionInOrder(ws, sp);
                    }
                    if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        sp.Password = null;
                    else
                    {
                        // Excel legacy password hash (ECMA-376 Part 4, 14.7.1)
                        int hash = 0;
                        for (int ci = value.Length - 1; ci >= 0; ci--)
                        {
                            hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7FFF);
                            hash ^= value[ci];
                        }
                        hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7FFF);
                        hash ^= value.Length;
                        hash ^= 0xCE4B;
                        sp.Password = HexBinaryValue.FromString(hash.ToString("X4"));
                    }
                    break;
                }
                // Verbatim hash write-back — the dump emitter surfaces the
                // stored hashes (legacy `password` attr / modern algorithm
                // set) so replay preserves protection strength. Values are
                // written as-is, never re-hashed.
                case "passwordhash":
                {
                    var spH = ws.GetFirstChild<SheetProtection>();
                    if (spH == null)
                    {
                        spH = new SheetProtection { Sheet = true, Objects = true, Scenarios = true };
                        InsertSheetProtectionInOrder(ws, spH);
                    }
                    spH.Password = string.IsNullOrEmpty(value) ? null : HexBinaryValue.FromString(value);
                    break;
                }
                case "protection.algorithm":
                case "protection.hash":
                case "protection.salt":
                case "protection.spincount":
                {
                    var spM = ws.GetFirstChild<SheetProtection>();
                    if (spM == null)
                    {
                        spM = new SheetProtection { Sheet = true, Objects = true, Scenarios = true };
                        InsertSheetProtectionInOrder(ws, spM);
                    }
                    switch (key.ToLowerInvariant())
                    {
                        case "protection.algorithm": spM.AlgorithmName = value; break;
                        case "protection.hash": spM.HashValue = value; break;
                        case "protection.salt": spM.SaltValue = value; break;
                        case "protection.spincount":
                            spM.SpinCount = uint.TryParse(value, out var spin) ? spin : null;
                            break;
                    }
                    break;
                }

                // ==================== Print Settings ====================
                case "printarea":
                {
                    var workbook = GetWorkbook();
                    // CONSISTENCY(workbook-child-order): use helper to create
                    // <definedNames> in schema-correct position when missing.
                    var definedNames = GetOrCreateDefinedNames(workbook);
                    // Find sheet index
                    var allSheets = workbook.GetFirstChild<Sheets>()?.Elements<Sheet>().ToList();
                    var sheetIdx = allSheets?.FindIndex(s =>
                        s.Name?.Value?.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true) ?? -1;
                    // Remove existing print area for this sheet
                    var existing = definedNames.Elements<DefinedName>()
                        .Where(d => d.Name == "_xlnm.Print_Area" && d.LocalSheetId?.Value == (uint)sheetIdx)
                        .ToList();
                    foreach (var e in existing) e.Remove();

                    if (!string.IsNullOrEmpty(value) && !value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        // A user-supplied "OtherSheet!A1:C10" prefix used to be
                        // blindly prepended with this sheet's name, producing a
                        // double-qualified defined name ("Sheet1!Bogus!A1:C10")
                        // that passes schema validation but makes real Excel
                        // refuse the file (0x800A03EC). The print area always
                        // targets the sheet being set — strip a prefix naming
                        // this sheet, reject any other.
                        var paRange = value;
                        var bangIdx = paRange.LastIndexOf('!');
                        if (bangIdx >= 0)
                        {
                            var paSheet = paRange[..bangIdx].Trim('\'');
                            if (!paSheet.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                                throw new ArgumentException(
                                    $"printArea '{value}' names sheet '{paSheet}', but this set targets '{sheetName}'. A print area always applies to its own sheet — pass just the range (e.g. printArea=A1:C10), or run set on /{paSheet} if that sheet exists.");
                            paRange = paRange[(bangIdx + 1)..];
                        }
                        // Bounds-check the range like freeze/merge/comment do:
                        // an out-of-grid ref (XFE, row 1048577) is invisible
                        // to schema validation (defined names are opaque
                        // strings) but makes real Excel fail to render.
                        foreach (var paCell in paRange.Replace("$", "").Split(':'))
                            ParseCellReference(paCell.Trim());
                        // Canonicalize an inverted range (D10:A1) — stored
                        // verbatim it passes validation and round-trips, but
                        // real Excel fails to render the sheet. Same per-axis
                        // swap as anchor/table/sqref.
                        var paParts = paRange.Replace("$", "").Split(':');
                        if (paParts.Length == 2)
                        {
                            var (paC1, paR1) = ParseCellReference(paParts[0].Trim());
                            var (paC2, paR2) = ParseCellReference(paParts[1].Trim());
                            int paI1 = ColumnNameToIndex(paC1), paI2 = ColumnNameToIndex(paC2);
                            if (paI2 < paI1 || paR2 < paR1)
                            {
                                // Only rewrite when actually inverted, so a
                                // well-formed input keeps its $ anchors.
                                if (paI2 < paI1) (paC1, paC2) = (paC2, paC1);
                                if (paR2 < paR1) (paR1, paR2) = (paR2, paR1);
                                paRange = $"{paC1}{paR1}:{paC2}{paR2}";
                            }
                        }
                        var dn = new DefinedName($"{Core.ModernFunctionQualifier.QuoteSheetNameForRef(sheetName)}!{paRange}") { Name = "_xlnm.Print_Area" };
                        if (sheetIdx >= 0) dn.LocalSheetId = (uint)sheetIdx;
                        definedNames.AppendChild(dn);
                    }
                    workbook.Save();
                    break;
                }
                case "printtitlerows" or "printtitlerow":
                case "printtitlecols" or "printtitlecol" or "printtitlecolumns":
                {
                    // Print_Titles definedName: combines repeating rows and
                    // repeating columns into a single comma-separated value
                    // for the sheet, e.g. "Sheet1!$A:$A,Sheet1!$1:$1".
                    var workbook = GetWorkbook();
                    // CONSISTENCY(workbook-child-order): use helper to create
                    // <definedNames> in schema-correct position when missing.
                    var definedNames = GetOrCreateDefinedNames(workbook);
                    var allSheets = workbook.GetFirstChild<Sheets>()?.Elements<Sheet>().ToList();
                    var sheetIdx = allSheets?.FindIndex(s =>
                        s.Name?.Value?.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true) ?? -1;
                    if (sheetIdx < 0)
                        throw new ArgumentException($"Sheet '{sheetName}' not found in workbook.");

                    // OrdinalIgnoreCase: the switch matched on the lowercased
                    // key but `key` keeps the caller's casing — the camelCase
                    // spelling printTitleRows silently routed to the cols branch.
                    bool isRows = key.StartsWith("printtitlerow", StringComparison.OrdinalIgnoreCase);

                    // Validate BEFORE the existing Print_Titles entry is
                    // removed below — a rejected value must not destroy the
                    // previously-set titles (failed-set atomicity).
                    if (!string.IsNullOrEmpty(value) && !value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var checkParts = value.Trim().Replace("$", "");
                        var checkBang = checkParts.IndexOf('!');
                        if (checkBang >= 0) checkParts = checkParts[(checkBang + 1)..];
                        foreach (var tok in checkParts.Split(':'))
                        {
                            var t = tok.Trim();
                            if (t.Length == 0) continue;
                            if (isRows)
                            {
                                if (!uint.TryParse(t, out var rowNum) || rowNum < 1 || rowNum > 1048576)
                                    throw new ArgumentException(
                                        $"Invalid print title row '{t}': rows must be 1..1048576 (e.g. printTitleRows=1:2).");
                            }
                            else
                            {
                                if (!t.All(char.IsLetter) || ColumnNameToIndex(t.ToUpperInvariant()) > 16384)
                                    throw new ArgumentException(
                                        $"Invalid print title column '{t}': columns must be A..XFD (e.g. printTitleCols=A:B).");
                            }
                        }
                    }

                    // Read existing Print_Titles for this sheet, parse row/col parts.
                    var existingDn = definedNames.Elements<DefinedName>()
                        .FirstOrDefault(d => d.Name == "_xlnm.Print_Titles" && d.LocalSheetId?.Value == (uint)sheetIdx);
                    string? rowsPart = null;
                    string? colsPart = null;
                    if (existingDn != null)
                    {
                        var raw = existingDn.Text ?? "";
                        foreach (var tok in raw.Split(',', StringSplitOptions.RemoveEmptyEntries))
                        {
                            var t = tok.Trim();
                            // Strip leading "SheetName!" if present
                            var bang = t.IndexOf('!');
                            var rangePart = bang >= 0 ? t[(bang + 1)..] : t;
                            // Row range looks like $1:$5 (digits only); col range like $A:$C (letters only)
                            var inner = rangePart.Replace("$", "");
                            var leftSide = inner.Split(':')[0];
                            if (leftSide.Length > 0 && char.IsDigit(leftSide[0]))
                                rowsPart = t;
                            else if (leftSide.Length > 0 && char.IsLetter(leftSide[0]))
                                colsPart = t;
                        }
                        existingDn.Remove();
                    }

                    static string Normalize(string sheet, string range, bool rows)
                    {
                        var v = range.Trim();
                        // Allow shorthand "1:1" or "A:A" (no $); add $ to columns/rows.
                        if (!v.Contains('$'))
                        {
                            var parts = v.Split(':');
                            if (parts.Length == 2)
                                v = rows ? $"${parts[0]}:${parts[1]}" : $"${parts[0]}:${parts[1]}";
                        }
                        // Allow user to pass already-qualified "Sheet1!$1:$1"; otherwise prefix.
                        // Quote the sheet name when Excel requires it (digit-start,
                        // space, hyphen, …) — an unquoted name corrupts the workbook.
                        return v.Contains('!') ? v : $"{Core.ModernFunctionQualifier.QuoteSheetNameForRef(sheet)}!{v}";
                    }

                    if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        if (isRows) rowsPart = null; else colsPart = null;
                    }
                    else
                    {
                        var normalized = Normalize(sheetName, value, isRows);
                        if (isRows) rowsPart = normalized; else colsPart = normalized;
                    }

                    var combined = string.Join(",", new[] { colsPart, rowsPart }.Where(s => !string.IsNullOrEmpty(s)));
                    if (!string.IsNullOrEmpty(combined))
                    {
                        var dn = new DefinedName(combined) { Name = "_xlnm.Print_Titles", LocalSheetId = (uint)sheetIdx };
                        definedNames.AppendChild(dn);
                    }
                    workbook.Save();
                    break;
                }
                case "orientation" or "pageorientation":
                {
                    var pageSetup = ws.GetFirstChild<PageSetup>();
                    if (pageSetup == null)
                    {
                        pageSetup = new PageSetup();
                        ws.AppendChild(pageSetup);
                    }
                    pageSetup.Orientation = value.ToLowerInvariant() == "landscape"
                        ? OrientationValues.Landscape
                        : OrientationValues.Portrait;
                    break;
                }
                case "papersize":
                {
                    var pageSetup = ws.GetFirstChild<PageSetup>();
                    if (pageSetup == null)
                    {
                        pageSetup = new PageSetup();
                        ws.AppendChild(pageSetup);
                    }
                    pageSetup.PaperSize = ParseHelpers.SafeParseUint(value, "paperSize");
                    break;
                }
                case "fittopage":
                {
                    // Treat "false"/"none"/"0" as a clear: drop FitToPage flag and any
                    // FitToWidth/FitToHeight overrides so readback no longer reports
                    // a fittopage value.
                    var fitParts = value.Split('x', 'X');
                    uint fw = 0, fh = 0;
                    bool isWxH = fitParts.Length == 2
                        && uint.TryParse(fitParts[0], out fw)
                        && uint.TryParse(fitParts[1], out fh);
                    // A value that LOOKS like the WxH form but fails to parse
                    // must not fall through to the boolean parser — the
                    // resulting "Invalid boolean value" message pointed users
                    // at true/false instead of the actual format.
                    if (!isWxH && fitParts.Length == 2)
                        throw new ArgumentException(
                            $"Invalid fitToPage value '{value}'. Expected WIDTHxHEIGHT with non-negative integers (e.g. 1x1, 2x0), or false/none to clear.");
                    // ECMA-376 caps fitToWidth/fitToHeight at 32767; larger
                    // values save fine and pass most validation but real Excel
                    // refuses the file (0x800A03EC). Mirror the margin guard.
                    if (isWxH && (fw > 32767 || fh > 32767))
                        throw new ArgumentException(
                            $"fitToPage value '{value}' is out of range: width and height must each be 0-32767 pages.");
                    bool clearing = !isWxH
                        && (string.IsNullOrEmpty(value)
                            || value.Equals("none", StringComparison.OrdinalIgnoreCase)
                            || !ParseHelpers.IsTruthy(value));

                    if (clearing)
                    {
                        var spExisting = ws.GetFirstChild<SheetProperties>();
                        var pspExisting = spExisting?.GetFirstChild<PageSetupProperties>();
                        if (pspExisting != null)
                        {
                            pspExisting.FitToPage = null;
                            // Drop the wrapper if it has no other attributes/children
                            if (!pspExisting.GetAttributes().Any() && !pspExisting.HasChildren)
                                pspExisting.Remove();
                        }
                        var psExisting = ws.GetFirstChild<PageSetup>();
                        if (psExisting != null)
                        {
                            psExisting.FitToWidth = null;
                            psExisting.FitToHeight = null;
                        }
                        break;
                    }

                    var sheetPr = ws.GetFirstChild<SheetProperties>();
                    if (sheetPr == null)
                    {
                        sheetPr = new SheetProperties();
                        ws.InsertAt(sheetPr, 0);
                    }
                    var psp = sheetPr.GetFirstChild<PageSetupProperties>();
                    if (psp == null)
                    {
                        psp = new PageSetupProperties();
                        sheetPr.AppendChild(psp);
                    }
                    psp.FitToPage = true;

                    var pageSetup = ws.GetFirstChild<PageSetup>();
                    if (pageSetup == null)
                    {
                        pageSetup = new PageSetup();
                        ws.AppendChild(pageSetup);
                    }
                    if (isWxH)
                    {
                        pageSetup.FitToWidth = fw;
                        pageSetup.FitToHeight = fh;
                    }
                    else
                    {
                        pageSetup.FitToWidth = 1;
                        pageSetup.FitToHeight = 1;
                    }
                    break;
                }
                case "header":
                {
                    // Reject XML-illegal control chars up front — otherwise the
                    // value is accepted and only fails at save ("data may be
                    // lost"). Same guard every other text property gets.
                    Core.ParseHelpers.ValidateXmlText(value, "header");
                    var hf = ws.GetFirstChild<HeaderFooter>();
                    if (hf == null)
                    {
                        hf = new HeaderFooter();
                        ws.AppendChild(hf);
                    }
                    hf.OddHeader = new OddHeader(value);
                    break;
                }
                case "footer":
                {
                    Core.ParseHelpers.ValidateXmlText(value, "footer");
                    var hf = ws.GetFirstChild<HeaderFooter>();
                    if (hf == null)
                    {
                        hf = new HeaderFooter();
                        ws.AppendChild(hf);
                    }
                    hf.OddFooter = new OddFooter(value);
                    break;
                }
                case "margin.top" or "margin.bottom" or "margin.left" or "margin.right" or "margin.header" or "margin.footer":
                {
                    var inches = ParseMarginInches(value);
                    // Negative margins pass schema validation (only an upper
                    // bound is declared) but real Excel refuses the file
                    // (0x800A03EC). Reject up front; 49in is the schema cap.
                    if (inches < 0 || inches >= 49)
                        throw new ArgumentException(
                            $"Invalid '{key}' value '{value}': margins must be between 0 and 49 inches.");
                    var pm = ws.GetFirstChild<PageMargins>();
                    if (pm == null)
                    {
                        // PageMargins requires all 6 attributes; default per Excel.
                        pm = new PageMargins
                        {
                            Top = 0.75, Bottom = 0.75,
                            Left = 0.7, Right = 0.7,
                            Header = 0.3, Footer = 0.3
                        };
                        // PageMargins must precede pageSetup, headerFooter, etc. but follow
                        // sheetProtection/printOptions. Insert before pageSetup if present.
                        var anchor = ws.GetFirstChild<PageSetup>() ?? (OpenXmlElement?)ws.GetFirstChild<HeaderFooter>();
                        if (anchor != null) ws.InsertBefore(pm, anchor);
                        else ws.AppendChild(pm);
                    }
                    var which = key.ToLowerInvariant().Substring("margin.".Length);
                    switch (which)
                    {
                        case "top": pm.Top = inches; break;
                        case "bottom": pm.Bottom = inches; break;
                        case "left": pm.Left = inches; break;
                        case "right": pm.Right = inches; break;
                        case "header": pm.Header = inches; break;
                        case "footer": pm.Footer = inches; break;
                    }
                    break;
                }

                // ==================== Sorting ====================
                // CONSISTENCY(range-action): sort is a region action like merge.
                // Sheet-level path auto-detects the full used range; explicit ranges
                // go through SetRange → SortRangeRows. Keep both entry points in
                // sync. See the project conventions "Consistency > Robustness".
                case "sort":
                {
                    // R7-3: remove ALL sortState children (malformed files may
                    // carry more than one; GetFirstChild leaves stragglers).
                    foreach (var __ss in ws.Descendants<SortState>().ToList()) __ss.Remove();
                    if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        break;

                    var sd = ws.GetFirstChild<SheetData>();
                    if (sd == null) sd = ws.AppendChild(new SheetData());
                    var rows = sd.Elements<Row>().ToList();
                    // R12-2: DO NOT early-return on empty sheet here. Empty sheet + invalid
                    // sort spec (e.g. "XFE asc", "AAAA asc", "sort=asc") used to silently
                    // succeed because we bailed before spec validation. Always dispatch into
                    // SortRangeRows so it validates the spec first; if spec is valid and there
                    // is no data, it no-ops cleanly via its existing dataStartRow > row2 guard.
                    int maxCol = 1;
                    foreach (var r in rows)
                        foreach (var c in r.Elements<Cell>())
                        {
                            var cref = c.CellReference?.Value;
                            if (cref == null) continue;
                            maxCol = Math.Max(maxCol, ColumnNameToIndex(ParseCellReference(cref).Column));
                        }
                    int minRowIdx = rows.Count == 0 ? 1 : (int)rows.Min(r => r.RowIndex?.Value ?? 1u);
                    int maxRowIdx = rows.Count == 0 ? 1 : (int)rows.Max(r => r.RowIndex?.Value ?? 1u);

                    // CONSISTENCY(sort-header-default): sortHeader defaults to false
                    // (row 1 participates in the reorder). This matches our general
                    // "caller states intent explicitly" rule and is documented in help.
                    // R4-D1 and R7-4 both proposed auto-detecting headers (type-mismatch
                    // heuristic, first-row-is-string warning). Rejected: heuristic
                    // warnings ship false positives on legitimately-heterogeneous
                    // row-1 data and are spammy in pipelines. Future revisit: make
                    // sortHeader default=true project-wide as a breaking change,
                    // documented in release notes — do NOT add a per-call warning.
                    // CONSISTENCY(sort-header-autofilter): when an AutoFilter is set on the
                    // sheet, row 1 is unambiguously the header (the filter's own header row),
                    // so default sortHeader=true to keep it out of the reorder — sorting it
                    // into the data silently destroys the dataset. Explicit sortHeader=false
                    // still overrides. Without an AutoFilter we keep the documented
                    // "caller states intent" default of false (no heuristic guessing).
                    bool sortHeader;
                    if (properties.TryGetValue("sortheader", out var shv))
                        sortHeader = IsTruthy(shv);
                    else
                        sortHeader = ws.GetFirstChild<AutoFilter>() != null;
                    SortRangeRows(worksheet, 1, minRowIdx, maxCol, maxRowIdx, value, sortHeader);
                    DeleteCalcChainIfPresent();
                    break;
                }
                case "sortheader":
                    // consumed by "sort" case above; ignore silently here so it doesn't show unsupported
                    break;

                default:
                    unsupported.Add(unsupported.Count == 0
                        ? $"{key} (valid sheet props: name, freeze, zoom, showGridLines, showRowColHeaders, tabcolor, autofilter, visibility, hidden, merge, protect, password, printarea, printTitleRows, printTitleCols, orientation, papersize, fittopage, header, footer, margin.top, margin.bottom, margin.left, margin.right, margin.header, margin.footer, sort, sortHeader)"
                        : key);
                    break;
            }
        }

        // An explicit protect=false wins over a password= in the SAME call.
        // The password case (re)creates a SheetProtection with sheet=true to
        // hold the hash, which would otherwise resurrect protection the caller
        // just asked to turn off (order-dependently) — e.g.
        // `--prop protect=false --prop password=none` left the sheet protected.
        // Re-assert the removal here so the explicit flag is authoritative.
        if (properties.TryGetValue("protect", out var protectFinal)
            && !ParseHelpers.IsTruthy(protectFinal))
        {
            ws.GetFirstChild<SheetProtection>()?.Remove();
        }

        SaveWorksheet(worksheet);
        return unsupported;
    }

}
