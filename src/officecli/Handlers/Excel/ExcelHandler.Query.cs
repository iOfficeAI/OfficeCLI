// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ==================== Query Layer ====================

    public DocumentNode Get(string path, int depth = 1)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("Path cannot be empty.");
        path = NormalizeExcelPath(path);
        path = ResolveSheetIndexInPath(path);
        if (path == "/")
        {
            var node = new DocumentNode { Path = "/", Type = "workbook" };

            // Core document properties
            var props = _doc.PackageProperties;
            if (props.Title != null) node.Format["title"] = props.Title;
            if (props.Creator != null) node.Format["author"] = props.Creator;
            if (props.Subject != null) node.Format["subject"] = props.Subject;
            if (props.Keywords != null) node.Format["keywords"] = props.Keywords;
            if (props.Description != null) node.Format["description"] = props.Description;
            if (props.Category != null) node.Format["category"] = props.Category;
            if (props.LastModifiedBy != null) node.Format["lastModifiedBy"] = props.LastModifiedBy;
            if (props.Revision != null) node.Format["revisionNumber"] = props.Revision;
            if (props.Created != null) node.Format["created"] = props.Created.Value.ToString("o");
            if (props.Modified != null) node.Format["modified"] = props.Modified.Value.ToString("o");

            foreach (var (name, part) in GetWorksheets())
            {
                var sheetNode = new DocumentNode { Path = $"/{name}", Type = "sheet", Preview = name };
                var sheetData = GetSheet(part).GetFirstChild<SheetData>();
                // R6-5: dedupe by RowIndex so a pivot placed on its own source
                // sheet doesn't double-count row children.
                var rowCount = sheetData?.Elements<Row>()
                    .Select(r => r.RowIndex?.Value ?? 0u)
                    .Where(i => i != 0)
                    .Distinct()
                    .Count() ?? 0;
                var chartCount = part.DrawingsPart != null ? CountExcelCharts(part.DrawingsPart) : 0;
                sheetNode.ChildCount = rowCount + chartCount;

                if (depth > 0 && sheetData != null)
                {
                    sheetNode.Children = GetSheetChildNodes(name, sheetData, depth, part);
                    // Children omit value-less empty cells/rows (issue #149);
                    // reflect the actual listed count, not the raw row count.
                    sheetNode.ChildCount = sheetNode.Children.Count;
                }

                node.Children.Add(sheetNode);
            }
            // Workbook-level settings
            PopulateWorkbookSettings(node);
            Core.ThemeHandler.PopulateTheme(_doc.WorkbookPart?.ThemePart, node);
            Core.ExtendedPropertiesHandler.PopulateExtendedProperties(_doc.ExtendedFilePropertiesPart, node);

            node.ChildCount = node.Children.Count;
            return node;
        }

        // Bare /namedrange (no index): list all defined names as children
        // instead of falling through to the sheet-name lookup, where the
        // literal word "namedrange" was treated as a missing sheet and
        // produced a confusing "sheet not found" error.
        var bareNamedRange = Regex.Match(path.TrimStart('/'), @"^namedrange$", RegexOptions.IgnoreCase);
        if (bareNamedRange.Success)
        {
            var listNode = new DocumentNode { Path = "/namedrange", Type = "namedrange_list" };
            var workbook = GetWorkbook();
            var allDefs = workbook.GetFirstChild<DefinedNames>()?.Elements<DefinedName>().ToList()
                ?? new List<DefinedName>();
            for (int i = 0; i < allDefs.Count; i++)
            {
                var dn = allDefs[i];
                var nrNode = new DocumentNode
                {
                    Path = $"/namedrange[{i + 1}]",
                    Type = "namedrange",
                    Text = dn.InnerText ?? dn.Name?.Value ?? "",
                    Preview = dn.InnerText
                };
                nrNode.Format["name"] = dn.Name?.Value ?? "";
                nrNode.Format["ref"] = dn.InnerText ?? "";
                if (dn.LocalSheetId?.HasValue == true)
                {
                    var sheets = workbook.GetFirstChild<Sheets>()?.Elements<Sheet>().ToList();
                    if (sheets != null && (int)dn.LocalSheetId.Value < sheets.Count)
                        nrNode.Format["scope"] = sheets[(int)dn.LocalSheetId.Value].Name?.Value ?? "";
                }
                else
                {
                    nrNode.Format["scope"] = "workbook";
                }
                if (!string.IsNullOrEmpty(dn.Comment?.Value))
                    nrNode.Format["comment"] = dn.Comment.Value;
                if (dn.Function?.Value == true)
                    nrNode.Format["volatile"] = true;
                listNode.Children.Add(nrNode);
            }
            listNode.ChildCount = listNode.Children.Count;
            return listNode;
        }

        // Bare /sheet (no index): list all sheets as children.
        // CONSISTENCY(bare-path-lister): mirrors /namedrange so agents can
        // discover top-level collections without first knowing names.
        var bareSheet = Regex.Match(path.TrimStart('/'), @"^sheet$", RegexOptions.IgnoreCase);
        if (bareSheet.Success)
        {
            var listNode = new DocumentNode { Path = "/sheet", Type = "sheet_list" };
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                var sheetNode = new DocumentNode { Path = $"/{sheetName}", Type = "sheet", Preview = sheetName };
                var sd = GetSheet(worksheetPart).GetFirstChild<SheetData>();
                var rowCount = sd?.Elements<Row>().Count() ?? 0;
                var chartCount = worksheetPart.DrawingsPart != null ? CountExcelCharts(worksheetPart.DrawingsPart) : 0;
                sheetNode.ChildCount = rowCount + chartCount;
                sheetNode.Format["name"] = sheetName;
                listNode.Children.Add(sheetNode);
            }
            listNode.ChildCount = listNode.Children.Count;
            return listNode;
        }

        // Bare /table (no index): list every table across every sheet.
        // CONSISTENCY(bare-path-lister): mirrors /namedrange.
        var bareTable = Regex.Match(path.TrimStart('/'), @"^table$", RegexOptions.IgnoreCase);
        if (bareTable.Success)
        {
            var listNode = new DocumentNode { Path = "/table", Type = "table_list" };
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                var tParts = worksheetPart.TableDefinitionParts.ToList();
                for (int i = 0; i < tParts.Count; i++)
                    listNode.Children.Add(TableToNode(sheetName, worksheetPart, i + 1, 0));
            }
            listNode.ChildCount = listNode.Children.Count;
            return listNode;
        }

        // Handle /namedrange[N] or /namedrange[Name] or /namedrange[@name=X]
        var namedRangeMatch = Regex.Match(path.TrimStart('/'), @"^namedrange\[(.+?)\]$", RegexOptions.IgnoreCase);
        if (namedRangeMatch.Success)
        {
            var selector = namedRangeMatch.Groups[1].Value;
            // BUG-R36-B4: accept attribute-style selector /namedrange[@name=X]
            // for parity with /formfield[@name=X]; previously the literal
            // "@name=X" string was treated as the defined-name to match,
            // matched nothing, and returning null! crashed downstream.
            var attrMatch = Regex.Match(selector, @"^@name=(.+)$", RegexOptions.IgnoreCase);
            if (attrMatch.Success)
                selector = attrMatch.Groups[1].Value.Trim('"', '\'');
            var workbook = GetWorkbook();
            var definedNames = workbook.GetFirstChild<DefinedNames>();
            // BUG-R36-B4: previously returned null! on miss, which the resident
            // caller dereferenced (NullReferenceException). Return a typed error
            // node so the standard "not found -> ArgumentException" path fires.
            if (definedNames == null)
                return new DocumentNode { Path = path, Type = "error", Text = $"Named range '{selector}' not found (no defined names in workbook)" };

            var allDefs = definedNames.Elements<DefinedName>().ToList();
            DefinedName? dn = null;
            int dnIndex;

            if (int.TryParse(selector, out dnIndex))
            {
                if (dnIndex < 1 || dnIndex > allDefs.Count)
                    return new DocumentNode { Path = path, Type = "error", Text = $"Named range {dnIndex} not found (total: {allDefs.Count})" };
                dn = allDefs[dnIndex - 1];
            }
            else
            {
                dn = allDefs.FirstOrDefault(d =>
                    d.Name?.Value?.Equals(selector, StringComparison.OrdinalIgnoreCase) == true);
                if (dn == null)
                    return new DocumentNode { Path = path, Type = "error", Text = $"Named range '{selector}' not found" };
                dnIndex = PathIndex.FromArrayIndex(allDefs.IndexOf(dn));
            }

            var nrNode = new DocumentNode
            {
                Path = $"/namedrange[{dnIndex}]",
                Type = "namedrange",
                Text = dn.InnerText ?? dn.Name?.Value ?? "",
                Preview = dn.InnerText
            };
            nrNode.Format["name"] = dn.Name?.Value ?? "";
            nrNode.Format["ref"] = dn.InnerText ?? "";
            if (dn.LocalSheetId?.HasValue == true)
            {
                var sheets = workbook.GetFirstChild<Sheets>()?.Elements<Sheet>().ToList();
                if (sheets != null && (int)dn.LocalSheetId.Value < sheets.Count)
                    nrNode.Format["scope"] = sheets[(int)dn.LocalSheetId.Value].Name?.Value ?? "";
            }
            else
            {
                // Schema declares scope get=true; emit "workbook" for workbook-scope names.
                nrNode.Format["scope"] = "workbook";
            }
            if (!string.IsNullOrEmpty(dn.Comment?.Value))
                nrNode.Format["comment"] = dn.Comment.Value;
            if (dn.Function?.Value == true)
                nrNode.Format["volatile"] = true;

            return nrNode;
        }

        // Parse path: /SheetName or /SheetName/A1 or /SheetName/A1:D10
        var segments = path.TrimStart('/').Split('/', 2);
        var sheetNameFromPath = segments[0];
        // workbook is a singleton at the document root — reject an indexed
        // /workbook[N] with a redirect rather than treating "workbook[N]" as a
        // sheet name (which fires a misleading SheetNotFoundException). Mirrors
        // the pptx notes[N]/theme[N] and docx watermark[N] redirects.
        if (Regex.IsMatch(sheetNameFromPath, @"^workbook\[\d+\]$", RegexOptions.IgnoreCase))
            throw new ArgumentException("workbook is a singleton; use /workbook or / (no index).");
        // docProps is a document-level part, not a sheet — same redirect class.
        if (Regex.IsMatch(sheetNameFromPath, @"^docProps\[\d+\]$", RegexOptions.IgnoreCase))
            throw new ArgumentException("docProps is a singleton; use /docProps or / (no index).");
        var worksheet = FindWorksheet(sheetNameFromPath);
        if (worksheet == null)
            throw SheetNotFoundException(sheetNameFromPath);
        // CONSISTENCY(path-stability): if the path used sheet[N] / sheet[last()],
        // rebuild the canonical path with the resolved sheet name so the returned
        // node.Path reflects the actual sheet (matches Word's last() echo behavior).
        var resolvedSheetName = ResolveSheetName(sheetNameFromPath);
        if (!resolvedSheetName.Equals(sheetNameFromPath, StringComparison.Ordinal))
        {
            sheetNameFromPath = resolvedSheetName;
            path = segments.Length == 1 ? $"/{resolvedSheetName}" : $"/{resolvedSheetName}/{segments[1]}";
        }

        var data = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (data == null)
            return new DocumentNode { Path = path, Type = "sheet", Preview = "(empty)" };

        if (segments.Length == 1)
        {
            // Return sheet overview
            var sheetNode = new DocumentNode
            {
                Path = path,
                Type = "sheet",
                Preview = sheetNameFromPath,
                ChildCount = data.Elements<Row>().Select(r => r.RowIndex?.Value ?? 0u).Where(i => i != 0).Distinct().Count() + (worksheet.DrawingsPart != null ? CountExcelCharts(worksheet.DrawingsPart) : 0)
            };

            // Include freeze pane info
            var ws = GetSheet(worksheet);
            var pane = ws.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>()?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Pane>();
            if (pane != null && pane.State?.Value == PaneStateValues.Frozen)
            {
                sheetNode.Format["freeze"] = pane.TopLeftCell?.Value ?? "";
            }

            // Include zoom and view properties
            var sheetView = ws.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>();
            if (sheetView?.ZoomScale?.HasValue == true && sheetView.ZoomScale.Value != 100)
                sheetNode.Format["zoom"] = (int)sheetView.ZoomScale.Value;
            if (sheetView?.ShowGridLines != null && !sheetView.ShowGridLines.Value)
                sheetNode.Format["gridlines"] = false;
            if (sheetView?.ShowRowColHeaders != null && !sheetView.ShowRowColHeaders.Value)
                sheetNode.Format["headings"] = false;
            if (sheetView?.RightToLeft?.HasValue == true && sheetView.RightToLeft.Value)
                sheetNode.Format["direction"] = "rtl";

            // Include tab color. Excel does not render tab transparency, so
            // strip any alpha component before formatting — `Add tabColor=80FF0000`
            // round-trips as `#FF0000`, mirroring how Excel stores 6-digit RGB
            // when the user picks a tab color in the UI.
            var tabColor = ws.GetFirstChild<SheetProperties>()?.GetFirstChild<TabColor>();
            if (tabColor?.Rgb?.HasValue == true)
            {
                var rgb = tabColor.Rgb.Value!;
                if (rgb.Length == 8) rgb = rgb[2..];
                sheetNode.Format["tabColor"] = ParseHelpers.FormatHexColor(rgb);
            }
            else if (tabColor?.Theme?.HasValue == true)
            {
                // CONSISTENCY(scheme-color): echo back the symbolic name
                // (e.g. "accent1") instead of the numeric theme index.
                var schemeName = ParseHelpers.ExcelThemeIndexToName(tabColor.Theme.Value);
                if (schemeName != null) sheetNode.Format["tabColor"] = schemeName;
            }

            // Include autofilter info
            var autoFilter = ws.GetFirstChild<AutoFilter>();
            if (autoFilter?.Reference?.Value != null)
            {
                sheetNode.Format["autoFilter"] = autoFilter.Reference.Value;
            }

            // Sheet-state (hidden / very hidden) readback — lives on the
            // workbook-level Sheet element, not on the Worksheet.
            var wbSheet = GetWorkbook().GetFirstChild<Sheets>()?.Elements<Sheet>()
                .FirstOrDefault(s => s.Name?.Value?.Equals(sheetNameFromPath, StringComparison.OrdinalIgnoreCase) == true);
            // bt-1 (R25): align with the project-wide toggle-on/key-missing
            // convention used by autoFilter / protect / row.hidden / col.hidden
            // (CONSISTENCY(default-omission)). Default-visible sheets emit no
            // hidden key; hidden=true only when State is Hidden/VeryHidden.
            // Reverts R24 d56ea9d5's always-emit behavior.
            if (wbSheet?.State?.Value is { } sheetState
                && (sheetState == SheetStateValues.Hidden || sheetState == SheetStateValues.VeryHidden))
            {
                sheetNode.Format["hidden"] = true;
                sheetNode.Format["visibility"] = sheetState == SheetStateValues.VeryHidden ? "veryHidden" : "hidden";
            }

            // Sheet protection readback
            var sheetProtection = ws.GetFirstChild<SheetProtection>();
            if (sheetProtection?.Sheet?.Value == true)
            {
                sheetNode.Format["protect"] = true;
                // Password hashes (legacy 16-bit `password` attr and the
                // modern algorithmName/hashValue/saltValue/spinCount set)
                // are surfaced verbatim so dump→batch preserves protection
                // strength instead of silently degrading to passwordless.
                if (sheetProtection.Password?.Value is { Length: > 0 } legacyPw)
                    sheetNode.Format["passwordHash"] = legacyPw;
                if (sheetProtection.AlgorithmName?.Value is { Length: > 0 } algo)
                    sheetNode.Format["protection.algorithm"] = algo;
                if (sheetProtection.HashValue?.Value is { Length: > 0 } hashV)
                    sheetNode.Format["protection.hash"] = hashV;
                if (sheetProtection.SaltValue?.Value is { Length: > 0 } saltV)
                    sheetNode.Format["protection.salt"] = saltV;
                if (sheetProtection.SpinCount?.HasValue == true)
                    sheetNode.Format["protection.spinCount"] = (int)sheetProtection.SpinCount.Value;
            }

            // Print settings readback
            var pageSetup = ws.GetFirstChild<PageSetup>();
            if (pageSetup != null)
            {
                if (pageSetup.Orientation?.HasValue == true)
                    sheetNode.Format["orientation"] = pageSetup.Orientation.InnerText;
                if (pageSetup.PaperSize?.HasValue == true)
                    sheetNode.Format["paperSize"] = (int)pageSetup.PaperSize.Value;
                if (pageSetup.FitToWidth?.HasValue == true)
                    sheetNode.Format["fitToPage"] = $"{pageSetup.FitToWidth.Value}x{pageSetup.FitToHeight?.Value ?? 1}";
            }

            // Print area readback
            var workbook = GetWorkbook();
            var allSheets = workbook.GetFirstChild<Sheets>()?.Elements<Sheet>().ToList();
            var sheetIdx = allSheets?.FindIndex(s =>
                s.Name?.Value?.Equals(sheetNameFromPath, StringComparison.OrdinalIgnoreCase) == true) ?? -1;
            var printAreaDn = workbook.GetFirstChild<DefinedNames>()?.Elements<DefinedName>()
                .FirstOrDefault(d => d.Name == "_xlnm.Print_Area" && d.LocalSheetId?.Value == (uint)sheetIdx);
            if (printAreaDn != null)
            {
                // Strip "SheetName!" prefix so Get output can round-trip to Set input
                var paText = printAreaDn.Text ?? "";
                var bangIdx = paText.IndexOf('!');
                if (bangIdx >= 0) paText = paText[(bangIdx + 1)..];
                sheetNode.Format["printArea"] = paText;
            }

            // Print title rows/cols readback (_xlnm.Print_Titles). Mirrors the
            // Print_Area path so `set printTitleRows/printTitleCols` round-trips.
            // The defined name combines both a row range ($1:$2) and a column
            // range ($A:$A), comma-joined; split by which side is digit/letter.
            var printTitlesDn = workbook.GetFirstChild<DefinedNames>()?.Elements<DefinedName>()
                .FirstOrDefault(d => d.Name == "_xlnm.Print_Titles" && d.LocalSheetId?.Value == (uint)sheetIdx);
            if (printTitlesDn != null)
            {
                foreach (var tok in (printTitlesDn.Text ?? "").Split(',', StringSplitOptions.RemoveEmptyEntries))
                {
                    var t = tok.Trim();
                    var bang = t.IndexOf('!');
                    var rangePart = (bang >= 0 ? t[(bang + 1)..] : t).Replace("$", "");
                    var leftSide = rangePart.Split(':')[0];
                    if (leftSide.Length == 0) continue;
                    if (char.IsDigit(leftSide[0]))
                        sheetNode.Format["printTitleRows"] = rangePart;
                    else if (char.IsLetter(leftSide[0]))
                        sheetNode.Format["printTitleCols"] = rangePart;
                }
            }

            // PageMargins readback
            var pm = ws.GetFirstChild<PageMargins>();
            if (pm != null)
            {
                static string Fmt(double v) => v.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + "in";
                if (pm.Top?.HasValue == true) sheetNode.Format["margin.top"] = Fmt(pm.Top.Value);
                if (pm.Bottom?.HasValue == true) sheetNode.Format["margin.bottom"] = Fmt(pm.Bottom.Value);
                if (pm.Left?.HasValue == true) sheetNode.Format["margin.left"] = Fmt(pm.Left.Value);
                if (pm.Right?.HasValue == true) sheetNode.Format["margin.right"] = Fmt(pm.Right.Value);
                if (pm.Header?.HasValue == true) sheetNode.Format["margin.header"] = Fmt(pm.Header.Value);
                if (pm.Footer?.HasValue == true) sheetNode.Format["margin.footer"] = Fmt(pm.Footer.Value);
            }

            // Header/Footer readback
            var headerFooter = ws.GetFirstChild<HeaderFooter>();
            if (headerFooter?.OddHeader?.Text != null)
                sheetNode.Format["header"] = headerFooter.OddHeader.Text;
            if (headerFooter?.OddFooter?.Text != null)
                sheetNode.Format["footer"] = headerFooter.OddFooter.Text;

            // Sort state readback
            var sortState = ws.GetFirstChild<SortState>();
            if (sortState != null)
            {
                var sortConditions = sortState.Elements<SortCondition>().ToList();
                var sortDesc = string.Join(",", sortConditions.Select(sc =>
                {
                    var colRef = sc.Reference?.Value?.Split(':')[0] ?? "";
                    var colName = Regex.Match(colRef, @"^([A-Z]+)").Groups[1].Value;
                    var dir = sc.Descending?.Value == true ? "desc" : "asc";
                    return $"{colName}:{dir}";
                }));
                sheetNode.Format["sort"] = sortDesc;
            }

            // Page breaks readback
            var rowBreaks = ws.GetFirstChild<RowBreaks>();
            if (rowBreaks != null && rowBreaks.Elements<Break>().Any())
            {
                var breaks = rowBreaks.Elements<Break>().Select(b => b.Id?.Value.ToString() ?? "").ToList();
                sheetNode.Format["rowBreaks"] = string.Join(",", breaks);
            }
            var colBreaks = ws.GetFirstChild<ColumnBreaks>();
            if (colBreaks != null && colBreaks.Elements<Break>().Any())
            {
                var cbreaks = colBreaks.Elements<Break>().Select(b => b.Id?.Value.ToString() ?? "").ToList();
                sheetNode.Format["colBreaks"] = string.Join(",", cbreaks);
            }

            if (depth > 0)
            {
                sheetNode.Children = GetSheetChildNodes(sheetNameFromPath, data, depth, worksheet);
                // Children omit value-less empty cells/rows (issue #149);
                // reflect the actual listed count, not the raw row count.
                sheetNode.ChildCount = sheetNode.Children.Count;
            }
            return sheetNode;
        }

        // BUG-R41-F2: reject cell reference segments that contain control characters
        // (e.g. \n, \r, \t). Without this check, "A1\n" passes the cell-ref regex
        // (Regex `$` matches before trailing \n in .NET) and resolves to a ghost cell.
        var cellRef = segments[1];
        if (cellRef.Any(c => c < ' ' && c != '\t' || c == '\x7f'))
            throw new ArgumentException(
                $"Cell reference '{cellRef.Replace("\n", "\\n").Replace("\r", "\\r")}' contains invalid control characters. " +
                $"Expected a clean cell address like 'A1' or 'B2'.");

        // Page break path: /Sheet1/rowbreak[N] or /Sheet1/colbreak[N]
        var rbMatch = Regex.Match(cellRef, @"^rowbreak\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (rbMatch.Success)
        {
            var rbIdx = int.Parse(rbMatch.Groups[1].Value);
            var rowBreaks = GetSheet(worksheet).GetFirstChild<RowBreaks>();
            var breaks = rowBreaks?.Elements<Break>().ToList() ?? new();
            if (rbIdx < 1 || rbIdx > breaks.Count)
                throw new ArgumentException($"Row break index {rbIdx} out of range (1-{breaks.Count})");
            var brk = breaks[rbIdx - 1];
            var rbNode = new DocumentNode
            {
                Path = path, Type = "rowbreak",
                Format = { ["row"] = brk.Id?.Value ?? 0u, ["manual"] = brk.ManualPageBreak?.Value ?? false }
            };
            // Restricted-span page break (<brk min max>): surface a non-default
            // column span so dump→replay reproduces it. Full-width default
            // (min 0 / max 16383) is omitted to keep the readback clean.
            if (brk.Min?.Value is { } rbMin && rbMin > 0) rbNode.Format["min"] = (int)rbMin;
            if (brk.Max?.Value is { } rbMax && rbMax != 16383u) rbNode.Format["max"] = (int)rbMax;
            return rbNode;
        }
        var cbMatch = Regex.Match(cellRef, @"^colbreak\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (cbMatch.Success)
        {
            var cbIdx = int.Parse(cbMatch.Groups[1].Value);
            var colBreaks = GetSheet(worksheet).GetFirstChild<ColumnBreaks>();
            var breaks = colBreaks?.Elements<Break>().ToList() ?? new();
            if (cbIdx < 1 || cbIdx > breaks.Count)
                throw new ArgumentException($"Column break index {cbIdx} out of range (1-{breaks.Count})");
            var brk = breaks[cbIdx - 1];
            var cbNode = new DocumentNode
            {
                Path = path, Type = "colbreak",
                Format = { ["col"] = (int)(brk.Id?.Value ?? 0u), ["manual"] = brk.ManualPageBreak?.Value ?? false }
            };
            // Restricted-span break: full-height default is min 0 / max 1048575.
            if (brk.Min?.Value is { } cbMin && cbMin > 0) cbNode.Format["min"] = (int)cbMin;
            if (brk.Max?.Value is { } cbMax && cbMax != 1048575u) cbNode.Format["max"] = (int)cbMax;
            return cbNode;
        }

        // Validation path: /Sheet1/dataValidation[N] (canonical) or
        // /Sheet1/validation[N] (legacy alias, R7-bt-6 CONSISTENCY)
        var validationMatch = Regex.Match(cellRef, @"^(?:dataValidation|validation)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (validationMatch.Success)
        {
            var dvIdx = int.Parse(validationMatch.Groups[1].Value);
            var dvs = GetSheet(worksheet).GetFirstChild<DataValidations>();
            if (dvs == null)
                return new DocumentNode { Path = path, Type = "error", Text = $"dataValidation[{dvIdx}] not found (sheet has no data validations)" };

            var dvList = dvs.Elements<DataValidation>().ToList();
            if (dvIdx < 1 || dvIdx > dvList.Count)
                return new DocumentNode { Path = path, Type = "error", Text = $"dataValidation[{dvIdx}] not found (sheet has {dvList.Count} validation(s))" };

            return DataValidationToNode(sheetNameFromPath, dvList[dvIdx - 1], dvIdx);
        }

        // Column path: /Sheet1/col[A]
        var colMatch = Regex.Match(cellRef, @"^col\[([A-Za-z0-9]+)\]$", RegexOptions.IgnoreCase);
        if (colMatch.Success)
        {
            var colValue = colMatch.Groups[1].Value;
            var colName = int.TryParse(colValue, out var numIdx) ? IndexToColumnName(numIdx) : colValue.ToUpperInvariant();
            var colIdx = (uint)ColumnNameToIndex(colName);
            var colNode = new DocumentNode { Path = path, Type = "column", Preview = colName };
            var columns = GetSheet(worksheet).GetFirstChild<Columns>();
            if (columns != null)
            {
                var col = columns.Elements<Column>().FirstOrDefault(c =>
                    c.Min?.Value <= colIdx && c.Max?.Value >= colIdx);
                if (col != null)
                {
                    if (col.Width?.Value != null) colNode.Format["width"] = col.Width.Value;
                    if (col.Hidden?.Value == true) colNode.Format["hidden"] = true;
                    if (col.CustomWidth?.Value == true) colNode.Format["customWidth"] = true;
                    if (col.OutlineLevel?.HasValue == true && col.OutlineLevel.Value > 0)
                        colNode.Format["outlineLevel"] = (int)col.OutlineLevel.Value;
                    if (col.Collapsed?.Value == true) colNode.Format["collapsed"] = true;
                    // Resolve a column-level number format (col @s -> cellXf ->
                    // numFmt) so `set col[A] --prop numFmt=...` round-trips as the
                    // canonical `numberformat` key, mirroring the cell reader.
                    if (col.Style?.Value is uint colStyleIdx && colStyleIdx != 0)
                    {
                        var (colNumFmtId, colFmtCode) = ExcelDataFormatter.GetCellFormat(
                            new Cell { StyleIndex = colStyleIdx }, _doc.WorkbookPart);
                        // A built-in numFmtId (e.g. 4 = "#,##0.00") carries no
                        // <numFmt> entry in the styles part, so GetCellFormat
                        // returns a null code for it. Resolve the built-in code
                        // so column Get surfaces `numberformat` the same way the
                        // cell reader does (CellToNode's built-in id map).
                        if (colNumFmtId > 0)
                        {
                            var colCode = !string.IsNullOrEmpty(colFmtCode)
                                ? colFmtCode
                                : ExcelDataFormatter.ResolveBuiltInFormatCode(colNumFmtId);
                            if (!string.IsNullOrEmpty(colCode))
                                colNode.Format["numberformat"] = colCode;
                            colNode.Format["numFmtId"] = (int)colNumFmtId;
                        }
                    }
                    // Long-tail CT_Col attributes (style, bestFit, phonetic, ...).
                    // Symmetric with column Set's case-preserving SetAttribute fallback.
                    FillUnknownAttrProps(col, colNode, "", CuratedColAttrs);
                }
            }
            // Include cells in this column as children (non-empty rows only)
            if (depth > 0)
            {
                var eval = new Core.FormulaEvaluator(data, _doc.WorkbookPart);
                foreach (var row in data.Elements<Row>().OrderBy(r => r.RowIndex?.Value ?? 0))
                {
                    var cell = row.Elements<Cell>().FirstOrDefault(c =>
                    {
                        if (c.CellReference?.Value == null) return false;
                        var (cn, _) = ParseCellReference(c.CellReference.Value);
                        return cn.Equals(colName, StringComparison.OrdinalIgnoreCase);
                    });
                    if (cell != null)
                        colNode.Children.Add(CellToNode(sheetNameFromPath, cell, worksheet, eval));
                }
                colNode.ChildCount = colNode.Children.Count;
            }
            return colNode;
        }

        // Row path: /Sheet1/row[N] or /Sheet1/row[last()]
        // CONSISTENCY(path-stability): mirrors sheet[last()] support in ResolveSheetName
        // and Word's p[last()] — resolve last() to the highest RowIndex present.
        var rowLastMatch = Regex.Match(cellRef, @"^row\[last\(\)\]$", RegexOptions.IgnoreCase);
        if (rowLastMatch.Success)
        {
            var lastRowIdx = data.Elements<Row>()
                .Select(r => r.RowIndex?.Value ?? 0u)
                .Where(i => i > 0)
                .DefaultIfEmpty(0u)
                .Max();
            if (lastRowIdx == 0)
                return new DocumentNode { Path = path, Type = "row", Text = "(empty)" };
            cellRef = $"row[{lastRowIdx}]";
            path = $"/{sheetNameFromPath}/row[{lastRowIdx}]";
        }
        var rowMatch = Regex.Match(cellRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = uint.Parse(rowMatch.Groups[1].Value);
            var row = data.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx);
            if (row == null)
                return new DocumentNode { Path = path, Type = "row", Preview = $"row {rowIdx}", Text = "(empty)" };
            var rowNode = new DocumentNode
            {
                Path = path, Type = "row", ChildCount = row.Elements<Cell>().Count()
            };
            // CONSISTENCY(unit-qualified-readback): row height is stored in
            // points in OOXML; emit as "{n}pt" so it matches pptx's
            // unit-qualified readback (the project conventions canonical value rule).
            if (row.Height?.Value != null)
                rowNode.Format["height"] = $"{row.Height.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}pt";
            if (row.Hidden?.Value == true) rowNode.Format["hidden"] = true;
            if (row.OutlineLevel?.HasValue == true && row.OutlineLevel.Value > 0)
                rowNode.Format["outlineLevel"] = (int)row.OutlineLevel.Value;
            if (row.Collapsed?.Value == true) rowNode.Format["collapsed"] = true;
            // Long-tail CT_Row attributes (spans, style, ph, thickTop, thickBot,
            // customFormat, ...). Symmetric with row Set's case-preserving fallback.
            FillUnknownAttrProps(row, rowNode, "", CuratedRowAttrs);
            if (depth > 0)
            {
                var eval = new Core.FormulaEvaluator(data, _doc.WorkbookPart);
                foreach (var c in row.Elements<Cell>())
                    rowNode.Children.Add(CellToNode(sheetNameFromPath, c, worksheet, eval));
            }
            return rowNode;
        }

        // Conditional formatting path: /Sheet1/cf[N]
        var cfMatch = Regex.Match(cellRef, @"^cf\[(\d+)\]$");
        if (cfMatch.Success)
        {
            var cfIdx = int.Parse(cfMatch.Groups[1].Value);
            var cfElements = GetSheet(worksheet).Elements<ConditionalFormatting>().ToList();
            if (cfIdx < 1 || cfIdx > cfElements.Count)
                return new DocumentNode { Path = path, Type = "error", Text = $"cf[{cfIdx}] not found (sheet has {cfElements.Count} conditional formatting rule(s))" };

            var cf = cfElements[cfIdx - 1];
            var cfNode = new DocumentNode { Path = path, Type = "conditionalFormatting" };
            cfNode.Format["ref"] = cf.SequenceOfReferences?.InnerText ?? "";

            var rule = cf.Elements<ConditionalFormattingRule>().FirstOrDefault();
            if (rule != null)
                PopulateCfNodeFromRule(worksheet, rule, cfNode);
            return cfNode;
        }

        // AutoFilter path: /Sheet1/autofilter
        if (cellRef.Equals("autofilter", StringComparison.OrdinalIgnoreCase))
        {
            var af = GetSheet(worksheet).GetFirstChild<AutoFilter>();
            var afNode = new DocumentNode { Path = path, Type = "autofilter" };
            if (af?.Reference?.Value != null) afNode.Format["range"] = af.Reference.Value;
            if (af != null) PopulateAutoFilterCriteria(af, afNode);
            return afNode;
        }

        // Chart axis-by-role sub-path: /Sheet1/chart[N]/axis[@role=ROLE].
        // Per schemas/help/pptx/chart-axis.json (shared contract).
        var chartAxisGetMatch = Regex.Match(cellRef,
            @"^chart\[(\d+)\]/axis\[@role=([a-zA-Z0-9_]+)\]$");
        if (chartAxisGetMatch.Success)
        {
            var caChartIdx = int.Parse(chartAxisGetMatch.Groups[1].Value);
            var caRole = chartAxisGetMatch.Groups[2].Value;
            var caDrawingsPart = worksheet.DrawingsPart;
            if (caDrawingsPart == null)
                throw new ArgumentException($"No charts found in sheet");
            var caAllCharts = GetExcelCharts(caDrawingsPart);
            if (caChartIdx < 1 || caChartIdx > caAllCharts.Count)
                throw new ArgumentException($"Chart index {caChartIdx} out of range (1-{caAllCharts.Count})");
            var caChartInfo = caAllCharts[caChartIdx - 1];
            if (caChartInfo.IsExtended || caChartInfo.StandardPart?.ChartSpace == null)
                throw new ArgumentException($"Axis not available on chart {caChartIdx}: extended charts not supported.");
            var axisNode = ChartHelper.BuildAxisNode(caChartInfo.StandardPart.ChartSpace, caRole, path);
            if (axisNode == null)
                throw new ArgumentException($"Axis with role '{caRole}' not found on chart {caChartIdx}.");
            return axisNode;
        }

        // Chart path: /Sheet1/chart[N] or /Sheet1/chart[N]/series[K]
        var chartMatch = Regex.Match(cellRef, @"^chart\[(\d+)\](?:/series\[(\d+)\])?$");
        if (chartMatch.Success)
        {
            var chartIdx = int.Parse(chartMatch.Groups[1].Value);
            var drawingsPart = worksheet.DrawingsPart;
            if (drawingsPart == null)
                throw new ArgumentException($"No charts found in sheet");

            var allCharts = GetExcelCharts(drawingsPart);
            if (chartIdx < 1 || chartIdx > allCharts.Count)
                throw new ArgumentException($"Chart index {chartIdx} out of range (1-{allCharts.Count})");

            var chartInfo = allCharts[chartIdx - 1];
            var chartNode = new DocumentNode { Path = $"/{sheetNameFromPath}/chart[{chartIdx}]", Type = "chart" };

            // BUG-R11-04: chart Get used to skip the TwoCellAnchor even though
            // `add chart --prop anchor=B2:F7` and `set ... anchor=...` both
            // support it. Round-trip requires Get to surface the anchor range
            // in the same `B2:F7` grammar. CONSISTENCY(ole-width-units) —
            // mirrors the Add/Set accepted grammar.
            var chartAnchorRange = GetChartAnchorRange(drawingsPart, chartIdx);
            if (chartAnchorRange != null)
                chartNode.Format["anchor"] = chartAnchorRange;

            // CONSISTENCY(ole-width-units): also surface x/y/width/height in cm,
            // matching the schema's add/set vocabulary so round-trip works.
            PopulateChartPositionFormat(drawingsPart, chartIdx, chartNode);

            if (chartInfo.IsExtended)
            {
                var cxChartSpace = chartInfo.ExtendedPart!.ChartSpace!;
                var cxType = Core.ChartExBuilder.DetectExtendedChartType(cxChartSpace);
                if (cxType != null) chartNode.Format["chartType"] = cxType;
                // Title
                var cxTitle = cxChartSpace.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.ChartTitle>().FirstOrDefault();
                var cxTitleText = cxTitle?.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault()?.Text;
                if (cxTitleText != null) chartNode.Format["title"] = cxTitleText;
                // Count series
                var cxSeries = cxChartSpace.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.Series>().ToList();
                chartNode.Format["seriesCount"] = cxSeries.Count;
            }
            else
            {
                var chart = chartInfo.StandardPart!.ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
                if (chart != null)
                    ChartHelper.ReadChartProperties(chart, chartNode, chartMatch.Groups[2].Success ? 1 : depth);
            }

            // If series sub-path requested, extract the specific series child
            if (chartMatch.Groups[2].Success)
            {
                var seriesIdx = int.Parse(chartMatch.Groups[2].Value);
                var seriesChildren = chartNode.Children.Where(c => c.Type == "series").ToList();
                if (seriesIdx < 1 || seriesIdx > seriesChildren.Count)
                    throw new ArgumentException($"Series {seriesIdx} not found (total: {seriesChildren.Count})");
                var seriesNode = seriesChildren[seriesIdx - 1];
                seriesNode.Path = path;
                return seriesNode;
            }
            return chartNode;
        }

        // Pivot table path: /Sheet1/pivottable[N]
        var pivotMatch = Regex.Match(cellRef, @"^pivottable\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (pivotMatch.Success)
        {
            var ptIdx = int.Parse(pivotMatch.Groups[1].Value);
            var pivotParts = worksheet.PivotTableParts.ToList();
            if (ptIdx < 1 || ptIdx > pivotParts.Count)
                throw new ArgumentException($"PivotTable index {ptIdx} out of range (1-{pivotParts.Count})");

            var pivotPart = pivotParts[ptIdx - 1];
            var ptNode = new DocumentNode { Path = path, Type = "pivottable" };
            if (pivotPart.PivotTableDefinition != null)
                PivotTableHelper.ReadPivotTableProperties(pivotPart.PivotTableDefinition, ptNode, pivotPart);
            return ptNode;
        }

        // Slicer path: /Sheet1/slicer[N]
        var slicerMatch = Regex.Match(cellRef, @"^slicer\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (slicerMatch.Success)
        {
            var slIdx = int.Parse(slicerMatch.Groups[1].Value);
            if (!TryFindSlicerByIndex(worksheet, slIdx, out var slicerElem, out var slicerCache) || slicerElem == null)
                throw new ArgumentException($"slicer[{slIdx}] not found on sheet '{sheetNameFromPath}'");
            var slNode = new DocumentNode { Path = path, Type = "slicer" };
            ReadSlicerProperties(slicerElem, slicerCache, slNode);
            return slNode;
        }

        // OLE object path: /Sheet1/ole[N]
        // CONSISTENCY(ole-alias): "oleobject" mirrors Add's case switch
        var oleMatch = Regex.Match(cellRef, @"^(?:ole|oleobject|object|embed)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (oleMatch.Success)
        {
            var oleIdx = int.Parse(oleMatch.Groups[1].Value);
            var oleList = CollectOleNodesForSheet(sheetNameFromPath, worksheet);
            if (oleIdx < 1 || oleIdx > oleList.Count)
                throw new ArgumentException($"OLE object {oleIdx} not found at /{sheetNameFromPath} (available: {oleList.Count}).");
            return oleList[oleIdx - 1];
        }

        // Comment path: /Sheet1/comment[N]
        var commentMatch = Regex.Match(cellRef, @"^comment\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (commentMatch.Success)
        {
            var cmtIndex = int.Parse(commentMatch.Groups[1].Value);
            var commentsPart = worksheet.WorksheetCommentsPart;
            if (commentsPart?.Comments == null)
                return new DocumentNode { Path = path, Type = "error", Text = $"comment[{cmtIndex}] not found (sheet has no comments)" };

            var cmtList = commentsPart.Comments.GetFirstChild<CommentList>();
            var cmtElement = cmtList?.Elements<Comment>().ElementAtOrDefault(cmtIndex - 1);
            if (cmtElement == null)
                return new DocumentNode { Path = path, Type = "error", Text = $"comment[{cmtIndex}] not found" };

            return CommentToNode(sheetNameFromPath, cmtElement, commentsPart.Comments, cmtIndex);
        }

        // Table path: /Sheet1/table[N]
        var tableMatch = Regex.Match(cellRef, @"^table\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (tableMatch.Success)
        {
            var tableIdx = int.Parse(tableMatch.Groups[1].Value);
            return TableToNode(sheetNameFromPath, worksheet, tableIdx, depth);
        }

        // Table column path: /Sheet1/table[N]/columns[M] or /column[M]
        var tableColMatch = Regex.Match(cellRef,
            @"^table\[(\d+)\]/(?:columns|column)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (tableColMatch.Success)
        {
            var tIdx = int.Parse(tableColMatch.Groups[1].Value);
            var cIdx = int.Parse(tableColMatch.Groups[2].Value);
            var tParts = worksheet.TableDefinitionParts.ToList();
            if (tIdx < 1 || tIdx > tParts.Count)
                throw new ArgumentException($"Table index {tIdx} out of range (1..{tParts.Count})");
            var tbl = tParts[tIdx - 1].Table
                ?? throw new ArgumentException($"Table {tIdx} has no definition");
            var tCols = tbl.GetFirstChild<TableColumns>()?.Elements<TableColumn>().ToList();
            if (tCols == null || cIdx < 1 || cIdx > tCols.Count)
                throw new ArgumentException($"Column index {cIdx} out of range (1..{tCols?.Count ?? 0})");
            var tCol = tCols[cIdx - 1];
            var tcNode = new DocumentNode
            {
                Path = $"/{sheetNameFromPath}/table[{tIdx}]/columns[{cIdx}]",
                Type = "tableColumn",
                Text = tCol.Name?.Value ?? ""
            };
            tcNode.Format["name"] = tCol.Name?.Value ?? "";
            if (tCol.Id?.Value != null) tcNode.Format["id"] = tCol.Id.Value;
            if (tCol.TotalsRowFunction?.HasValue == true)
                // Open XML SDK v3 EnumValue<T>.ToString() returns
                // "TotalsRowFunctionValues { }" — use InnerText for the
                // OOXML-canonical lowercase token. CONSISTENCY(enum-innertext).
                tcNode.Format["totalFunction"] = tCol.TotalsRowFunction.InnerText;
            if (tCol.TotalsRowLabel?.Value != null)
                tcNode.Format["totalLabel"] = tCol.TotalsRowLabel.Value;
            var ccf = tCol.CalculatedColumnFormula?.Text;
            if (!string.IsNullOrEmpty(ccf)) tcNode.Format["formula"] = ccf;
            return tcNode;
        }

        // Cell reference: A1 or range A1:D10
        // Check if it's a cell reference or a generic XML path
        var firstPart = cellRef.Split('/')[0].Split('[')[0];
        bool isCellRef = System.Text.RegularExpressions.Regex.IsMatch(firstPart, @"^[A-Z]+\d+", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        if (!isCellRef)
        {
            // Handle sparkline[N] path segment
            var spkMatch = Regex.Match(cellRef, @"^sparkline\[(\d+)\]$", RegexOptions.IgnoreCase);
            if (spkMatch.Success)
            {
                var spkIndex = int.Parse(spkMatch.Groups[1].Value);
                var spkGroup = GetSparklineGroup(worksheet, spkIndex)
                    ?? throw new ArgumentException($"Sparkline[{spkIndex}] not found in sheet '{sheetNameFromPath}'");
                return SparklineGroupToNode(sheetNameFromPath, spkGroup, spkIndex);
            }

            // Handle picture[N] path segment
            var picMatch = Regex.Match(cellRef, @"^picture\[(\d+)\]$", RegexOptions.IgnoreCase);
            if (picMatch.Success)
            {
                var picIndex = int.Parse(picMatch.Groups[1].Value);
                // GetPictureNode returns null for out-of-range indices (incl.
                // picture[0]); the bare `!` leaked a NullReferenceException as
                // an opaque internal_error instead of the not-found message
                // every sibling element type produces.
                return GetPictureNode(sheetNameFromPath, worksheet, picIndex, path)
                    ?? throw new ArgumentException(
                        $"Picture[{picIndex}] not found in sheet '{sheetNameFromPath}' (indices are 1-based).");
            }

            // Handle shape[N] path segment
            var shpMatch = Regex.Match(cellRef, @"^shape\[(\d+)\]$", RegexOptions.IgnoreCase);
            if (shpMatch.Success)
            {
                var shpIndex = int.Parse(shpMatch.Groups[1].Value);
                // Same null-leak as picture[N] above.
                return GetShapeNode(sheetNameFromPath, worksheet, shpIndex, path)
                    ?? throw new ArgumentException(
                        $"Shape[{shpIndex}] not found in sheet '{sheetNameFromPath}' (indices are 1-based).");
            }


            // If it looks like it could be a malformed cell reference (digits only, etc.), reject it
            if (Regex.IsMatch(cellRef, @"^\d+$"))
                throw new ArgumentException($"Invalid cell reference: '{cellRef}'. Expected format like 'A1', 'B2'.");

            // CONSISTENCY(axis-ref-compat): Excel-style whole-column/row
            // references (B:B, 1:1) are input aliases for col[X]/row[N] —
            // re-dispatch a single-axis span to the canonical path (readback
            // Path stays canonical). Multi-axis spans (B:D) have no single
            // node to return; point at the bracket syntax instead.
            if (TryExpandAxisRef(cellRef) is { } axisSegments)
            {
                if (axisSegments.Count == 1)
                    return Get($"/{sheetNameFromPath}/{axisSegments[0]}", depth);
                throw new ArgumentException(
                    $"{cellRef} spans multiple {(char.IsDigit(cellRef[0]) ? "rows" : "columns")} — get them one at a time ({axisSegments[0]} … {axisSegments[^1]}); set accepts the whole span.");
            }

            // Generic XML fallback: navigate worksheet XML tree
            var xmlSegments = GenericXmlQuery.ParsePathSegments(cellRef);
            var target = GenericXmlQuery.NavigateByPath(GetSheet(worksheet), xmlSegments);
            if (target == null)
                return new DocumentNode { Path = path, Type = "error", Text = $"Element not found: {cellRef}" };
            return GenericXmlQuery.ElementToNode(target, path, depth);
        }

        // Handle /SheetName/A1/run[N] (rich text run direct access)
        var runGetMatch = Regex.Match(cellRef, @"^([A-Z]+\d+)/run\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (runGetMatch.Success)
        {
            var runCellRef = runGetMatch.Groups[1].Value.ToUpperInvariant();
            var runIdx = int.Parse(runGetMatch.Groups[2].Value);
            ParseCellReference(runCellRef);
            var runCell = FindCell(data, runCellRef);
            if (runCell == null)
                throw new ArgumentException($"Cell {runCellRef} not found");
            if (runCell.DataType?.Value != CellValues.SharedString ||
                !int.TryParse(runCell.CellValue?.Text, out var sstIdx))
                throw new ArgumentException($"Cell {runCellRef} is not a rich text cell");
            var sstPart = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            var ssi = sstPart?.SharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(sstIdx);
            if (ssi == null) throw new ArgumentException($"SharedString entry {sstIdx} not found");
            var runs = ssi.Elements<Run>().ToList();
            if (runIdx < 1 || runIdx > runs.Count)
                throw new ArgumentException($"Run index {runIdx} out of range (1-{runs.Count})");
            return RunToNode(runs[PathIndex.ToArrayIndex(runIdx)], $"/{sheetNameFromPath}/{runCellRef}/run[{runIdx}]");
        }

        if (cellRef.Contains(':'))
        {
            // Range — validate both endpoints
            var rangeParts = cellRef.Split(':');
            ParseCellReference(rangeParts[0]);
            if (rangeParts.Length > 1) ParseCellReference(rangeParts[1]);
            return GetCellRange(sheetNameFromPath, data, cellRef, depth, worksheet);
        }
        else
        {
            // Single cell — validate cell reference
            ParseCellReference(cellRef);
            var cell = FindCell(data, cellRef);
            if (cell == null)
            {
                var emptyNode = new DocumentNode { Path = path, Type = "cell", Text = "(empty)", Preview = cellRef };
                // Still check merge status for empty cells — they may be part of a merged range
                if (worksheet != null)
                {
                    var mergeCells = GetSheet(worksheet).GetFirstChild<MergeCells>();
                    if (mergeCells != null)
                    {
                        var mergeCell = mergeCells.Elements<MergeCell>()
                            .FirstOrDefault(m => IsCellInMergeRange(cellRef, m.Reference?.Value));
                        if (mergeCell != null)
                        {
                            var mergeRef = mergeCell.Reference?.Value ?? "";
                            emptyNode.Format["merge"] = mergeRef;
                            if (mergeRef.Split(':')[0].Equals(cellRef, StringComparison.OrdinalIgnoreCase))
                                emptyNode.Format["mergeAnchor"] = true;
                        }
                    }
                }
                return emptyNode;
            }
            return CellToNode(sheetNameFromPath, cell, worksheet);
        }
    }

    public List<DocumentNode> Query(string selector)
        => FilterSelectorPositionalIndex(selector, QueryDispatch(selector));

    // Collection element types whose Query result Paths carry a positional
    // `/elem[N]` tail (row uses sparse RowIndex, col uses a letter, the rest are
    // 1-based ordinals). `hyperlink` is excluded — it is cell-backed (Path
    // `/Sheet/A1`), so it has no positional `[N]` to resolve.
    private static readonly HashSet<string> PositionalIndexElements = new(StringComparer.OrdinalIgnoreCase)
    {
        "row", "col", "column", "shape", "picture", "image", "media",
        "chart", "table", "listobject", "ole", "oleobject", "object", "embed",
        "comment", "note", "slicer", "pivottable", "pivot",
        "namedrange", "definedname", "sparkline", "validation", "datavalidation",
        "rowbreak", "colbreak",
    };

    // A pure positional bracket in the selector / Excel-native form —
    // `Sheet1!shape[2]`, `row[2]`, `col[A]`, `col[2]` — selects that ONE element,
    // matching the Get path `/Sheet1/shape[2]`. Without this the bracket was
    // silently dropped (ParseCellSelector only captures operator brackets) and
    // EVERY element of that type matched. Harmless for read-only query, but a
    // footgun once Set/Remove began routing non-slash selectors through Query
    // (`set "Sheet1!shape[2]"` then mutated all shapes while reporting success).
    //
    // A bracket carrying an operator (`value>100`, `hidden=true`) or a bare
    // has-attribute filter (`row[hidden]`) is NOT positional — it is left to the
    // existing handler/AttributeFilter pipeline (the numeric/letter guards below
    // exclude those). Within a single Query the results are homogeneous (one
    // element type), so matching the trailing `[token]` uniquely identifies the
    // indexed node regardless of canonical-vs-alias element naming.
    private static List<DocumentNode> FilterSelectorPositionalIndex(string selector, List<DocumentNode> results)
    {
        if (results.Count == 0) return results;

        // Strip the sheet prefix (Sheet1! — but not a != operator) and any
        // leading /Sheet/ so we see the bare `elem[token]`.
        var s = StripTopLevelSheetPrefix(selector);
        if (s.StartsWith('/'))
        {
            var t = s.TrimStart('/');
            var i = t.IndexOf('/');
            s = i < 0 ? t : t[(i + 1)..];
        }

        var m = Regex.Match(s, @"^(\w+)\[\s*([A-Za-z0-9]+)\s*\]$");
        if (!m.Success) return results;
        var elem = m.Groups[1].Value;
        if (!PositionalIndexElements.Contains(elem)) return results;
        var token = m.Groups[2].Value;

        string wanted;
        if (elem.Equals("col", StringComparison.OrdinalIgnoreCase)
            || elem.Equals("column", StringComparison.OrdinalIgnoreCase))
        {
            // col[A] (column letter) or col[2] (1-based numeric → letter), per
            // the Get path semantics. A non-column-letter token (e.g. a stray
            // word) is not positional — leave the results untouched.
            if (int.TryParse(token, out var ci))
                wanted = IndexToColumnName(ci);
            else if (Regex.IsMatch(token, @"^[A-Za-z]{1,3}$"))
                wanted = token.ToUpperInvariant();
            else
                return results;
        }
        else
        {
            // Every other collection indexes by a 1-based ordinal. A pure-alpha
            // token on these is a has-attribute filter (`row[hidden]`), not an
            // index — leave it to the filter pipeline.
            if (!int.TryParse(token, out _)) return results;
            wanted = token;
        }

        return results.Where(n =>
            n.Path.EndsWith($"[{wanted}]", StringComparison.OrdinalIgnoreCase)).ToList();
    }

    // True when an `=` (equality predicate) sits outside every bracket and quote
    // — i.e. a predicate was written without its brackets (`Dept=IT`,
    // `col.部门=销售`). Only `=` is checked: `>` / `<` double as the descendant
    // combinator (`row > cell`, `table > row`), so flagging them would break
    // those selectors; the bare-`>` predicate case is instead steered by the
    // ambiguity/collision errors, which now spell out the bracketed form. `=` is
    // never a combinator, so a top-level `=` is unambiguously a missing-brackets
    // predicate — fail loud instead of returning a silent empty set (exit 0).
    private static bool HasTopLevelComparison(string selector)
    {
        int bracket = 0, paren = 0;
        char? quote = null;
        foreach (var c in selector)
        {
            if (quote.HasValue) { if (c == quote.Value) quote = null; continue; }
            if (c == '"' || c == '\'') { quote = c; continue; }
            else if (c == '[') bracket++;
            else if (c == ']') bracket = System.Math.Max(0, bracket - 1);
            else if (c == '(') paren++;
            else if (c == ')') paren = System.Math.Max(0, paren - 1);
            else if (bracket == 0 && paren == 0 && c == '=')
                return true;
        }
        return false;
    }

    // Strip a `Sheet!` prefix, recognizing only a TOP-LEVEL '!' (outside
    // brackets and quotes) as the sheet separator. The old regex ^.+?!(?!=)
    // was quote-blind: a predicate value containing '!' — row[F="x!"],
    // row[Msg~=hello!] — was truncated at the value's bang and the selector
    // silently dispatched as an unknown element (0 matches, no warning). A
    // top-level '!' followed by '=' (a bare != filter) is not a separator.
    private static string StripTopLevelSheetPrefix(string selector)
    {
        var i = Core.SelectorCommaSplit.TopLevelIndexOf(selector, '!');
        if (i < 0 || (i + 1 < selector.Length && selector[i + 1] == '=')) return selector;
        return selector[(i + 1)..];
    }

    // True when a '/' sits outside every bracket and quote — the sheet/path
    // separator of a slash path, as opposed to a '/' inside a predicate value.
    // Shared scan with MutationSelectorGuard so query and set/remove agree on
    // what counts as a scoped slash path.
    private static bool HasTopLevelSlash(string selector)
        => Core.SelectorCommaSplit.ContainsTopLevelChar(selector, '/');

    private List<DocumentNode> QueryDispatch(string selector)
    {
        var results = new List<DocumentNode>();

        // A slash-path copied from Get output ("/Sheet1/row[2]") that lost its
        // leading slash ("Sheet1/row[2]") would otherwise read as an unknown
        // element type and return a silent empty set. A top-level '/' (outside
        // brackets — a value's '/' like row[url~=a/b] does not count) with no
        // leading slash is unambiguously that mistake; restore the slash so it
        // resolves the same as the copied path.
        if (!selector.StartsWith("/") && HasTopLevelSlash(selector))
            selector = "/" + selector;

        // Handle Excel-native direct cell ref: Sheet1!A1 or Sheet1!A1:D10
        // For ranges (containing ':'), expand the "range" container node into its
        // individual cell children so Query returns a flat list consistent with all
        // other Query branches. Single-cell refs return a one-element list.
        var nativeCellRef = Regex.Match(selector, @"^([^/!]+)!([A-Z]+\d+(:[A-Z]+\d+)?)$", RegexOptions.IgnoreCase);
        if (nativeCellRef.Success)
        {
            // 'My Data (2024)'!A1 — Excel requires quoting names with spaces;
            // strip the quotes so the DOM path resolves the real sheet.
            var node = Get($"/{UnquoteSheetName(nativeCellRef.Groups[1].Value)}/{nativeCellRef.Groups[2].Value}");
            if (node.Type == "range" && node.Children.Count > 0)
                return node.Children;
            return [node];
        }

        // A comparison operator OUTSIDE any bracket means the predicate was
        // written without its brackets (`col.2024>150`, `foo>1`, `Dept=IT`).
        // Such a selector matches no element type and would otherwise return an
        // empty list with exit 0 — a silent "no rows" that a data user reads as a
        // real result. Fail loud with the bracketed form instead.
        if (HasTopLevelComparison(selector))
            throw new Core.CliException(
                $"'{selector}' is not a valid selector: a predicate must be inside brackets, " +
                $"e.g. row[col.2024>150] to filter table rows by a column, or cell[value>150] to filter cells.")
                { Code = "invalid_selector" };

        // CONSISTENCY(excel-sheet-separator-warn): Detect the PPT-style `>`
        // separator form (e.g. `Sheet1>ole`) that users familiar with the
        // PowerPoint query grammar may try against Excel. Excel uses `!`
        // (Sheet1!cell[...]) — the legacy spreadsheet separator — so a `>`
        // in the sheet-prefix slot will silently fall through to generic
        // XML and return an empty result. We emit a single stderr warning
        // pointing to the correct `!` form, then let the normal flow run.
        // Only fire when the prefix looks like a sheet name (no `/`) and
        // the suffix is a known Excel element type we would have handled.
        {
            var pptStyle = Regex.Match(selector, @"^([^/!>]+)>(\w+)");
            if (pptStyle.Success)
            {
                var suffixType = pptStyle.Groups[2].Value.ToLowerInvariant();
                if (suffixType is "ole" or "oleobject" or "object" or "embed" or "cell" or "row"
                    or "chart" or "pivottable" or "pivot" or "slicer" or "shape"
                    or "picture" or "table" or "listobject" or "comment" or "note"
                    or "validation" or "namedrange" or "definedname" or "media"
                    or "image" or "sparkline")
                {
                    Console.Error.WriteLine(
                        $"Warning: Excel uses '!' not '>' as sheet separator " +
                        $"(e.g. '{pptStyle.Groups[1].Value}!{suffixType}' not " +
                        $"'{pptStyle.Groups[1].Value}>{suffixType}').");
                }
            }
        }

        // CONSISTENCY(merge-alias): OOXML element is <mergeCell>, but users
        // naturally type the semantic name `merge` (matches the `merge` key
        // returned by Get on a cell, and the `merge=...` prop on Set). Also
        // accept `mergedrange`. Rewrite to the real element name so the
        // generic-XML fallback below matches.
        selector = Regex.Replace(selector, @"(^|!)(merge|mergedrange)\b", "$1mergeCell", RegexOptions.IgnoreCase);

        // Check if element type is known (Scheme A) or should fall back to generic XML (Scheme B)
        // Strip sheet prefix (Sheet1!cell[...]) but not != operator.
        // Also accept path-style: "/element" (bare) and "/Sheet/element[...]"
        // so the same dispatch table catches query "/sheet", "/table",
        // "/namedrange", and "/Sheet1/table[1]". Mirrors GET's bare-path
        // listers (see ecb36111) — without this normalization, the bare
        // path falls through to ParseCellSelector and returns cells.
        var selectorForType = StripTopLevelSheetPrefix(selector);
        if (selectorForType.StartsWith('/'))
        {
            var trimmed = selectorForType.TrimStart('/');
            var slashIdx = trimmed.IndexOf('/');
            // "/element..."   → "element..."
            // "/Sheet/elem"   → "elem" (the trailing element governs dispatch;
            //                  parsed.Sheet captures the prefix separately).
            selectorForType = slashIdx < 0 ? trimmed : trimmed[(slashIdx + 1)..];
        }
        var elementMatch = Regex.Match(selectorForType, @"^(\w+)");
        // Lowercase once so all downstream `elementName is "..."` dispatch is
        // case-insensitive. CONSISTENCY(query-case-insensitive): matches how
        // WordHandler.Query normalizes selector.element to lowercase.
        var elementName = elementMatch.Success ? elementMatch.Groups[1].Value.ToLowerInvariant() : "";
        bool isKnownType = string.IsNullOrEmpty(elementName)
            // CONSISTENCY(ole-alias): "oleobject" mirrors Add's case switch
            || elementName is "cell" or "row" or "col" or "column" or "sheet" or "validation" or "comment" or "note" or "table" or "listobject" or "chart" or "pivottable" or "pivot" or "slicer" or "shape" or "picture" or "sparkline" or "namedrange" or "definedname" or "media" or "image" or "ole" or "oleobject" or "object" or "embed" or "hyperlink" or "rowbreak" or "colbreak"
            || (elementName.Length <= 3 && Regex.IsMatch(elementName, @"^[A-Z]+$", RegexOptions.IgnoreCase));
        if (!isKnownType)
        {
            // Scheme B: generic XML fallback
            var genericParsed = GenericXmlQuery.ParseSelector(selector);
            foreach (var (_, worksheetPart) in GetWorksheets())
            {
                results.AddRange(GenericXmlQuery.Query(
                    GetSheet(worksheetPart), genericParsed.element, genericParsed.attrs, genericParsed.containsText));
            }
            return results;
        }

        // CONSISTENCY(query-combinator-xlsx): "row > cell" (and space combinator
        // "row cell") — LHS is a parent scope hint, RHS is the target element type.
        // ParseCellSelector ignores the combinator and extracts only the LHS type,
        // so "row > cell" dispatches as "row" and returns rows instead of cells.
        // Detect the pattern early and re-dispatch with the RHS selector so the
        // correct branch fires.  Same fix applies to any "X > cell" variant.
        var xlCombinatorMatch = Regex.Match(selectorForType, @"^\w[\w\[\]!=@'""\.]*\s*[> ]\s*(.+)$");
        if (xlCombinatorMatch.Success)
        {
            var rhsSelector = xlCombinatorMatch.Groups[1].Value.Trim();
            var rhsType = Regex.Match(rhsSelector, @"^(\w+)").Groups[1].Value.ToLowerInvariant();
            // Only redirect when RHS is a known cell-level type; otherwise fall through
            // to let ParseCellSelector handle it (e.g. "sheet > row" should stay "row").
            if (rhsType is "cell")
                return Query(rhsSelector);
        }

        var parsed = ParseCellSelector(selector);

        // Handle validation queries
        if (elementName == "validation")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var dvs = GetSheet(worksheetPart).GetFirstChild<DataValidations>();
                if (dvs == null) continue;

                var dvList = dvs.Elements<DataValidation>().ToList();
                for (int i = 0; i < dvList.Count; i++)
                    results.Add(DataValidationToNode(sheetName, dvList[i], i + 1));
            }
            return results;
        }

        // Handle comment queries
        if (elementName is "comment" or "note")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var commentsPart = worksheetPart.WorksheetCommentsPart;
                if (commentsPart?.Comments == null) continue;

                var cmtList = commentsPart.Comments.GetFirstChild<CommentList>();
                if (cmtList == null) continue;

                var cmtElements = cmtList.Elements<Comment>().ToList();
                for (int i = 0; i < cmtElements.Count; i++)
                    results.Add(CommentToNode(sheetName, cmtElements[i], commentsPart.Comments, i + 1));
            }
            return results;
        }

        // Handle table queries.
        // CONSISTENCY(xlsx/detected-table): `listobject` returns only real
        // ListObjects (OOXML <table> elements). `table` additionally returns
        // heuristically detected table-shaped blocks (type="detectedtable",
        // stable=false) so an agent gets both real and inferred tables in one
        // call. Detection never fabricates /table[N] paths — detected blocks
        // carry honest range paths (/Sheet!A1:D10).
        if (elementName is "table" or "listobject")
        {
            bool includeDetected = elementName == "table";
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var tableParts = worksheetPart.TableDefinitionParts.ToList();
                var realRanges = new List<(int c1, int r1, int c2, int r2)>();
                for (int i = 0; i < tableParts.Count; i++)
                {
                    results.Add(TableToNode(sheetName, worksheetPart, i + 1, 0));
                    if (includeDetected && TryParseRange(tableParts[i].Table?.Reference?.Value, out var rng))
                        realRanges.Add(rng);
                }
                if (includeDetected)
                    results.AddRange(DetectTables(sheetName, worksheetPart, realRanges));
            }
            return results;
        }

        // Handle chart queries
        if (elementName == "chart")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart == null) continue;

                var allCharts = GetExcelCharts(drawingsPart);
                for (int i = 0; i < allCharts.Count; i++)
                {
                    var chartInfo = allCharts[i];
                    var node = new DocumentNode { Path = $"/{sheetName}/chart[{i + 1}]", Type = "chart" };

                    if (chartInfo.IsExtended)
                    {
                        var cxChartSpace = chartInfo.ExtendedPart!.ChartSpace!;
                        var cxType = Core.ChartExBuilder.DetectExtendedChartType(cxChartSpace);
                        if (cxType != null) node.Format["chartType"] = cxType;
                        var cxTitle = cxChartSpace.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.ChartTitle>().FirstOrDefault();
                        var cxTitleText = cxTitle?.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault()?.Text;
                        if (cxTitleText != null) node.Format["title"] = cxTitleText;
                        var cxSeries = cxChartSpace.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.Series>().ToList();
                        node.Format["seriesCount"] = cxSeries.Count;
                    }
                    else
                    {
                        var chart = chartInfo.StandardPart!.ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
                        if (chart != null)
                            ChartHelper.ReadChartProperties(chart, node, 0);
                    }

                    // Filter by contains text (match on title)
                    if (parsed.ValueContains != null)
                    {
                        var title = node.Format.TryGetValue("title", out var t) ? t?.ToString() : null;
                        if (title == null || !title.Contains(parsed.ValueContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    results.Add(node);
                }
            }
            return results;
        }

        // Handle sheet queries
        if (elementName == "sheet")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var sheetNode = new DocumentNode { Path = $"/{sheetName}", Type = "sheet", Preview = sheetName };
                var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
                var rowCount = sheetData?.Elements<Row>().Count() ?? 0;
                var chartCount = worksheetPart.DrawingsPart != null ? CountExcelCharts(worksheetPart.DrawingsPart) : 0;
                sheetNode.ChildCount = rowCount + chartCount;
                results.Add(sheetNode);
            }
            return results;
        }

        // Handle pivottable queries
        if (elementName == "pivottable" || elementName == "pivot")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var pivotParts = worksheetPart.PivotTableParts.ToList();
                for (int i = 0; i < pivotParts.Count; i++)
                {
                    var node = new DocumentNode { Path = $"/{sheetName}/pivottable[{i + 1}]", Type = "pivottable" };
                    var pivotDef = pivotParts[i].PivotTableDefinition;
                    if (pivotDef != null)
                        PivotTableHelper.ReadPivotTableProperties(pivotDef, node, pivotParts[i]);

                    if (parsed.ValueContains != null)
                    {
                        var name = node.Format.TryGetValue("name", out var n) ? n?.ToString() : null;
                        if (name == null || !name.Contains(parsed.ValueContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    results.Add(node);
                }
            }
            return results;
        }

        // Handle slicer queries
        if (elementName == "slicer")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var slicersPart = worksheetPart.GetPartsOfType<SlicersPart>().FirstOrDefault();
                if (slicersPart?.Slicers == null) continue;

                var slicers = slicersPart.Slicers.Elements<X14.Slicer>().ToList();
                for (int i = 0; i < slicers.Count; i++)
                {
                    if (!TryFindSlicerByIndex(worksheetPart, i + 1, out var slElem, out var slCache) || slElem == null)
                        continue;
                    var node = new DocumentNode
                    {
                        Path = $"/{sheetName}/slicer[{i + 1}]",
                        Type = "slicer"
                    };
                    ReadSlicerProperties(slElem, slCache, node);

                    if (parsed.ValueContains != null)
                    {
                        var nm = node.Format.TryGetValue("name", out var n) ? n?.ToString() : null;
                        if (nm == null || !nm.Contains(parsed.ValueContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    results.Add(node);
                }
            }
            return results;
        }

        // Handle sparkline queries
        if (elementName == "sparkline")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var ws = GetSheet(worksheetPart);
                var extList = ws.GetFirstChild<WorksheetExtensionList>();
                if (extList == null) continue;

                var spkExt = extList.Elements<WorksheetExtension>()
                    .FirstOrDefault(e => e.Uri == "{05C60535-1F16-4fd2-B633-E4A46CF9E463}");
                if (spkExt == null) continue;

                var spkGroups = spkExt.GetFirstChild<X14.SparklineGroups>();
                if (spkGroups == null) continue;

                var groups = spkGroups.Elements<X14.SparklineGroup>().ToList();
                for (int i = 0; i < groups.Count; i++)
                    results.Add(SparklineGroupToNode(sheetName, groups[i], i + 1));
            }
            return results;
        }

        // Page break queries: rowbreak / colbreak. Enumerate the RowBreaks /
        // ColumnBreaks child of each sheet, mirroring the per-sheet sparkline
        // dispatch above and producing the same node shape as the Get path
        // (/Sheet1/{row,col}break[N], see :407-435).
        if (elementName is "rowbreak" or "colbreak")
        {
            var wantRow = elementName == "rowbreak";
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var ws = GetSheet(worksheetPart);
                var breaks = wantRow
                    ? ws.GetFirstChild<RowBreaks>()?.Elements<Break>().ToList()
                    : ws.GetFirstChild<ColumnBreaks>()?.Elements<Break>().ToList();
                if (breaks == null) continue;

                for (int i = 0; i < breaks.Count; i++)
                {
                    var brk = breaks[i];
                    var node = new DocumentNode
                    {
                        Path = $"/{sheetName}/{elementName}[{i + 1}]",
                        Type = elementName,
                        Format = wantRow
                            ? new() { ["row"] = brk.Id?.Value ?? 0u, ["manual"] = brk.ManualPageBreak?.Value ?? false }
                            : new() { ["col"] = (int)(brk.Id?.Value ?? 0u), ["manual"] = brk.ManualPageBreak?.Value ?? false }
                    };
                    results.Add(node);
                }
            }
            return results;
        }

        // Handle shape queries
        if (elementName == "shape")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart?.WorksheetDrawing == null) continue;

                // CONSISTENCY(xlsx-group-flatten): enumerate every leaf shape
                // across all anchors, so an anchor that wraps a GroupShape
                // with N inner shapes contributes N results (was 1). Anchors
                // without a group still contribute exactly one shape so the
                // common-case shape index numbering does not change.
                var leafShapeCount = EnumerateLeafShapes(drawingsPart.WorksheetDrawing).Count();
                for (int i = 0; i < leafShapeCount; i++)
                {
                    var node = GetShapeNode(sheetName, worksheetPart, i + 1, $"/{sheetName}/shape[{i + 1}]");
                    if (node == null) continue;

                    if (parsed.ValueContains != null)
                    {
                        if (node.Text == null || !node.Text.Contains(parsed.ValueContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    if (MatchesFormatAttributes(node, parsed))
                        results.Add(node);
                }
            }
            return results;
        }

        // Handle OLE object queries. Excel stores OLE objects in two
        // parallel structures:
        //   1. <oleObjects> inside the worksheet (schema-typed OleObject
        //      elements with progId + shapeId + r:id)
        //   2. EmbeddedObjectParts/EmbeddedPackageParts on the WorksheetPart
        //      (the actual binary payloads, joined via rel id)
        // We enumerate (1) as the source of truth for path indexing and
        // join (2) for contentType/fileSize enrichment. Worksheets that
        // somehow have orphan parts without a matching oleObjects entry
        // are still surfaced from the parts side so nothing is missed.
        // CONSISTENCY(ole-alias): "oleobject" mirrors Add's case switch
        if (elementName is "ole" or "oleobject" or "object" or "embed")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var oleNodes = CollectOleNodesForSheet(sheetName, worksheetPart);
                foreach (var node in oleNodes)
                {
                    if (parsed.ValueContains != null)
                    {
                        var pid = node.Format.TryGetValue("progId", out var p) ? p?.ToString() : null;
                        if (pid == null || !pid.Contains(parsed.ValueContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    if (MatchesFormatAttributes(node, parsed))
                        results.Add(node);
                }
            }
            return results;
        }

        // Handle picture queries
        if (elementName == "picture")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart?.WorksheetDrawing == null) continue;

                var picAnchors = EnumeratePictureAnchors(drawingsPart.WorksheetDrawing).ToList();

                for (int i = 0; i < picAnchors.Count; i++)
                {
                    var node = GetPictureNode(sheetName, worksheetPart, i + 1, $"/{sheetName}/picture[{i + 1}]");
                    if (node == null) continue;

                    if (parsed.ValueContains != null)
                    {
                        var alt = node.Format.TryGetValue("alt", out var a) ? a?.ToString() : null;
                        if (alt == null || !alt.Contains(parsed.ValueContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    if (MatchesFormatAttributes(node, parsed))
                        results.Add(node);
                }
            }
            return results;
        }

        // Handle media/image queries
        if (elementName is "media" or "image")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart?.WorksheetDrawing == null) continue;

                var picAnchors = EnumeratePictureAnchors(drawingsPart.WorksheetDrawing).ToList();

                for (int i = 0; i < picAnchors.Count; i++)
                {
                    var node = GetPictureNode(sheetName, worksheetPart, i + 1, $"/{sheetName}/picture[{i + 1}]");
                    if (node == null) continue;

                    // Add content type from image part
                    var pic = picAnchors[i].Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>().First();
                    var blip = pic.BlipFill?.Blip;
                    if (blip?.Embed?.Value != null)
                    {
                        var part = drawingsPart.GetPartById(blip.Embed.Value);
                        if (part != null)
                        {
                            node.Format["contentType"] = part.ContentType;
                            node.Format["fileSize"] = part.GetStream().Length;
                        }
                    }
                    results.Add(node);
                }
            }
            return results;
        }

        // Handle row queries. Symmetric to col/column above: each <row r="N">
        // surfaces as one DocumentNode pointing at /SheetName/row[N]. Without
        // this branch, `query row` fell through to the generic cell loop and
        // returned cell nodes (BUG-BT-R33-2).
        if (elementName is "row")
        {
            // row[<colname|colletter> op val] → match table data rows by a
            // column's cell value (e.g. row[Salary>5000]); distinct from
            // row[hidden=true] which filters row attributes. Column predicates
            // resolve header names against a ListObject — see
            // ExcelHandler.Query.RowWhere.cs.
            // Parse the bracket(s) as one expression tree (single predicate,
            // stacked [a][b], in-bracket `and`, or `or` / parens — uniformly).
            // Classify the leaf predicates into table COLUMNS vs row PROPERTIES
            // (height/hidden/...), then either run the tree-aware row-where over
            // the columns or fall through to the generic post-filter for row
            // properties. e.g. row[金额>5000 or 金额<100], row[Salary>5000][Age<40].
            var rowExpr = AttributeFilter.ParseExpr(selector);
            if (rowExpr != null)
            {
                // `@key` forces ROW PROPERTY; `col.key` forces COLUMN. The '@' is
                // stripped from leaf keys, so recover the forced set from the raw
                // selector to keep it out of the column-shadow collision check.
                var atForced = new HashSet<string>(
                    Regex.Matches(selector, @"@([\w.]+)").Select(m => m.Groups[1].Value),
                    StringComparer.OrdinalIgnoreCase);
                int colCount = 0, attrCount = 0;
                var collisionCandidates = new List<string>();
                foreach (var lc in AttributeFilter.LeafConditions(rowExpr))
                {
                    bool forcedCol = Regex.IsMatch(lc.Key, @"^col(?:umn)?\.", RegexOptions.IgnoreCase);
                    if (!forcedCol && int.TryParse(lc.Key, out _))
                    {
                        var full = $"row[col.{lc.Key}{AttributeFilter.OpToString(lc.Op)}{lc.Value}]";
                        throw new Core.CliException(
                            $"row[{lc.Key} …] is ambiguous: a bare number is the row index, not a column filter. " +
                            $"Use '{full}' to filter by a column named '{lc.Key}', or 'row[{lc.Key}]' (no operator) for that row.")
                            { Code = "invalid_selector" };
                    }
                    if (!forcedCol && atForced.Contains(lc.Key) && !RowAttributeKeys.Contains(lc.Key)
                        && !IsRowSetAttributeKey(lc.Key))
                        // '@' names a row PROPERTY; an unknown one must not fall
                        // through to the generic post-filter, where the absent
                        // key evaluates false and a not(@typo=…) flips to
                        // matching EVERY row (deletion-scale hazard on remove).
                        throw new Core.CliException(
                            $"row[@{lc.Key} …]: '{lc.Key}' is not a row property. " +
                            $"Row properties: {string.Join(", ", RowAttributeKeys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase))}. " +
                            $"For the table column, drop the '@' (row[{lc.Key} …]) or force it with row[col.{lc.Key} …].")
                        {
                            Code = "invalid_selector",
                            ValidValues = RowAttributeKeys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase).ToArray(),
                        };
                    if (!forcedCol && (atForced.Contains(lc.Key) || RowAttributeKeys.Contains(lc.Key)))
                    {
                        attrCount++;
                        // A bare row-property name that ALSO names a real column is
                        // ambiguous (col.<name> / @<name> disambiguate).
                        if (!atForced.Contains(lc.Key)) collisionCandidates.Add(lc.Key);
                    }
                    else colCount++;
                }
                foreach (var ak in collisionCandidates)
                    if (RowKeyCollidesWithColumn(ak, parsed.Sheet))
                        throw new Core.CliException(
                            $"row[{ak} …] is ambiguous: '{ak}' names both a row property and a table column. " +
                            $"Use 'row[col.{ak} …]' to match the column, or 'row[@{ak} …]' to match the row property.")
                        {
                            Code = "invalid_selector",
                            Suggestion = $"row[col.{ak} …] (table column) or row[@{ak} …] (row property)",
                            ValidValues = new[] { $"col.{ak}", $"@{ak}" },
                        };
                if (colCount > 0 && attrCount > 0)
                    throw new Core.CliException(
                        "row[...] cannot mix table columns and row properties in one expression. Split into separate queries.")
                        { Code = "invalid_selector" };
                if (colCount > 0)
                    return QueryRowsByColumnPredicate(parsed.Sheet, rowExpr);
                // pure row properties → fall through to the generic post-filter.
            }

            // A pure-numeric `row[N]` is a positional index, not a filter; it is
            // resolved to that single row by FilterSelectorPositionalIndex (the
            // shared chokepoint that also handles col/shape/chart/...), so this
            // branch returns every row and the wrapper narrows.
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
                if (sheetData == null) continue;

                foreach (var row in sheetData.Elements<Row>())
                {
                    var rowIdx = row.RowIndex?.Value ?? 0u;
                    if (rowIdx == 0) continue;
                    var node = new DocumentNode
                    {
                        Path = $"/{sheetName}/row[{rowIdx}]",
                        Type = "row",
                        ChildCount = row.Elements<Cell>().Count(),
                        Preview = rowIdx.ToString()
                    };
                    // Row properties are emitted only when set (height on 3 of
                    // 1000 rows). Declare the full queryable-key set so the
                    // post-filter treats an absent-but-known key (row[@height>…]
                    // on a sheet where no row has a custom height — the exact
                    // form the collision error recommends) as 0 matches, not
                    // "unknown key".
                    node.InternalFormat["declaredKeys"] = RowAttributeKeys;
                    if (row.Height?.Value != null)
                        node.Format["height"] = $"{row.Height.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}pt";
                    if (row.Hidden?.Value == true) node.Format["hidden"] = true;
                    if (row.CustomHeight?.Value == true) node.Format["customHeight"] = true;
                    if (row.OutlineLevel?.HasValue == true && row.OutlineLevel.Value > 0)
                        node.Format["outlineLevel"] = (int)row.OutlineLevel.Value;
                    if (row.Collapsed?.Value == true) node.Format["collapsed"] = true;
                    if (MatchesFormatAttributes(node, parsed))
                        results.Add(node);
                }
            }
            return results;
        }

        // Handle column queries. OOXML stores columns as <col min=".." max="..">,
        // which can span a range of column indices. We expand spans into one
        // DocumentNode per concrete column so `/SheetName/col[X]` paths align
        // with the Get path format.
        if (elementName is "col" or "column")
        {
            // A positional `col[A]` / `col[2]` is resolved to that single column
            // by FilterSelectorPositionalIndex (shared chokepoint); this branch
            // returns every column and the wrapper narrows.
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var columns = GetSheet(worksheetPart).GetFirstChild<Columns>();
                if (columns == null) continue;

                foreach (var col in columns.Elements<Column>())
                {
                    var min = col.Min?.Value ?? 0u;
                    var max = col.Max?.Value ?? min;
                    if (min == 0) continue;
                    for (uint ci = min; ci <= max; ci++)
                    {
                        var colName = IndexToColumnName((int)ci);
                        var node = new DocumentNode
                        {
                            Path = $"/{sheetName}/col[{colName}]",
                            Type = "column",
                            Preview = colName
                        };
                        if (col.Width?.Value != null) node.Format["width"] = col.Width.Value;
                        if (col.Hidden?.Value == true) node.Format["hidden"] = true;
                        if (col.CustomWidth?.Value == true) node.Format["customWidth"] = true;
                        if (col.OutlineLevel?.HasValue == true && col.OutlineLevel.Value > 0)
                            node.Format["outlineLevel"] = (int)col.OutlineLevel.Value;
                        if (col.Collapsed?.Value == true) node.Format["collapsed"] = true;
                        if (MatchesFormatAttributes(node, parsed))
                            results.Add(node);
                    }
                }
            }
            return results;
        }

        // Handle hyperlink queries. In xlsx, hyperlinks are cell-level metadata
        // (worksheet <hyperlinks><hyperlink ref=".." r:id=".."/></hyperlinks>),
        // not standalone addressable elements. We surface them as discoverable
        // nodes whose Path points at the owning cell so the agent can Get/Set
        // the hyperlink via cell `link` / `tooltip` / `display` props.
        // CONSISTENCY(xlsx-hyperlink-cell-backed): Add/Set live on cells, not here.
        if (elementName is "hyperlink")
        {
            foreach (var (sheetName, worksheetPart) in GetWorksheets())
            {
                if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                    continue;

                var ws = GetSheet(worksheetPart);
                var hyperlinksEl = ws.GetFirstChild<Hyperlinks>();
                if (hyperlinksEl == null) continue;

                foreach (var hl in hyperlinksEl.Elements<Hyperlink>())
                {
                    var cellRef = hl.Reference?.Value ?? "";
                    var node = new DocumentNode
                    {
                        Path = string.IsNullOrEmpty(cellRef) ? $"/{sheetName}" : $"/{sheetName}/{cellRef}",
                        Type = "hyperlink",
                        Preview = cellRef
                    };
                    if (!string.IsNullOrEmpty(cellRef)) node.Format["ref"] = cellRef;
                    // Resolve external URL via relationship id
                    if (hl.Id?.Value != null)
                    {
                        try
                        {
                            var rel = worksheetPart.HyperlinkRelationships
                                .FirstOrDefault(r => r.Id == hl.Id.Value);
                            if (rel != null) node.Format["url"] = rel.Uri.ToString();
                        }
                        catch { }
                    }
                    if (hl.Location?.Value != null) node.Format["location"] = hl.Location.Value;
                    if (hl.Display?.Value != null) node.Format["display"] = hl.Display.Value;
                    if (hl.Tooltip?.Value != null) node.Format["tooltip"] = hl.Tooltip.Value;

                    if (MatchesFormatAttributes(node, parsed))
                        results.Add(node);
                }
            }
            return results;
        }

        // Handle namedrange / definedname queries
        if (elementName is "namedrange" or "definedname")
        {
            var workbook = GetWorkbook();
            var definedNames = workbook.GetFirstChild<DefinedNames>();
            if (definedNames != null)
            {
                var allDefs = definedNames.Elements<DefinedName>().ToList();
                for (int i = 0; i < allDefs.Count; i++)
                {
                    var dn = allDefs[i];
                    var nrNode = new DocumentNode
                    {
                        Path = $"/namedrange[{i + 1}]",
                        Type = "namedrange",
                        Text = dn.InnerText ?? dn.Name?.Value ?? "",
                        Preview = dn.InnerText
                    };
                    if (dn.Name?.Value != null) nrNode.Format["name"] = dn.Name.Value;
                    nrNode.Format["ref"] = dn.InnerText ?? "";
                    if (dn.LocalSheetId?.HasValue == true)
                    {
                        var sheets = workbook.GetFirstChild<Sheets>()?.Elements<Sheet>().ToList();
                        if (sheets != null && (int)dn.LocalSheetId.Value < sheets.Count)
                            nrNode.Format["scope"] = sheets[(int)dn.LocalSheetId.Value].Name?.Value ?? "";
                    }
                    else
                    {
                        // Schema declares scope get=true; emit "workbook" for workbook-scope names.
                        nrNode.Format["scope"] = "workbook";
                    }
                    if (dn.Comment?.HasValue == true) nrNode.Format["comment"] = dn.Comment!.Value!;
                    if (dn.Function?.Value == true) nrNode.Format["volatile"] = true;

                    if (parsed.ValueContains != null)
                    {
                        var name = dn.Name?.Value ?? "";
                        if (!name.Contains(parsed.ValueContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    results.Add(nrNode);
                }
            }
            return results;
        }

        foreach (var (sheetName, worksheetPart) in GetWorksheets())
        {
            // If selector specifies a sheet, skip non-matching sheets
            if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                continue;

            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            var eval = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    if (MatchesCellSelector(cell, sheetName, parsed))
                    {
                        var node = CellToNode(sheetName, cell, worksheetPart, eval);
                        // Carry the stored value under `value` so the CLI
                        // post-filter compares `value>0.3` / `value=0.5` on the
                        // underlying number (0.5), not the display ("50%").
                        // node.Text keeps the display for output; MatchOne also
                        // falls back to it so `value=50%` still matches.
                        node.Format["value"] = GetCellRawComparisonValue(cell, eval);
                        if (MatchesFormatAttributes(node, parsed))
                            results.Add(node);
                    }
                }
            }
        }

        return results;
    }

    // ==================== CF DXF resolution ====================

    /// <summary>
    /// Resolves a conditional formatting rule's dxfId to fill and font colors
    /// from the workbook stylesheet, and populates the DocumentNode accordingly.
    /// </summary>
    private void PopulateCfNodeFromDxf(DocumentNode cfNode, int dxfId)
    {
        var stylesheet = _doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
        if (stylesheet == null) return;

        var dxfs = stylesheet.GetFirstChild<DifferentialFormats>();
        if (dxfs == null) return;

        var dxfList = dxfs.Elements<DifferentialFormat>().ToList();
        if (dxfId < 0 || dxfId >= dxfList.Count) return;

        var dxf = dxfList[dxfId];

        // Resolve fill color
        var fill = dxf.GetFirstChild<Fill>();
        if (fill != null)
        {
            var patternFill = fill.GetFirstChild<PatternFill>();
            if (patternFill != null)
            {
                var bgColor = patternFill.GetFirstChild<BackgroundColor>();
                if (bgColor?.Rgb?.Value != null)
                    cfNode.Format["fill"] = ParseHelpers.FormatHexColor(bgColor.Rgb.Value);
                else
                {
                    var fgColor = patternFill.GetFirstChild<ForegroundColor>();
                    if (fgColor?.Rgb?.Value != null)
                        cfNode.Format["fill"] = ParseHelpers.FormatHexColor(fgColor.Rgb.Value);
                }
            }
        }

        // Resolve font props. The dxf font carries whatever BuildFormulaCfFont
        // wrote (bold/italic/strike/underline/size/name/color); surface them all
        // so `get`/`query` round-trips what `add`/`set` accepts — otherwise an
        // applied `font.bold` reads back as nothing (it IS in the XML).
        var font = dxf.GetFirstChild<Font>();
        if (font != null)
        {
            var bold = font.GetFirstChild<Bold>();
            if (bold != null && (bold.Val == null || bold.Val.Value))
                cfNode.Format["font.bold"] = true;
            var italic = font.GetFirstChild<Italic>();
            if (italic != null && (italic.Val == null || italic.Val.Value))
                cfNode.Format["font.italic"] = true;
            var strike = font.GetFirstChild<Strike>();
            if (strike != null && (strike.Val == null || strike.Val.Value))
                cfNode.Format["font.strike"] = true;
            var underline = font.GetFirstChild<Underline>();
            if (underline != null)
                cfNode.Format["font.underline"] = underline.Val?.InnerText ?? "single";
            var fontSize = font.GetFirstChild<FontSize>();
            if (fontSize?.Val?.Value != null)
                cfNode.Format["font.size"] = $"{fontSize.Val.Value:0.##}pt";
            var fontName = font.GetFirstChild<FontName>();
            if (!string.IsNullOrEmpty(fontName?.Val?.Value))
                cfNode.Format["font.name"] = fontName.Val.Value;
            var fontColor = font.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>();
            if (fontColor?.Rgb?.Value != null)
                cfNode.Format["font.color"] = ParseHelpers.FormatHexColor(fontColor.Rgb.Value);
        }
    }

    /// <summary>
    /// Resolve the x14:cfRule that pairs with a 2007 dataBar rule via x14:id reference,
    /// by scanning the worksheet's extLst x14:conditionalFormattings.
    /// </summary>
    private static X14.ConditionalFormattingRule? FindMatchingX14DataBarRule(
        Worksheet ws,
        ConditionalFormattingRuleExtensionList extList)
    {
        var idExt = extList.Elements<ConditionalFormattingRuleExtension>()
            .FirstOrDefault(e => string.Equals(e.Uri?.Value, "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}", StringComparison.OrdinalIgnoreCase));
        var idEl = idExt?.GetFirstChild<X14.Id>();
        var refId = idEl?.Text;
        if (string.IsNullOrEmpty(refId)) return null;

        const string cfExtUri = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";
        var wsExtList = ws.GetFirstChild<WorksheetExtensionList>();
        if (wsExtList == null) return null;
        foreach (var wsExt in wsExtList.Elements<WorksheetExtension>().Where(e => e.Uri == cfExtUri))
        {
            foreach (var x14Cfs in wsExt.Elements<X14.ConditionalFormattings>())
            foreach (var x14Cf in x14Cfs.Elements<X14.ConditionalFormatting>())
            foreach (var x14Rule in x14Cf.Elements<X14.ConditionalFormattingRule>())
            {
                if (string.Equals(x14Rule.Id?.Value, refId, StringComparison.OrdinalIgnoreCase))
                    return x14Rule;
            }
        }
        return null;
    }

    // CONSISTENCY(xlsx/detected-table): strict, high-precision detection of
    // table-shaped data blocks that are NOT real ListObjects. Design goals
    // (in priority order): (1) whatever is reported is almost certainly a real
    // table — favour false negatives over false positives; (2) fast — relies on
    // Excel's sparse cell storage, so cost is O(non-empty cells), independent of
    // the declared sheet dimension; (3) report every qualifying block on the
    // sheet. Detected blocks carry honest range paths and stable=false.
    //
    // A block qualifies only if ALL hold:
    //   - its top-left anchor has no occupied cell directly above or to the left
    //   - the anchor row is a contiguous run of >= 2 non-empty TEXT cells (the
    //     header) — text = non-empty, not numeric, not a formula
    //   - at least one data row exists below the header within the header span
    //   - it does not overlap any real ListObject range (dedup)
    private List<DocumentNode> DetectTables(
        string sheetName, WorksheetPart worksheetPart,
        List<(int c1, int r1, int c2, int r2)> realRanges)
    {
        var results = new List<DocumentNode>();
        var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
        if (sheetData == null) return results;

        // 1. One sparse pass: occupied map keyed by (col,row) → (value, isText).
        var occupied = new Dictionary<(int col, int row), (string val, bool isText)>();
        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value is not string cref) continue;
                int colIdx, rowNum;
                try { var (cn, rn) = ParseCellReference(cref); colIdx = ColumnNameToIndex(cn); rowNum = rn; }
                catch { continue; }
                var val = GetCellDisplayValue(cell);
                if (string.IsNullOrEmpty(val)) continue;
                bool isText = cell.CellFormula == null
                    && !double.TryParse(val, System.Globalization.NumberStyles.Any,
                                        System.Globalization.CultureInfo.InvariantCulture, out _);
                occupied[(colIdx, rowNum)] = (val, isText);
            }
        }
        if (occupied.Count == 0) return results;

        // 2. Top-left anchors: occupied, nothing above, nothing to the left.
        var anchors = occupied.Keys
            .Where(k => !occupied.ContainsKey((k.col, k.row - 1))
                     && !occupied.ContainsKey((k.col - 1, k.row)))
            .OrderBy(k => k.row).ThenBy(k => k.col)
            .ToList();

        var claimed = new HashSet<(int, int)>();

        foreach (var (ac, ar) in anchors)
        {
            if (claimed.Contains((ac, ar))) continue;

            // 3. Header span: contiguous non-empty cells rightward from a TEXT
            //    anchor. The anchor (table's leading column header) must be text,
            //    but interior/trailing header cells MAY be numeric so a common
            //    year/number header ("Region | 2024 | 2025") is still recognised
            //    — a single numeric header used to truncate the span to one
            //    column and drop the whole table.
            if (!occupied.TryGetValue((ac, ar), out var anchorCell) || !anchorCell.isText) continue;
            int c = ac;
            while (occupied.ContainsKey((c, ar))) c++;
            int headerEndCol = c - 1;
            if (headerEndCol - ac + 1 < 2) continue;  // strict: >= 2 columns

            // 4. Row extent: scan down while any cell within the header span is
            //    occupied. A fully blank row is the block boundary.
            int lastDataRow = ar;
            for (int r = ar + 1; ; r++)
            {
                bool rowHasData = false;
                for (int cc = ac; cc <= headerEndCol; cc++)
                    if (occupied.ContainsKey((cc, r))) { rowHasData = true; break; }
                if (!rowHasData) break;
                lastDataRow = r;
            }
            if (lastDataRow == ar) continue;  // strict: >= 1 data row

            // 5. Dedup against real ListObjects.
            var block = (ac, ar, headerEndCol, lastDataRow);
            if (realRanges.Any(rr => RangesOverlap(rr, block))) continue;

            for (int rr2 = ar; rr2 <= lastDataRow; rr2++)
                for (int cc = ac; cc <= headerEndCol; cc++)
                    claimed.Add((cc, rr2));

            var colNames = new List<string>();
            for (int cc = ac; cc <= headerEndCol; cc++)
                colNames.Add(occupied.TryGetValue((cc, ar), out var hv) ? hv.val : "");

            var startRef = IndexToColumnName(ac) + ar;
            var endRef = IndexToColumnName(headerEndCol) + lastDataRow;
            var rangeRef = $"{startRef}:{endRef}";

            var node = new DocumentNode
            {
                Path = $"/{sheetName}/{rangeRef}",
                Type = "detectedtable",
                Text = colNames.FirstOrDefault() ?? "",
                Preview = rangeRef,
            };
            node.Format["source"] = "header-sniff";
            node.Format["stable"] = false;
            node.Format["ref"] = rangeRef;
            node.Format["columns"] = string.Join(",", colNames);
            // Structured column list — the comma-joined Format["columns"] is
            // lossy when a header itself contains a comma ("Amount, USD"), which
            // silently corrupts header→column resolution downstream. Consumers
            // that resolve a column by name (row-where, set-by-column, hints)
            // read this list instead of re-splitting the string. See
            // DetectedTableColumns().
            node.InternalFormat["columnList"] = colNames;
            node.Format["dataRange"] = $"{IndexToColumnName(ac)}{ar + 1}:{endRef}";
            node.ChildCount = colNames.Count;
            results.Add(node);
        }

        return results;
    }

    private static bool TryParseRange(string? rangeRef, out (int c1, int r1, int c2, int r2) range)
    {
        range = default;
        if (string.IsNullOrEmpty(rangeRef)) return false;
        var parts = rangeRef.Split(':');
        try
        {
            var (c1, r1) = ParseCellReference(parts[0].Replace("$", ""));
            if (parts.Length == 1)
            {
                range = (ColumnNameToIndex(c1), r1, ColumnNameToIndex(c1), r1);
                return true;
            }
            var (c2, r2) = ParseCellReference(parts[1].Replace("$", ""));
            range = (ColumnNameToIndex(c1), r1, ColumnNameToIndex(c2), r2);
            return true;
        }
        catch { return false; }
    }

    private static bool RangesOverlap((int c1, int r1, int c2, int r2) a, (int c1, int r1, int c2, int r2) b)
        => a.c1 <= b.c2 && b.c1 <= a.c2 && a.r1 <= b.r2 && b.r1 <= a.r2;
}
