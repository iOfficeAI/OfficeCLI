// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    public string ViewAsText(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var sb = new StringBuilder();
        var sheets = GetWorksheets();
        int sheetIdx = 0;
        int emitted = 0;
        bool truncated = false;

        foreach (var (sheetName, worksheetPart) in sheets)
        {
            if (truncated) break;
            sb.AppendLine($"=== Sheet: {sheetName} ===");
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            int totalRows = sheetData.Elements<Row>().Count();
            var evaluator = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);
            int lineNum = 0;
            foreach (var row in sheetData.Elements<Row>())
            {
                lineNum++;
                if (startLine.HasValue && lineNum < startLine.Value) continue;
                if (endLine.HasValue && lineNum > endLine.Value) break;

                if (maxLines.HasValue && emitted >= maxLines.Value)
                {
                    sb.AppendLine($"... (showed {emitted} rows, {totalRows} total in sheet, use --start/--end to view more)");
                    truncated = true;
                    break;
                }

                var cellElements = row.Elements<Cell>();
                if (cols != null)
                    cellElements = cellElements.Where(c => cols.Contains(ParseCellReference(c.CellReference?.Value ?? "A1").Column));
                var cells = cellElements.Select(c => GetCellDisplayValue(c, evaluator)).ToArray();
                var rowRef = row.RowIndex?.Value ?? (uint)lineNum;
                sb.AppendLine($"[/{sheetName}/row[{rowRef}]] {string.Join("\t", cells)}");
                emitted++;
            }

            sheetIdx++;
            if (sheetIdx < sheets.Count) sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsAnnotated(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var sb = new StringBuilder();
        var sheets = GetWorksheets();
        int emitted = 0;
        bool truncated = false;

        foreach (var (sheetName, worksheetPart) in sheets)
        {
            if (truncated) break;
            sb.AppendLine($"=== Sheet: {sheetName} ===");
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            int totalRows = sheetData.Elements<Row>().Count();
            int lineNum = 0;
            foreach (var row in sheetData.Elements<Row>())
            {
                lineNum++;
                if (startLine.HasValue && lineNum < startLine.Value) continue;
                if (endLine.HasValue && lineNum > endLine.Value) break;

                if (maxLines.HasValue && emitted >= maxLines.Value)
                {
                    sb.AppendLine($"... (showed {emitted} rows, {totalRows} total in sheet, use --start/--end to view more)");
                    truncated = true;
                    break;
                }

                var cellElements = row.Elements<Cell>();
                if (cols != null)
                    cellElements = cellElements.Where(c => cols.Contains(ParseCellReference(c.CellReference?.Value ?? "A1").Column));

                foreach (var cell in cellElements)
                {
                    var cellRef = cell.CellReference?.Value ?? "?";
                    var value = GetCellDisplayValue(cell);
                    var formula = cell.CellFormula?.Text;
                    var type = cell.DataType?.InnerText ?? "Number";

                    var annotation = formula != null ? $"={formula}" : type;
                    var warn = "";

                    if (string.IsNullOrEmpty(value) && formula == null)
                        warn = " \u26a0 empty";
                    else if (formula != null && (value == "#REF!" || value == "#VALUE!" || value == "#NAME?"))
                        warn = " \u26a0 formula error";

                    sb.AppendLine($"  {cellRef}: [{value}] \u2190 {annotation}{warn}");
                }
                emitted++;
            }
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsOutline()
    {
        var sb = new StringBuilder();
        var workbook = _doc.WorkbookPart?.Workbook;
        if (workbook == null) return "(empty workbook)";

        var sheets = workbook.GetFirstChild<Sheets>();
        if (sheets == null) return "(no sheets)";

        sb.AppendLine($"File: {Path.GetFileName(_filePath)}");

        foreach (var sheet in sheets.Elements<Sheet>())
        {
            var name = sheet.Name?.Value ?? "?";
            var sheetId = sheet.Id?.Value;
            if (sheetId == null) continue;

            var worksheetPart = (WorksheetPart)_doc.WorkbookPart!.GetPartById(sheetId);
            var worksheet = GetSheet(worksheetPart);
            var sheetData = worksheet.GetFirstChild<SheetData>();

            int rowCount = sheetData?.Elements<Row>().Count() ?? 0;
            int colCount = GetSheetColumnCount(worksheet, sheetData);

            int formulaCount = 0;
            if (sheetData != null)
            {
                formulaCount = sheetData.Descendants<CellFormula>().Count();
            }

            var formulaInfo = formulaCount > 0 ? $", {formulaCount} formula(s)" : "";

            // Pivot tables are stored as pivotTableDefinition XML; their rendered cells
            // are NOT materialized into sheetData (Excel/Calc re-render from pivotCacheRecords
            // at display time). Without this hint, a pivot-only sheet looks like "0 rows × 0 cols"
            // and users think it's empty. Surface the pivot count explicitly — same strategy POI
            // takes via XSSFSheet.getPivotTables(). See also: query pivottable.
            int pivotCount = worksheetPart.PivotTableParts.Count();
            var pivotInfo = pivotCount > 0 ? $", {pivotCount} pivot table(s)" : "";

            int oleCount = CountSheetOleObjects(worksheetPart);
            var oleInfo = oleCount > 0 ? $", {oleCount} ole object(s)" : "";

            sb.AppendLine($"\u251c\u2500\u2500 \"{name}\" ({rowCount} rows \u00d7 {colCount} cols{formulaInfo}{pivotInfo}{oleInfo})");
        }

        return sb.ToString().TrimEnd();
    }

    // CONSISTENCY(ole-stats): per-sheet OLE counter shared by outline and
    // outlineJson. Same dedup rule as ViewAsStats — referenced oleObject
    // elements count once, orphan embedded/package parts add extras.
    private int CountSheetOleObjects(WorksheetPart worksheetPart)
    {
        int count = 0;
        var referenced = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var oleEl in GetSheet(worksheetPart).Descendants<OleObject>())
        {
            count++;
            if (oleEl.Id?.Value is string rid && !string.IsNullOrEmpty(rid))
                referenced.Add(rid);
        }
        count += worksheetPart.EmbeddedObjectParts.Count(p => !referenced.Contains(worksheetPart.GetIdOfPart(p)));
        count += worksheetPart.EmbeddedPackageParts.Count(p => !referenced.Contains(worksheetPart.GetIdOfPart(p)));
        return count;
    }

    public string ViewAsStats()
    {
        var sb = new StringBuilder();
        var sheets = GetWorksheets();
        int totalCells = 0;
        int emptyCells = 0;
        int formulaCells = 0;
        int errorCells = 0;
        var typeCounts = new Dictionary<string, int>();

        foreach (var (sheetName, worksheetPart) in sheets)
        {
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    totalCells++;
                    var value = GetCellDisplayValue(cell);
                    if (string.IsNullOrEmpty(value)) emptyCells++;
                    if (cell.CellFormula != null) formulaCells++;
                    if (value is "#REF!" or "#VALUE!" or "#NAME?" or "#DIV/0!") errorCells++;

                    var type = cell.DataType?.InnerText ?? "Number";
                    typeCounts[type] = typeCounts.GetValueOrDefault(type) + 1;
                }
            }
        }

        // OLE object count across all sheets. Same dedup rule as
        // CollectOleNodesForSheet: referenced parts count as one entry
        // (via their oleObject element), orphan parts add extras.
        int oleCount = 0;
        foreach (var (_, worksheetPart) in sheets)
        {
            var referenced = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var oleEl in GetSheet(worksheetPart).Descendants<OleObject>())
            {
                oleCount++;
                if (oleEl.Id?.Value is string rid && !string.IsNullOrEmpty(rid))
                    referenced.Add(rid);
            }
            oleCount += worksheetPart.EmbeddedObjectParts.Count(p => !referenced.Contains(worksheetPart.GetIdOfPart(p)));
            oleCount += worksheetPart.EmbeddedPackageParts.Count(p => !referenced.Contains(worksheetPart.GetIdOfPart(p)));
        }

        sb.AppendLine($"Sheets: {sheets.Count}");
        sb.AppendLine($"Total Cells: {totalCells}");
        sb.AppendLine($"Empty Cells: {emptyCells}");
        sb.AppendLine($"Formula Cells: {formulaCells}");
        sb.AppendLine($"Error Cells: {errorCells}");
        if (oleCount > 0) sb.AppendLine($"OLE Objects: {oleCount}");
        sb.AppendLine();
        sb.AppendLine("Data Type Distribution:");
        foreach (var (type, count) in typeCounts.OrderByDescending(kv => kv.Value))
            sb.AppendLine($"  {type}: {count}");

        return sb.ToString().TrimEnd();
    }

    public JsonNode ViewAsStatsJson()
    {
        var sheets = GetWorksheets();
        int totalCells = 0, emptyCells = 0, formulaCells = 0, errorCells = 0;
        var typeCounts = new Dictionary<string, int>();

        foreach (var (sheetName, worksheetPart) in sheets)
        {
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>())
                foreach (var cell in row.Elements<Cell>())
                {
                    totalCells++;
                    var value = GetCellDisplayValue(cell);
                    if (string.IsNullOrEmpty(value)) emptyCells++;
                    if (cell.CellFormula != null) formulaCells++;
                    if (value is "#REF!" or "#VALUE!" or "#NAME?" or "#DIV/0!") errorCells++;
                    var type = cell.DataType?.InnerText ?? "Number";
                    typeCounts[type] = typeCounts.GetValueOrDefault(type) + 1;
                }
        }

        int oleCountJson = 0;
        foreach (var (_, worksheetPart) in sheets)
        {
            var refSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var oleEl in GetSheet(worksheetPart).Descendants<OleObject>())
            {
                oleCountJson++;
                if (oleEl.Id?.Value is string rid && !string.IsNullOrEmpty(rid))
                    refSet.Add(rid);
            }
            oleCountJson += worksheetPart.EmbeddedObjectParts.Count(p => !refSet.Contains(worksheetPart.GetIdOfPart(p)));
            oleCountJson += worksheetPart.EmbeddedPackageParts.Count(p => !refSet.Contains(worksheetPart.GetIdOfPart(p)));
        }

        var result = new JsonObject
        {
            ["sheets"] = sheets.Count,
            ["totalCells"] = totalCells,
            ["emptyCells"] = emptyCells,
            ["formulaCells"] = formulaCells,
            ["errorCells"] = errorCells,
            ["oleObjects"] = oleCountJson,
        };

        var types = new JsonObject();
        foreach (var (type, count) in typeCounts.OrderByDescending(kv => kv.Value))
            types[type] = count;
        result["dataTypeDistribution"] = types;

        return result;
    }

    public JsonNode ViewAsOutlineJson()
    {
        var workbook = _doc.WorkbookPart?.Workbook;
        if (workbook == null) return new JsonObject();

        var sheetsEl = workbook.GetFirstChild<Sheets>();
        if (sheetsEl == null) return new JsonObject { ["fileName"] = Path.GetFileName(_filePath), ["sheets"] = new JsonArray() };

        var sheetsArray = new JsonArray();
        foreach (var sheet in sheetsEl.Elements<Sheet>())
        {
            var name = sheet.Name?.Value ?? "?";
            var sheetId = sheet.Id?.Value;
            if (sheetId == null) continue;

            var worksheetPart = (WorksheetPart)_doc.WorkbookPart!.GetPartById(sheetId);
            var worksheet = GetSheet(worksheetPart);
            var sheetData = worksheet.GetFirstChild<SheetData>();
            int rowCount = sheetData?.Elements<Row>().Count() ?? 0;
            int colCount = GetSheetColumnCount(worksheet, sheetData);
            int formulaCount = sheetData?.Descendants<CellFormula>().Count() ?? 0;

            int oleCount = CountSheetOleObjects(worksheetPart);
            var sheetObj = new JsonObject
            {
                ["name"] = name,
                ["rows"] = rowCount,
                ["cols"] = colCount,
                ["formulas"] = formulaCount,
                ["oleObjects"] = oleCount
            };
            sheetsArray.Add((JsonNode)sheetObj);
        }

        return new JsonObject
        {
            ["fileName"] = Path.GetFileName(_filePath),
            ["sheets"] = sheetsArray
        };
    }

    public JsonNode ViewAsTextJson(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var sheetsArray = new JsonArray();
        var worksheets = GetWorksheets();
        int emitted = 0;
        bool truncated = false;

        foreach (var (sheetName, worksheetPart) in worksheets)
        {
            if (truncated) break;
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            var rowsArray = new JsonArray();
            int lineNum = 0;
            foreach (var row in sheetData.Elements<Row>())
            {
                lineNum++;
                if (startLine.HasValue && lineNum < startLine.Value) continue;
                if (endLine.HasValue && lineNum > endLine.Value) break;
                if (maxLines.HasValue && emitted >= maxLines.Value) { truncated = true; break; }

                var cellElements = row.Elements<Cell>();
                if (cols != null)
                    cellElements = cellElements.Where(c => cols.Contains(ParseCellReference(c.CellReference?.Value ?? "A1").Column));

                var cellsObj = new JsonObject();
                foreach (var cell in cellElements)
                {
                    var cellRef = cell.CellReference?.Value ?? "?";
                    cellsObj[cellRef] = GetCellDisplayValue(cell);
                }

                var rowRef = row.RowIndex?.Value ?? (uint)lineNum;
                rowsArray.Add((JsonNode)new JsonObject
                {
                    ["row"] = (int)rowRef,
                    ["cells"] = cellsObj
                });
                emitted++;
            }

            sheetsArray.Add((JsonNode)new JsonObject
            {
                ["name"] = sheetName,
                ["rows"] = rowsArray
            });
        }

        return new JsonObject { ["sheets"] = sheetsArray };
    }

    private static int GetSheetColumnCount(Worksheet worksheet, SheetData? sheetData)
    {
        // Try SheetDimension first (e.g., <dimension ref="A1:F20"/>)
        var dimRef = worksheet.GetFirstChild<SheetDimension>()?.Reference?.Value;
        if (!string.IsNullOrEmpty(dimRef))
        {
            var parts = dimRef.Split(':');
            if (parts.Length == 2)
            {
                var endRef = parts[1];
                var col = new string(endRef.TakeWhile(char.IsLetter).ToArray());
                if (!string.IsNullOrEmpty(col))
                    return ColumnNameToIndex(col);
            }
            // Single-cell dimension like "A1" means 1 column
            if (parts.Length == 1)
            {
                var col = new string(parts[0].TakeWhile(char.IsLetter).ToArray());
                if (!string.IsNullOrEmpty(col))
                    return ColumnNameToIndex(col);
            }
        }

        // Fallback: scan all rows for max cell count
        if (sheetData == null) return 0;
        int maxCols = 0;
        foreach (var row in sheetData.Elements<Row>())
        {
            var count = row.Elements<Cell>().Count();
            if (count > maxCols) maxCols = count;
        }
        return maxCols;
    }

    public List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null)
    {
        var issues = new List<DocumentIssue>();
        int issueNum = 0;

        var sheets = GetWorksheets();
        foreach (var (sheetName, worksheetPart) in sheets)
        {
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var cellRef = cell.CellReference?.Value ?? "?";
                    var value = GetCellDisplayValue(cell);

                    if (cell.CellFormula != null && value is "#REF!" or "#VALUE!" or "#NAME?" or "#DIV/0!")
                    {
                        issues.Add(new DocumentIssue
                        {
                            Id = $"F{++issueNum}",
                            Type = IssueType.Content,
                            Severity = IssueSeverity.Error,
                            Path = $"{sheetName}!{cellRef}",
                            Message = $"Formula error: {value}",
                            Context = $"={cell.CellFormula.Text}"
                        });
                    }

                    if (limit.HasValue && issues.Count >= limit.Value) break;
                }
                if (limit.HasValue && issues.Count >= limit.Value) break;
            }
        }

        return issues;
    }
}
