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
    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        Modified = true;
        // Batch Set: route to the shared filter engine when the path is a bare
        // selector (no `/`) OR a `/`-scoped path that carries a content filter
        // (e.g. `/Sheet1/cell[value>5000 or value<300]`). The latter would
        // otherwise fall to the positional-index navigator and reject the
        // predicate — query already resolves it, so set must too (parity).
        // Structural `/`-paths (`/Sheet1/A1`, `/Sheet1/row[2]`,
        // `/Sheet1/chart[1]/axis[@role=…]`) stay false in IsContentFilterPath
        // and take the direct dispatch below.
        if (!string.IsNullOrEmpty(path)
            && (!path.StartsWith("/") || AttributeFilter.IsContentFilterPath(path)))
        {
            var unsupported = new List<string>();
            // FilterSelector narrows the mutation set with the same engine `query`
            // uses (CommandBuilder.GetQuery.cs): a pure-AND selector takes the
            // legacy path (handler pre-filter + flat re-apply — so `cell[value>100]`
            // no longer over-matches every cell), and an `or` selector is queried
            // with its filter brackets stripped (Query returns broadly) then
            // narrowed by the boolean expression tree. ResolveCellAttributeAlias
            // maps cell short keys (bold -> font.bold, ...); a no-op on other keys.
            var (targets, _) = AttributeFilter.FilterSelector(path, Query, ResolveCellAttributeAlias);
            if (targets.Count == 0)
                // A selector that resolves to zero rows is an ordinary empty
                // WHERE result, not a tool failure — surface not_found so the
                // JSON envelope reads clean instead of internal_error.
                throw new Core.CliException($"No elements matched selector: {path}") { Code = "not_found" };
            LastSelectorSetCount = targets.Count;
            foreach (var target in targets)
            {
                var targetUnsupported = Set(target.Path, properties);
                foreach (var u in targetUnsupported)
                    if (!unsupported.Contains(u)) unsupported.Add(u);
            }
            return unsupported;
        }

        // Normalize to case-insensitive lookup so camelCase keys match lowercase lookups
        if (properties != null && properties.Comparer != StringComparer.OrdinalIgnoreCase)
            properties = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
        properties ??= new Dictionary<string, string>();

        path = NormalizeExcelPath(path);
        path = ResolveSheetIndexInPath(path);

        // Excel only supports find+replace — reject find without replace early (before path dispatch)
        if (properties.ContainsKey("find") && !properties.ContainsKey("replace"))
            throw new ArgumentException("Excel only supports 'find' with 'replace'. Use 'find' + 'replace' for text replacement. find+format (without replace) is not supported in Excel.");
        // CONSISTENCY(find-regex): Excel find/replace now honours the `r"..."`
        // raw-string regex prefix (see ApplyFindReplace), matching docx/pptx.
        // The legacy `regex=true` flag is normalized to the r"..." form so both
        // spellings work, mirroring the WordHandler.Set.cs find path.
        if (properties.TryGetValue("regex", out var xlRegexFlag)
            && ParseHelpers.IsTruthySafe(xlRegexFlag)
            && properties.TryGetValue("find", out var xlFindRaw)
            && !xlFindRaw.StartsWith("r\"") && !xlFindRaw.StartsWith("r'"))
            properties["find"] = $"r\"{xlFindRaw}\"";

        // Handle root path "/" — document properties
        if (path == "/")
        {
            // Find & Replace: special handling before document properties
            if (properties.TryGetValue("find", out var findText) && properties.TryGetValue("replace", out var replaceText))
            {
                var count = FindAndReplace(findText, replaceText, null);
                LastFindMatchCount = count;
                var remaining = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
                remaining.Remove("find");
                remaining.Remove("replace");
                if (remaining.Count > 0)
                    return Set(path, remaining);
                return [];
            }

            var unsupported = new List<string>();
            var pkg = _doc.PackageProperties;
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "title": pkg.Title = value; break;
                    case "author" or "creator": pkg.Creator = value; break;
                    case "subject": pkg.Subject = value; break;
                    case "description": pkg.Description = value; break;
                    case "category": pkg.Category = value; break;
                    case "keywords": pkg.Keywords = value; break;
                    case "lastmodifiedby": pkg.LastModifiedBy = value; break;
                    case "revisionnumber": pkg.Revision = value; break;
                    default:
                        var lowerKey = key.ToLowerInvariant();
                        if (!TrySetWorkbookSetting(lowerKey, value)
                            && !Core.ThemeHandler.TrySetTheme(_doc.WorkbookPart?.ThemePart, lowerKey, value)
                            && !Core.ExtendedPropertiesHandler.TrySetExtendedProperty(
                                Core.ExtendedPropertiesHandler.GetOrCreateExtendedPart(_doc), lowerKey, value))
                            unsupported.Add(key);
                        break;
                }
            }
            return unsupported;
        }

        // Handle /SheetName/sparkline[N]
        var sparklineSetMatch = Regex.Match(path.TrimStart('/'), @"^([^/]+)/sparkline\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (sparklineSetMatch.Success) return SetSparklineByPath(sparklineSetMatch, properties);

        // Handle /namedrange[N] or /namedrange[Name]
        var namedRangeMatch = Regex.Match(path.TrimStart('/'), @"^namedrange\[(.+?)\]$", RegexOptions.IgnoreCase);
        if (namedRangeMatch.Success) return SetNamedRangeByPath(namedRangeMatch, properties);

        // Parse path: /SheetName, /SheetName/A1, /SheetName/A1:D1, /SheetName/col[A], /SheetName/row[1], /SheetName/autofilter
        var segments = path.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];

        var worksheet = FindWorksheet(sheetName);
        if (worksheet == null)
            throw SheetNotFoundException(sheetName);

        // Sheet-level Set (path is just /SheetName)
        if (segments.Length < 2)
        {
            return SetSheetLevel(worksheet, sheetName, properties);
        }

        // BUG-R41-F2: reject cell reference segments that contain control characters
        // (e.g. \n, \r, \t). In .NET, Regex `$` matches before a trailing \n, so
        // without this check "A1\n" would pass ParseCellReference and create a ghost
        // cell with CellReference="A1\n" — an address that never resolves to A1.
        // Reject up-front so the caller gets a clear error instead of silent corruption.
        var cellRef = segments[1];
        if (cellRef.Any(c => c < ' ' && c != '\t' || c == '\x7f'))
            throw new ArgumentException(
                $"Cell reference '{cellRef.Replace("\n", "\\n").Replace("\r", "\\r")}' contains invalid control characters. " +
                $"Expected a clean cell address like 'A1' or 'B2'.");

        // Handle /SheetName/dataValidation[N] (canonical) and
        // /SheetName/validation[N] (legacy alias, R7-bt-6 CONSISTENCY)
        var validationSetMatch = Regex.Match(cellRef, @"^(?:dataValidation|validation)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (validationSetMatch.Success) return SetValidationByPath(validationSetMatch, worksheet, properties);

        // Handle /SheetName/ole[N]
        var oleSetMatch = Regex.Match(cellRef, @"^(?:ole|object|embed)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (oleSetMatch.Success) return SetOleByPath(oleSetMatch, worksheet, properties);

        // Handle /SheetName/picture[N]
        var picSetMatch = Regex.Match(cellRef, @"^picture\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (picSetMatch.Success) return SetPictureByPath(picSetMatch, worksheet, properties);

        // Handle /SheetName/shape[N]
        var shapeSetMatch = Regex.Match(cellRef, @"^shape\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (shapeSetMatch.Success) return SetShapeByPath(shapeSetMatch, worksheet, properties);

        // Handle /SheetName/slicer[N] — caption/style/columnCount/rowHeight/name
        var slicerSetMatch = Regex.Match(cellRef, @"^slicer\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (slicerSetMatch.Success) return SetSlicerByPath(slicerSetMatch, worksheet, properties);

        // Handle /SheetName/table[N]/columns[M] or /SheetName/table[N]/column[M]
        // CONSISTENCY(table-column-path): mirror the col[M].prop= dotted form already
        // accepted on /Sheet/table[N] by exposing the column as a sub-path so users can
        // address it as a node and call Set with a flat property bag.
        var tableColPathMatch = Regex.Match(cellRef,
            @"^table\[(\d+)\]/(?:columns|column)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (tableColPathMatch.Success) return SetTableColumnByPath(tableColPathMatch, worksheet, properties);

        // Handle /SheetName/table[N]
        var tableSetMatch = Regex.Match(cellRef, @"^table\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (tableSetMatch.Success) return SetTableByPath(tableSetMatch, worksheet, properties);

        // Handle /SheetName/comment[N]
        var commentSetMatch = Regex.Match(cellRef, @"^comment\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (commentSetMatch.Success) return SetCommentByPath(commentSetMatch, worksheet, sheetName, properties);

        // Handle /SheetName/autofilter
        if (cellRef.Equals("autofilter", StringComparison.OrdinalIgnoreCase))
        {
            return SetAutoFilter(worksheet, properties);
        }

        // Handle /SheetName/cf[N] or /SheetName/conditionalformatting[N]
        var cfSetMatch = Regex.Match(cellRef, @"^(?:cf|conditionalformatting)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (cfSetMatch.Success) return SetCfByPath(cfSetMatch, worksheet, properties);

        // Handle /SheetName/rowbreak[N] or /SheetName/colbreak[N] — reposition
        // (row/col) and toggle manual. Mirrors the Get readback keys so a queried
        // page break round-trips into set.
        var brkSetMatch = Regex.Match(cellRef, @"^(rowbreak|colbreak)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (brkSetMatch.Success)
            return SetPageBreak(worksheet,
                brkSetMatch.Groups[1].Value.Equals("rowbreak", StringComparison.OrdinalIgnoreCase),
                int.Parse(brkSetMatch.Groups[2].Value), properties);

        // CONSISTENCY(axis-ref-compat): accept Excel-style whole-column/row
        // references (B:B, B:D, 1:1, 2:5) as input aliases — expand to the
        // canonical col[X]/row[N] segments and apply per column/row, the way
        // range set (A1:D1) applies per cell.
        if (TryExpandAxisRef(cellRef) is { } axisSegments)
        {
            var axisUnsupported = new List<string>();
            foreach (var seg in axisSegments)
            {
                var segResult = Set($"/{sheetName}/{seg}", properties);
                // Intersect: a prop is unsupported only if no column/row took it.
                axisUnsupported = axisSegments[0] == seg
                    ? segResult
                    : axisUnsupported.Intersect(segResult, StringComparer.OrdinalIgnoreCase).ToList();
            }
            return axisUnsupported;
        }

        // Handle /SheetName/col[X] where X is a column letter (A) or numeric index (1)
        var colMatch = Regex.Match(cellRef, @"^col\[([A-Za-z0-9]+)\]$", RegexOptions.IgnoreCase);
        if (colMatch.Success)
        {
            var colValue = colMatch.Groups[1].Value;
            var colName = int.TryParse(colValue, out var colNumIdx) ? IndexToColumnName(colNumIdx) : colValue.ToUpperInvariant();
            return SetColumn(worksheet, colName, properties);
        }

        // Handle /SheetName/row[N]
        var rowMatch = Regex.Match(cellRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = uint.Parse(rowMatch.Groups[1].Value);
            return SetRow(worksheet, rowIdx, properties);
        }

        // Handle /SheetName/chart[N]/axis[@role=ROLE]
        var chartAxisSetMatch = Regex.Match(cellRef,
            @"^chart\[(\d+)\]/axis\[@role=([a-zA-Z0-9_]+)\]$");
        if (chartAxisSetMatch.Success) return SetChartAxisByPath(chartAxisSetMatch, worksheet, properties);

        // Handle /SheetName/chart[N] or /SheetName/chart[N]/series[K]
        var chartMatch = Regex.Match(cellRef, @"^chart\[(\d+)\](?:/series\[(\d+)\])?$");
        if (chartMatch.Success) return SetChartByPath(chartMatch, worksheet, properties);

        // Handle /SheetName/pivottable[N]
        var pivotSetMatch = Regex.Match(cellRef, @"^pivottable\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (pivotSetMatch.Success) return SetPivotTableByPath(pivotSetMatch, worksheet, properties);

        // Handle /SheetName/A1/run[N] (rich text run)
        var runSetMatch = Regex.Match(cellRef, @"^([A-Z]+\d+)/run\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (runSetMatch.Success) return SetCellRunByPath(runSetMatch, worksheet, properties);

        // Handle /SheetName/A1:D1 (range — merge/unmerge)
        if (cellRef.Contains(':'))
        {
            var firstPartRange = cellRef.Split(':')[0];
            bool isRangeRef = Regex.IsMatch(firstPartRange, @"^[A-Z]+\d+$", RegexOptions.IgnoreCase);
            if (isRangeRef)
            {
                return SetRange(worksheet, cellRef.ToUpperInvariant(), properties);
            }
        }

        // Check if path is a cell reference or generic XML path
        var firstPart = cellRef.Split('/')[0].Split('[')[0];
        bool isCellRef = Regex.IsMatch(firstPart, @"^[A-Z]+\d+", RegexOptions.IgnoreCase);
        if (!isCellRef)
        {
            // Generic XML fallback: navigate to element and set attributes
            var xmlSegments = GenericXmlQuery.ParsePathSegments(cellRef);
            var target = GenericXmlQuery.NavigateByPath(GetSheet(worksheet), xmlSegments);
            if (target == null)
            {
                // Common selector misspellings: "column A" / "col/A" /
                // "row 5" / "row/5" are the space and slash forms of the
                // bracket selectors col[A] / row[5] — point at the exact
                // form instead of a bare not_found (Issue #351 family).
                var selMatch = Regex.Match(cellRef,
                    @"^(?:column|col)[\s/]+([A-Za-z]{1,3})$|^(?:row)[\s/]+(\d{1,7})$",
                    RegexOptions.IgnoreCase);
                if (selMatch.Success)
                {
                    var hint = selMatch.Groups[1].Success
                        ? $"col[{selMatch.Groups[1].Value.ToUpperInvariant()}]"
                        : $"row[{selMatch.Groups[2].Value}]";
                    throw new ArgumentException(
                        $"Element not found: {cellRef}. Did you mean '{hint}'? Columns and rows use the bracket form (col[A], row[1]) on the target sheet.");
                }
                throw new ArgumentException($"Element not found: {cellRef}");
            }
            var unsup = new List<string>();
            foreach (var (key, value) in properties)
            {
                if (!GenericXmlQuery.SetGenericAttribute(target, key, value))
                    unsup.Add(key);
            }
            SaveWorksheet(worksheet);
            return unsup;
        }

        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            GetSheet(worksheet).Append(sheetData);
        }

        // Did the cell exist before this Set? If not, FindOrCreateCell
        // materializes it, and a mid-Set validation throw must remove it
        // entirely — restoring an empty clone (below) would leave a ghost
        // <c> stub behind on a failed create (exit 1 must mean no change).
        var cellPreExisted = sheetData.Elements<Row>()
            .SelectMany(r => r.Elements<Cell>())
            .Any(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));

        var cell = FindOrCreateCell(sheetData, cellRef);

        // Clone cell for rollback on failure (atomic: no partial modifications)
        var cellBackup = cell.CloneNode(true);

        try
        {
        return SetCellProperties(cell, cellRef, worksheet, properties);
        }
        catch
        {
            if (cellPreExisted)
                // Rollback: restore cell to pre-modification state.
                cell.Parent?.ReplaceChild(cellBackup, cell);
            else
                // Newly created by this Set — remove it so a failed create
                // leaves no ghost cell.
                cell.Remove();
            throw;
        }
    }
}
