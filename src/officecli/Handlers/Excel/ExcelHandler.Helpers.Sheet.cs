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

    /// <summary>
    /// Validate a sheet name against Excel's rules. Throws ArgumentException
    /// with a clear message on the first rule violation. Rules:
    ///   - non-empty, non-whitespace
    ///   - max 31 chars
    ///   - cannot contain  \  /  ?  *  :  [  ]
    ///   - cannot start or end with apostrophe '
    ///   - cannot equal reserved "History" (case-insensitive)
    /// </summary>
    /// <summary>
    /// Insert a fresh SheetProtection element in schema-correct position.
    /// CT_Worksheet order requires sheetProtection before autoFilter, sortState,
    /// dataConsolidate, customSheetViews, mergeCells, phoneticPr,
    /// conditionalFormatting, dataValidations, hyperlinks, printOptions,
    /// pageMargins, pageSetup, headerFooter, rowBreaks, colBreaks, customProperties,
    /// cellWatches, ignoredErrors, smartTags, drawing, legacyDrawing,
    /// legacyDrawingHF, drawingHF, picture, oleObjects, controls, webPublishItems,
    /// tableParts, extLst. Excel rejects out-of-order placements.
    /// </summary>
    internal static void InsertSheetProtectionInOrder(Worksheet ws, SheetProtection sp)
    {
        OpenXmlElement? anchor = null;
        foreach (var child in ws.ChildElements)
        {
            if (child is AutoFilter
                || child is SortState
                || child is DataConsolidate
                || child is CustomSheetViews
                || child is MergeCells
                || child is PhoneticProperties
                || child is ConditionalFormatting
                || child is DataValidations
                || child is Hyperlinks
                || child is PrintOptions
                || child is PageMargins
                || child is PageSetup
                || child is HeaderFooter
                || child is RowBreaks
                || child is ColumnBreaks
                || child is CustomProperties
                || child is CellWatches
                || child is IgnoredErrors
                || child is DocumentFormat.OpenXml.Spreadsheet.Drawing
                || child is LegacyDrawing
                || child is LegacyDrawingHeaderFooter
                || child is Picture
                || child is OleObjects
                || child is Controls
                || child is WebPublishItems
                || child is TableParts
                || child is WorksheetExtensionList)
            {
                anchor = child;
                break;
            }
        }
        if (anchor != null)
            ws.InsertBefore(sp, anchor);
        else
            ws.AppendChild(sp);
    }

    // ==================== Path Normalization ====================

    /// <summary>
    /// Normalize Excel-native path notation to DOM style.
    /// Sheet1!A1 → /Sheet1/A1
    /// Sheet1!A1:D10 → /Sheet1/A1:D10
    /// Sheet1!row[2] → /Sheet1/row[2]
    /// Sheet1!1:1 → /Sheet1/row[1]   (whole row)
    /// Sheet1!A:A → /Sheet1/col[A]   (whole column)
    /// Paths already starting with '/' are returned unchanged.
    /// </summary>
    // Excel-quoted sheet name: 'My Data (2024)'!A1 — strip one pair of single
    // quotes and un-double embedded quotes ('It''s'!A1 → It's). Excel REQUIRES
    // the quoted form for names with spaces/punctuation, so refs pasted from
    // formulas arrive quoted. Unquoted names pass through unchanged.
    internal static string UnquoteSheetName(string s)
        => s.Length >= 2 && s[0] == '\'' && s[^1] == '\''
            ? s[1..^1].Replace("''", "'")
            : s;

    internal string NormalizeExcelPath(string path)
    {
        // Reject malformed segment separators that previously slipped past
        // the regex matchers and exposed raw OOXML local names. DOCX already
        // rejects these; bring XLSX up to parity.
        if (path.Length > 1 && path != "/" && path.EndsWith("/"))
            throw new ArgumentException($"Invalid path '{path}': trailing '/' is not allowed.");
        if (path.StartsWith("//"))
            throw new ArgumentException($"Invalid path '{path}': leading '//' is not allowed.");
        if (path.Contains("//"))
            throw new ArgumentException($"Invalid path '{path}': empty path segment ('//') is not allowed.");
        // Handle "/Sheet1!A1" — strip leading '/' when '!' is present so native
        // notation is parsed correctly. EXCEPT when the first slash segment
        // names an existing sheet: '!' is legal inside a sheet NAME (Excel
        // forbids only : \ / ? * [ ]), and reinterpreting "/Q1!Results" as
        // bang notation made such a sheet permanently unaddressable.
        if (path.StartsWith('/') && path.Contains('!'))
        {
            var seg0 = path[1..];
            var seg0Slash = seg0.IndexOf('/');
            var first = seg0Slash < 0 ? seg0 : seg0[..seg0Slash];
            if (!GetWorksheets().Any(w => w.Name.Equals(first, StringComparison.OrdinalIgnoreCase)))
                path = path[1..];
        }
        if (path.Equals("/workbook", StringComparison.OrdinalIgnoreCase)) return "/";
        if (path.StartsWith('/')) return path;
        // Excel-quoted sheet ref: 'My Data'!A1 — the '!' separator is the one
        // FOLLOWING the closing quote (the name itself may contain '!').
        string? qSheet = null; var rest = "";
        if (path.StartsWith('\''))
        {
            int qi = 1;
            while (qi < path.Length)
            {
                if (path[qi] == '\'')
                {
                    if (qi + 1 < path.Length && path[qi + 1] == '\'') { qi += 2; continue; }
                    break;
                }
                qi++;
            }
            if (qi < path.Length - 1 && path[qi + 1] == '!')
            {
                qSheet = UnquoteSheetName(path[..(qi + 1)]);
                rest = path[(qi + 2)..];
            }
        }
        var bang = qSheet != null ? 0 : path.IndexOf('!');
        if (qSheet != null || bang > 0)
        {
            var sheet = qSheet ?? path[..bang];
            var selector = qSheet != null ? rest : path[(bang + 1)..];

            // Whole-row notation: "1:1" or "3:3"
            var wholeRow = System.Text.RegularExpressions.Regex.Match(selector, @"^(\d+):\1$");
            if (wholeRow.Success)
                return $"/{sheet}/row[{wholeRow.Groups[1].Value}]";

            // Whole-column notation: "A:A" or "AB:AB"
            var wholeCol = System.Text.RegularExpressions.Regex.Match(selector, @"^([A-Za-z]+):\1$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (wholeCol.Success)
                return $"/{sheet}/col[{wholeCol.Groups[1].Value.ToUpperInvariant()}]";

            return $"/{sheet}/{selector}";
        }
        return path;
    }

    /// <summary>
    /// Resolve sheet[N] index references in the first segment of a normalized path.
    /// E.g. /sheet[1]/A1 → /Sheet1/A1 (if the first sheet is named "Sheet1").
    /// Must be called after NormalizeExcelPath.
    /// </summary>
    private string ResolveSheetIndexInPath(string path)
    {
        if (!path.StartsWith('/')) return path;
        var trimmed = path[1..]; // remove leading '/'
        var slashIdx = trimmed.IndexOf('/');
        var firstSegment = slashIdx >= 0 ? trimmed[..slashIdx] : trimmed;
        var resolved = ResolveSheetName(firstSegment);
        if (resolved == firstSegment) return path;
        return slashIdx >= 0 ? $"/{resolved}/{trimmed[(slashIdx + 1)..]}" : $"/{resolved}";
    }

    // ==================== Private Helpers ====================

    private static Worksheet GetSheet(WorksheetPart part) =>
        part.Worksheet ?? throw new InvalidOperationException("Corrupt file: worksheet data missing");

    /// <summary>
    /// Mark a worksheet as dirty. The actual save (with schema-order reorder) is
    /// deferred to <see cref="FlushDirtyParts"/> which runs in Dispose().
    /// This replaces per-mutation Save() calls — batch operations over many cells
    /// previously triggered one disk write per cell (O(n) saves); now they all
    /// flush in a single pass at the end.
    /// </summary>
    private void SaveWorksheet(WorksheetPart part)
    {
        // A dirty worksheet is itself proof that the editable package was
        // modified.  Keep this invariant here instead of relying on every
        // caller to set Modified separately: Import writes rows directly and
        // used to enqueue the part without setting Modified, so Dispose()
        // treated the session as read-only and discarded the in-memory copy.
        Modified = true;
        _dirtyWorksheets.Add(part);
    }

    /// <summary>
    /// Flush all pending worksheet and stylesheet saves. Called from Dispose().
    /// Each dirty WorksheetPart is reordered and saved exactly once regardless
    /// of how many mutations targeted it.
    /// </summary>
    private void FlushDirtyParts()
    {
        // Reconcile any formula caches that went stale since they were written
        // (e.g. a SUMIFS authored before its data was imported). Runs before the
        // dirty-part flush so cells it touches are picked up by the loop below.
        RefreshStaleFormulaCaches();
        // Re-seed chart numCache/strCache from current cell values so offline
        // consumers (dump/batch, view html) don't read stale series data and
        // dump→replay stays idempotent. Runs after the formula sweep so charts
        // referencing formula cells see the reconciled values.
        RefreshStaleChartCaches();
        foreach (var part in _dirtyWorksheets)
        {
            ReorderWorksheetChildren(GetSheet(part));
            GetSheet(part).Save();
        }
        _dirtyWorksheets.Clear();
        if (_dirtyStylesheet)
        {
            _doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet?.Save();
            _dirtyStylesheet = false;
        }
    }

    /// <summary>
    /// Delete the calculation chain part if present.
    /// Excel will recalculate and recreate it on next open.
    /// This avoids stale calc chain references after cell/formula mutations.
    /// </summary>
    private void DeleteCalcChainIfPresent()
    {
        var calcChainPart = _doc.WorkbookPart?.CalculationChainPart;
        if (calcChainPart != null)
            _doc.WorkbookPart!.DeletePart(calcChainPart);
    }

    /// <summary>
    /// Reorder worksheet children to match OpenXML schema sequence.
    /// Schema: sheetPr, dimension, sheetViews, sheetFormatPr, cols, sheetData,
    ///   autoFilter, sortState, mergeCells, conditionalFormatting,
    ///   dataValidations, hyperlinks, printOptions, pageMargins, pageSetup,
    ///   headerFooter, drawing, legacyDrawing, tableParts, extLst
    /// </summary>
    private static void ReorderWorksheetChildren(Worksheet ws)
    {
        var order = new Dictionary<string, int>
        {
            ["sheetPr"] = 0, ["dimension"] = 1, ["sheetViews"] = 2, ["sheetFormatPr"] = 3,
            ["cols"] = 4, ["sheetData"] = 5, ["sheetCalcPr"] = 6, ["sheetProtection"] = 7,
            ["protectedRanges"] = 8, ["scenarios"] = 9, ["autoFilter"] = 10, ["sortState"] = 11,
            ["dataConsolidate"] = 12, ["customSheetViews"] = 13, ["mergeCells"] = 14,
            ["phoneticPr"] = 15, ["conditionalFormatting"] = 16, ["dataValidations"] = 17,
            ["hyperlinks"] = 18, ["printOptions"] = 19, ["pageMargins"] = 20,
            ["pageSetup"] = 21, ["headerFooter"] = 22, ["rowBreaks"] = 23, ["colBreaks"] = 24,
            ["drawing"] = 25, ["legacyDrawing"] = 26, ["tableParts"] = 27, ["extLst"] = 99
        };

        var children = ws.ChildElements.ToList();
        var sorted = children
            .OrderBy(c => order.TryGetValue(c.LocalName, out var idx) ? idx : 50)
            .ToList();

        bool needsReorder = false;
        for (int i = 0; i < children.Count; i++)
        {
            if (!ReferenceEquals(children[i], sorted[i]))
            {
                needsReorder = true;
                break;
            }
        }

        if (needsReorder)
        {
            foreach (var child in children) child.Remove();
            foreach (var child in sorted) ws.AppendChild(child);
        }
    }

    private Workbook GetWorkbook() =>
        _doc.WorkbookPart?.Workbook ?? throw new InvalidOperationException("Corrupt file: workbook missing");

    private List<(string Name, WorksheetPart Part)> GetWorksheets() => GetWorksheets(_doc);

    /// <summary>
    /// True when the workbook-level &lt;sheet&gt; element for the given name has
    /// State Hidden or VeryHidden. Used by the HTML preview to omit hidden sheets
    /// from the tab strip and content slider (matching real Excel).
    /// </summary>
    private bool IsSheetHidden(string sheetName)
    {
        var wbSheet = GetWorkbook().GetFirstChild<Sheets>()?.Elements<Sheet>()
            .FirstOrDefault(s => s.Name?.Value?.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true);
        return wbSheet?.State?.Value is { } st
            && (st == SheetStateValues.Hidden || st == SheetStateValues.VeryHidden);
    }

    private static List<(string Name, WorksheetPart Part)> GetWorksheets(SpreadsheetDocument doc)
    {
        var result = new List<(string, WorksheetPart)>();
        var workbook = doc.WorkbookPart?.Workbook;
        if (workbook == null) return result;

        var sheets = workbook.GetFirstChild<Sheets>();
        if (sheets == null) return result;

        foreach (var sheet in sheets.Elements<Sheet>())
        {
            var name = sheet.Name?.Value ?? "?";
            var id = sheet.Id?.Value;
            if (id == null) continue;
            var part = (WorksheetPart)doc.WorkbookPart!.GetPartById(id);
            result.Add((name, part));
        }

        return result;
    }

    private static readonly System.Text.RegularExpressions.Regex SheetIndexPattern =
        new(@"^sheet\[(\d+)\]$", System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);

    private static readonly System.Text.RegularExpressions.Regex SheetLastPattern =
        new(@"^sheet\[last\(\)\]$", System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);

    /// <summary>
    /// Resolve a sheet name that may be a 1-based index reference like "sheet[1]"
    /// or the XPath-style "sheet[last()]" predicate to the actual sheet name.
    /// Returns the original name if not an index pattern.
    /// </summary>
    private string ResolveSheetName(string sheetName)
    {
        var m = SheetIndexPattern.Match(sheetName);
        if (m.Success && int.TryParse(m.Groups[1].Value, out var idx) && idx >= 1)
        {
            var sheets = GetWorksheets();
            if (idx <= sheets.Count)
                return sheets[idx - 1].Name;
        }
        // CONSISTENCY(path-stability): align with Word's p[last()] support
        // (commit 5b03d7a7) so sheet[last()] resolves to the last worksheet.
        if (SheetLastPattern.IsMatch(sheetName))
        {
            var sheets = GetWorksheets();
            if (sheets.Count > 0)
                return sheets[^1].Name;
        }
        return sheetName;
    }

    private WorksheetPart? FindWorksheet(string sheetName)
    {
        sheetName = ResolveSheetName(sheetName);
        foreach (var (name, part) in GetWorksheets())
        {
            if (name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                return part;
        }
        return null;
    }

    private ArgumentException SheetNotFoundException(string sheetName)
    {
        var available = GetWorksheets().Select(w => w.Name).ToList();
        var availableStr = available.Count > 0
            ? string.Join(", ", available)
            : "(none)";
        return new ArgumentException(
            $"Sheet not found: \"{sheetName}\". Available sheets: [{availableStr}]. " +
            $"Use DOM path \"/{available.FirstOrDefault() ?? "SheetName"}/A1\" or Excel notation \"{available.FirstOrDefault() ?? "SheetName"}!A1\".");
    }

    /// <summary>
    /// Find and replace text across all sheets (or a specific sheet). Returns the number of replacements made.
    /// Handles SharedStringTable entries as well as inline strings and direct cell values.
    /// </summary>
    private int FindAndReplace(string find, string replace, WorksheetPart? targetSheet)
    {
        if (string.IsNullOrEmpty(find)) return 0;
        int totalCount = 0;
        var workbookPart = _doc.WorkbookPart;
        if (workbookPart == null) return 0;

        // Replace in SharedStringTable (affects all sheets sharing these strings)
        if (targetSheet == null)
        {
            var sst = workbookPart.SharedStringTablePart?.SharedStringTable;
            if (sst != null)
            {
                foreach (var si in sst.Elements<SharedStringItem>())
                {
                    // Handle simple text items
                    var textEl = si.GetFirstChild<Text>();
                    if (textEl?.Text != null)
                    {
                        textEl.Text = ApplyFindReplace(textEl.Text, find, replace, out int count);
                        totalCount += count;
                    }

                    // Handle rich text runs
                    foreach (var run in si.Elements<Run>())
                    {
                        var runText = run.GetFirstChild<Text>();
                        if (runText?.Text != null)
                        {
                            runText.Text = ApplyFindReplace(runText.Text, find, replace, out int count);
                            totalCount += count;
                        }
                    }
                }
                sst.Save();
            }
        }

        // Replace in inline strings and direct cell values
        var sheets = targetSheet != null
            ? [targetSheet]
            : workbookPart.WorksheetParts.ToList();

        foreach (var wsPart in sheets)
        {
            var sheetData = wsPart.Worksheet?.GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    // Inline string
                    var inlineStr = cell.GetFirstChild<InlineString>();
                    if (inlineStr != null)
                    {
                        var t = inlineStr.GetFirstChild<Text>();
                        if (t?.Text != null)
                        {
                            t.Text = ApplyFindReplace(t.Text, find, replace, out int count);
                            totalCount += count;
                        }
                        // Rich text runs inside inline string
                        foreach (var run in inlineStr.Elements<Run>())
                        {
                            var runText = run.GetFirstChild<Text>();
                            if (runText?.Text != null)
                            {
                                runText.Text = ApplyFindReplace(runText.Text, find, replace, out int count);
                                totalCount += count;
                            }
                        }
                        continue;
                    }

                    // Direct string value (DataType is null or String)
                    if (cell.DataType?.Value == CellValues.String)
                    {
                        // R20-1: a t="str" cell from `set formula=...` carries BOTH
                        // <f> (formula) and <v> (cached value); the evaluator reads
                        // <f>, so touching only <v> silently reverts on next eval.
                        if (cell.CellFormula?.Text is { } fText)
                        {
                            // Only rewrite <f> when the formula is a PURE string
                            // literal ("...") — a single quoted string with no
                            // interior quote and nothing else. Blindly replacing
                            // inside a real formula corrupts cell refs / function
                            // tokens (find "2026" → B2026&"..." becomes B2027&...;
                            // find "A1" → SUM(A1:A10) becomes SUM(X:A10)). Anything
                            // that isn't a bare literal is skipped untouched + warned,
                            // so <f>/<v> stay consistent and no formula is mangled.
                            var trimmedF = fText.Trim();
                            bool isPureLiteral = trimmedF.Length >= 2
                                && trimmedF[0] == '"' && trimmedF[^1] == '"'
                                && trimmedF.IndexOf('"', 1) == trimmedF.Length - 1;
                            if (isPureLiteral)
                            {
                                cell.CellFormula.Text = ApplyFindReplace(fText, find, replace, out int fCount);
                                totalCount += fCount;
                                if (cell.CellValue?.Text != null)
                                    cell.CellValue.Text = ApplyFindReplace(cell.CellValue.Text, find, replace, out _);
                            }
                            else if (CountOccurrences(fText, find) > 0
                                     || (cell.CellValue?.Text != null && CountOccurrences(cell.CellValue.Text, find) > 0))
                            {
                                // Pattern would have matched, but this is a real
                                // formula — skip it (no <f>/<v> change) and warn, so
                                // the user knows find/replace did not touch formulas.
                                Console.Error.WriteLine(
                                    $"warning: find/replace skipped formula cell {cell.CellReference?.Value ?? "?"} "
                                    + "(formula text not modified to avoid corrupting references/functions)");
                            }
                        }
                        else
                        {
                            var cv = cell.CellValue;
                            if (cv?.Text != null)
                            {
                                cv.Text = ApplyFindReplace(cv.Text, find, replace, out int count);
                                totalCount += count;
                            }
                        }
                    }

                    // SharedStringTable reference — if targeting a specific sheet, replace inline
                    if (targetSheet != null && cell.DataType?.Value == CellValues.SharedString)
                    {
                        var sst = workbookPart.SharedStringTablePart?.SharedStringTable;
                        if (sst != null && cell.CellValue?.Text != null
                            && int.TryParse(cell.CellValue.Text, out var sstIdx))
                        {
                            var items = sst.Elements<SharedStringItem>().ToList();
                            if (sstIdx >= 0 && sstIdx < items.Count)
                            {
                                var si = items[sstIdx];
                                var siText = si.GetFirstChild<Text>();
                                if (siText?.Text != null)
                                {
                                    siText.Text = ApplyFindReplace(siText.Text, find, replace, out int count);
                                    totalCount += count;
                                }
                                foreach (var run in si.Elements<Run>())
                                {
                                    var runText = run.GetFirstChild<Text>();
                                    if (runText?.Text != null)
                                    {
                                        runText.Text = ApplyFindReplace(runText.Text, find, replace, out int count);
                                        totalCount += count;
                                    }
                                }
                                sst.Save();
                            }
                        }
                    }
                }
            }

            SaveWorksheet(wsPart);
        }

        return totalCount;
    }

    private static int CountOccurrences(string text, string find)
    {
        int count = 0;
        int idx = 0;
        while ((idx = text.IndexOf(find, idx, StringComparison.Ordinal)) >= 0)
        {
            count++;
            idx += find.Length;
        }
        return count;
    }

    /// <summary>
    /// CONSISTENCY(find-regex): apply one find/replace to a text fragment,
    /// honouring the `r"..."` raw-string regex prefix the same way docx/pptx
    /// Set and the query `~=` operator do (shared parser
    /// AttributeFilter.TryParseRegexPrefix). r-prefixed → Regex.Replace
    /// (capture groups like $1 expand); plain literal → ordinal Contains/Replace.
    /// Returns the rewritten text and sets <paramref name="count"/> to the
    /// number of matches; returns the input unchanged with count 0 on no match.
    /// A malformed regex falls back to literal handling (never throws here).
    /// </summary>
    private static string ApplyFindReplace(string text, string find, string replace, out int count)
    {
        count = 0;
        if (AttributeFilter.TryParseRegexPrefix(find, out var pattern))
        {
            try
            {
                // CONSISTENCY(find-regex-timeout): mirror the 5s catastrophic-
                // backtracking guard Word/Pptx find use (WordHandler.Helpers.cs:20,
                // PowerPointHandler.Helpers.cs:18) so a pathological xlsx pattern
                // can't hang the process.
                var rx = new System.Text.RegularExpressions.Regex(
                    pattern, System.Text.RegularExpressions.RegexOptions.None, TimeSpan.FromSeconds(5));
                count = rx.Matches(text).Count;
                return count > 0 ? rx.Replace(text, replace) : text;
            }
            catch (System.ArgumentException)
            {
                // Malformed regex — fall through to literal handling.
            }
        }
        if (text.Contains(find, StringComparison.Ordinal))
        {
            count = CountOccurrences(text, find);
            return text.Replace(find, replace, StringComparison.Ordinal);
        }
        return text;
    }

    // R24-1: detect whether a styleProps bag asks for the text number format
    // ("@"). All three accepted aliases are checked: numberformat, numfmt,
    // format. Whitespace is trimmed; quoting is not expected here because
    // ExcelStyleManager already strips surrounding quotes upstream.
    // CT_Workbook schema order: ...sheets, functionGroups, externalReferences,
    // definedNames, calcPr, oleSize, customWorkbookViews, pivotCaches...
    // Returns existing <definedNames> or creates+inserts one in schema-correct
    // position. AppendChild lands after calcPr, which fails strict validators.
    private static DefinedNames GetOrCreateDefinedNames(Workbook workbook)
    {
        var definedNames = workbook.GetFirstChild<DefinedNames>();
        if (definedNames != null) return definedNames;
        definedNames = new DefinedNames();
        var insertBefore = (OpenXmlElement?)workbook.GetFirstChild<CalculationProperties>()
            ?? (OpenXmlElement?)workbook.GetFirstChild<OleSize>()
            ?? (OpenXmlElement?)workbook.GetFirstChild<CustomWorkbookViews>()
            ?? (OpenXmlElement?)workbook.GetFirstChild<PivotCaches>();
        if (insertBefore != null)
            workbook.InsertBefore(definedNames, insertBefore);
        else
            workbook.AppendChild(definedNames);
        return definedNames;
    }
}
