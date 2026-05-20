// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;
using SpreadsheetDrawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeCli.Handlers;

// Per-element-type Add helpers for table-like paths (namedrange, comment, validation, autofilter, table, pivottable). Mechanically extracted from the Add() god-method.
public partial class ExcelHandler
{
    private string AddNamedRange(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        // R4-4: accept `/namedrange[NAME]` path form so users don't
        // have to repeat the name in --prop name=. Path brackets take
        // precedence only when --prop name= is absent (explicit prop
        // still wins on mismatch, to keep other `/namedrange[N]` int
        // indexing semantics elsewhere in the handler usable as-is).
        var pathNrName = "";
        {
            var mNr = System.Text.RegularExpressions.Regex.Match(
                parentPath, @"^/namedrange\[([^\]]+)\]/?$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (mNr.Success)
            {
                var captured = mNr.Groups[1].Value;
                // Only treat as a name if it is not a pure integer
                // (preserves existing `/namedrange[1]` semantics).
                if (!int.TryParse(captured, out _))
                    pathNrName = captured;
            }
        }
        var nrName = properties.GetValueOrDefault("name", pathNrName);
        if (string.IsNullOrEmpty(nrName))
            throw new ArgumentException("'name' property is required for namedrange");
        // Per OOXML §18.2.5: defined-name identifiers must start with
        // letter/underscore/backslash, contain only letter/digit/
        // underscore/period/backslash, and must not parse as a cell
        // reference. Otherwise Excel rejects the file with 0x800A03EC.
        if (!System.Text.RegularExpressions.Regex.IsMatch(nrName, @"^[A-Za-z_\\][A-Za-z0-9_\\.]*$"))
            throw new ArgumentException($"Invalid defined-name '{nrName}': must start with a letter/underscore and contain only letters, digits, underscores, or periods (no spaces).");
        // Excel caps defined-name identifier length at 255 characters; longer
        // names are silently truncated on open (or the file is rejected with
        // 0x800A03EC depending on host). Refuse up front instead of letting
        // a 256+ char name land on disk and round-trip-differ on re-open.
        if (nrName.Length > 255)
            throw new ArgumentException($"Invalid defined-name '{nrName}': length {nrName.Length} exceeds the Excel 255-character limit.");
        if (LooksLikeCellReference(nrName))
            throw new ArgumentException($"Invalid defined-name '{nrName}': name parses as a cell reference; choose a different name.");
        // R39-5: Excel reserves the single letters R and C (case-insensitive)
        // because they collide with R1C1 reference notation. Excel rejects
        // the file with 0x800A03EC if either is used as a defined name.
        if (nrName.Length == 1 && (nrName[0] == 'R' || nrName[0] == 'r' || nrName[0] == 'C' || nrName[0] == 'c'))
            throw new ArgumentException($"Invalid defined-name '{nrName}': single letter 'R' / 'C' is reserved by Excel for R1C1 reference notation; choose a different name.");
        // `refersTo` is the common Excel-documented alias for `ref`;
        // silently map it so users don't end up with an empty
        // <x:definedName/> that corrupts the file.
        var refVal = properties.GetValueOrDefault("ref",
            properties.GetValueOrDefault("refersTo",
                properties.GetValueOrDefault("formula", "")));
        // R15/bt-2: reject up-front when the required ref/refersTo/formula
        // value is missing so an empty <x:definedName/> never gets written
        // (the resulting zombie polluted the workbook and broke later Set
        // calls). Unsupported aliases like `range=` previously silently
        // landed here as empty and produced the zombie.
        if (string.IsNullOrEmpty(refVal))
            throw new ArgumentException("'ref' (or 'refersTo' / 'formula') property is required for namedrange");
        // R7-2: per ECMA-376 §18.2.5, <x:definedName> content must NOT
        // have a leading '=' (unlike the formula-bar form in Excel UI).
        // Excel rejects the file with 0x800A03EC if '=' is present.
        if (refVal.StartsWith('='))
            refVal = refVal.TrimStart('=');

        // R27-1: cross-workbook references like "[Other.xlsx]Sheet1!$A$1"
        // or "[1]Sheet1!$A$1" need an externalReferences part to resolve.
        // Without one, Excel opens the file but formulas referencing the
        // name show #REF!. Reject up-front rather than write a silently
        // broken defined name.
        // CONSISTENCY(xref-detect): bt-5/fuzz-NR01 — also catch the
        // single-quoted form `'[Book.xlsx]Sheet'!A1` (Excel's standard
        // quoting for sheet names with spaces) which previously slipped
        // through and produced a silently broken defined name.
        if (System.Text.RegularExpressions.Regex.IsMatch(refVal, @"^\s*'?\["))
            throw new ArgumentException(
                $"Cross-workbook references like '{refVal}' require an externalLinks part which officecli doesn't expose; use raw-set for this case");

        var workbook = GetWorkbook();
        // CONSISTENCY(workbook-child-order): helper inserts <definedNames>
        // in schema-correct position (before calcPr/oleSize/...).
        var definedNames = GetOrCreateDefinedNames(workbook);

        var dn = new DefinedName(refVal) { Name = nrName };

        // Scope resolution: explicit --prop scope= wins. Otherwise, if the
        // parentPath addresses a specific sheet (e.g. "/Sheet1"), default
        // the scope to that sheet — callers who picked a sheet-level
        // parent almost always wanted a sheet-scoped name, and previously
        // the parentPath was silently ignored and the name landed at
        // workbook scope. Pass scope=workbook explicitly to keep the old
        // workbook-global behavior under a sheet parent.
        string? effectiveScope = null;
        bool explicitWorkbookScope = false;
        if (properties.TryGetValue("scope", out var scopeRaw) && !string.IsNullOrEmpty(scopeRaw))
        {
            if (string.Equals(scopeRaw, "workbook", StringComparison.OrdinalIgnoreCase))
                explicitWorkbookScope = true;
            else
                effectiveScope = scopeRaw;
        }
        if (effectiveScope == null && !explicitWorkbookScope)
        {
            var parentTrim = (parentPath ?? "").TrimStart('/').TrimEnd('/');
            if (!string.IsNullOrEmpty(parentTrim) && !parentTrim.Contains('/'))
                effectiveScope = parentTrim; // single-segment sheet name
        }
        if (effectiveScope != null)
        {
            var nrSheets = workbook.GetFirstChild<Sheets>()?.Elements<Sheet>().ToList();
            var nrSheetIdx = nrSheets?.FindIndex(s =>
                s.Name?.Value?.Equals(effectiveScope, StringComparison.OrdinalIgnoreCase) == true) ?? -1;
            if (nrSheetIdx >= 0) dn.LocalSheetId = (uint)nrSheetIdx;
        }
        if (properties.TryGetValue("comment", out var nrComment))
            dn.Comment = nrComment;
        // 'volatile' surfaces as DefinedName.Function in OOXML — Excel's
        // recalc engine treats function-flagged defined names as volatile,
        // forcing recalc on every workbook change.
        if (properties.TryGetValue("volatile", out var nrVolatile) && IsTruthy(nrVolatile))
            dn.Function = true;

        // CONSISTENCY(definedname-unique): Excel rejects two
        // <definedName> entries that share both name AND scope
        // (LocalSheetId) with a "found a problem" repair dialog.
        // Same name across different scopes (workbook-global vs
        // per-sheet, or two distinct sheets) is legal — only the
        // (name, localSheetId) pair must be unique.
        var dnLocalId = dn.LocalSheetId?.Value;
        foreach (var existingDn in definedNames.Elements<DefinedName>())
        {
            var existingName = existingDn.Name?.Value;
            if (existingName == null) continue;
            if (!string.Equals(existingName, nrName, StringComparison.OrdinalIgnoreCase)) continue;
            if (existingDn.LocalSheetId?.Value == dnLocalId)
                throw new ArgumentException(
                    $"Defined name '{nrName}' already exists" +
                    (dnLocalId.HasValue ? $" in sheet scope (localSheetId={dnLocalId})" : " in workbook scope") +
                    "; remove it before adding a new one or pick a different name.");
        }

        definedNames.AppendChild(dn);

        // R7-3: if the defined-name body is a formula (not just a pure
        // range reference), set fullCalcOnLoad so Excel recomputes on
        // first open — otherwise the name evaluates to 0 until the
        // user triggers a recalc.
        if (LooksLikeFormulaBody(refVal))
        {
            var calcPr = workbook.GetFirstChild<CalculationProperties>();
            if (calcPr == null)
            {
                calcPr = new CalculationProperties();
                var insertBefore = (DocumentFormat.OpenXml.OpenXmlElement?)workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.OleSize>()
                    ?? (DocumentFormat.OpenXml.OpenXmlElement?)workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.CustomWorkbookViews>()
                    ?? (DocumentFormat.OpenXml.OpenXmlElement?)workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.PivotCaches>();
                if (insertBefore != null)
                    workbook.InsertBefore(calcPr, insertBefore);
                else
                    workbook.AppendChild(calcPr);
            }
            calcPr.FullCalculationOnLoad = true;
        }

        workbook.Save();

        var nrIdx = definedNames.Elements<DefinedName>().ToList().IndexOf(dn) + 1;
        return $"/namedrange[{nrIdx}]";
    }

    private string AddComment(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        var cmtSegments = parentPath.TrimStart('/').Split('/', 2);
        var cmtSheetName = cmtSegments[0];
        // Extract cell reference from path if present (e.g., /Sheet1/A1 -> A1)
        string? cmtRefFromPath = null;
        if (cmtSegments.Length > 1 && Regex.IsMatch(cmtSegments[1], @"^[A-Z]+\d+$", RegexOptions.IgnoreCase))
            cmtRefFromPath = cmtSegments[1];
        var cmtWorksheet = FindWorksheet(cmtSheetName)
            ?? throw new ArgumentException($"Sheet not found: {cmtSheetName}");

        var cmtRef = properties.GetValueOrDefault("ref") ?? cmtRefFromPath
            ?? throw new ArgumentException("Property 'ref' is required for comment");
        // Validate cell reference up-front; ParseCellReference rejects bad
        // syntax, out-of-range rows (>1048576), and out-of-range columns (>XFD)
        // with a clear ArgumentException — matches the validation surface
        // already enforced for cells/ranges elsewhere.
        ParseCellReference(cmtRef);
        var cmtText = properties.GetValueOrDefault("text", "");
        var cmtAuthor = properties.GetValueOrDefault("author", "Author");

        var commentsPart = cmtWorksheet.WorksheetCommentsPart
            ?? cmtWorksheet.AddNewPart<WorksheetCommentsPart>();

        if (commentsPart.Comments == null)
        {
            commentsPart.Comments = new Comments(
                new Authors(new Author(cmtAuthor)),
                new CommentList()
            );
        }

        var comments = commentsPart.Comments;
        var authors = comments.GetFirstChild<Authors>()!;
        var commentList = comments.GetFirstChild<CommentList>()!;

        // CONSISTENCY(overlap-reject): duplicate comment on the same
        // cell is ambiguous — mirror the table T4 overlap-reject
        // pattern. User must `remove comment` first to replace it.
        var cmtRefUpper = cmtRef.ToUpperInvariant();
        if (commentList.Elements<Comment>().Any(c =>
                string.Equals(c.Reference?.Value, cmtRefUpper, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException(
                $"comment already exists on {cmtRefUpper}. Remove it first before adding a new comment.");

        uint authorId = 0;
        var existingAuthors = authors.Elements<Author>().ToList();
        var authorIdx = existingAuthors.FindIndex(a => a.Text == cmtAuthor);
        if (authorIdx >= 0)
            authorId = (uint)authorIdx;
        else
        {
            authors.AppendChild(new Author(cmtAuthor));
            authorId = (uint)existingAuthors.Count;
        }

        var comment = new Comment { Reference = cmtRef.ToUpperInvariant(), AuthorId = authorId };
        // Support user-supplied `\n` (literal two-char sequence from
        // CLI) and real LF as line breaks — Excel renders the
        // preserved newline in the comment body. Matches the shape
        // `text` behavior documented in add-shape help.
        // CONSISTENCY(text-escape-boundary): \n / \t resolution lives at the
        // CLI --prop parse boundary; this value already has real newlines.
        var cmtNormalized = (cmtText ?? "").Replace("\r\n", "\n");
        comment.CommentText = new CommentText(
            new Run(
                BuildCommentRunProperties(properties),
                new Text(cmtNormalized) { Space = SpaceProcessingModeValues.Preserve }
            )
        );
        commentList.AppendChild(comment);
        commentsPart.Comments.Save();

        if (!cmtWorksheet.VmlDrawingParts.Any())
        {
            var vmlPart = cmtWorksheet.AddNewPart<VmlDrawingPart>();
            using var writer = new System.IO.StreamWriter(vmlPart.GetStream());
            writer.Write("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout><v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype></xml>");
        }

        var cmtIdx = commentList.Elements<Comment>().ToList().IndexOf(comment) + 1;
        return $"/{cmtSheetName}/comment[{cmtIdx}]";
    }

    private string AddValidation(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        var dvSegments = parentPath.TrimStart('/').Split('/', 2);
        var dvSheetName = dvSegments[0];
        var dvWorksheet = FindWorksheet(dvSheetName)
            ?? throw new ArgumentException($"Sheet not found: {dvSheetName}");

        var dvSqref = properties.GetValueOrDefault("sqref")
            ?? properties.GetValueOrDefault("ref")
            ?? throw new ArgumentException("Property 'sqref' (or 'ref') is required for validation");

        var dv = new DataValidation
        {
            SequenceOfReferences = new ListValue<StringValue>(
                dvSqref.Split(' ').Select(s => new StringValue(s)))
        };

        if (properties.TryGetValue("type", out var dvType))
        {
            dv.Type = dvType.ToLowerInvariant() switch
            {
                "list" => DataValidationValues.List,
                "whole" => DataValidationValues.Whole,
                "decimal" => DataValidationValues.Decimal,
                "date" => DataValidationValues.Date,
                "time" => DataValidationValues.Time,
                "textlength" => DataValidationValues.TextLength,
                "custom" => DataValidationValues.Custom,
                _ => throw new ArgumentException($"Unknown validation type: {dvType}. Use: list, whole, decimal, date, time, textLength, custom")
            };
        }

        if (properties.TryGetValue("operator", out var dvOp))
        {
            dv.Operator = dvOp.ToLowerInvariant() switch
            {
                "between" => DataValidationOperatorValues.Between,
                "notbetween" => DataValidationOperatorValues.NotBetween,
                "equal" => DataValidationOperatorValues.Equal,
                "notequal" => DataValidationOperatorValues.NotEqual,
                "greaterthan" => DataValidationOperatorValues.GreaterThan,
                "lessthan" => DataValidationOperatorValues.LessThan,
                "greaterthanorequal" => DataValidationOperatorValues.GreaterThanOrEqual,
                "lessthanorequal" => DataValidationOperatorValues.LessThanOrEqual,
                _ => throw new ArgumentException($"Unknown operator: {dvOp}")
            };
        }

        if (properties.TryGetValue("formula1", out var dvFormula1))
        {
            // R28-A1 — reject empty formula1 for type=list. Excel renders an empty
            // dropdown (or rejects the file outright depending on form), and the
            // user almost certainly meant to provide options like "1,2,3".
            if (dv.Type?.Value == DataValidationValues.List
                && string.IsNullOrWhiteSpace(dvFormula1.Trim('"')))
                throw new ArgumentException(
                    "Property 'formula1' is empty for validation type=list; supply options like formula1=\"1,2,3\" or a range reference.");
            dv.Formula1 = new Formula1(NormalizeValidationFormula(dvFormula1, dv.Type?.Value));
        }
        else if (dv.Type?.Value == DataValidationValues.List)
        {
            // R28-A1 — type=list with no formula1 at all is also nonsense.
            throw new ArgumentException(
                "Property 'formula1' is required for validation type=list; supply options like formula1=\"1,2,3\" or a range reference.");
        }

        if (properties.TryGetValue("formula2", out var dvFormula2))
            dv.Formula2 = new Formula2(NormalizeValidationFormula(dvFormula2, dv.Type?.Value));
        else if (dv.Operator?.Value == DataValidationOperatorValues.Between
                 || dv.Operator?.Value == DataValidationOperatorValues.NotBetween)
        {
            // operator=between/notBetween needs both bounds. Without formula2
            // Excel silently treats the rule as "anything passes" — the file
            // opens but validates nothing. Reject up front rather than land a
            // permissive no-op on disk.
            throw new ArgumentException(
                $"Property 'formula2' is required when operator='{dv.Operator.InnerText}'; supply both bounds (formula1=lower, formula2=upper).");
        }

        // Build case-insensitive lookup for validation properties
        var dvProps = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);

        dv.AllowBlank = !dvProps.TryGetValue("allowBlank", out var dvAllowBlank)
            || IsTruthy(dvAllowBlank);
        dv.ShowErrorMessage = !dvProps.TryGetValue("showError", out var dvShowError)
            || IsTruthy(dvShowError);
        dv.ShowInputMessage = !dvProps.TryGetValue("showInput", out var dvShowInput)
            || IsTruthy(dvShowInput);

        if (dvProps.TryGetValue("errorTitle", out var dvErrorTitle))
            dv.ErrorTitle = dvErrorTitle;
        if (dvProps.TryGetValue("error", out var dvError))
            dv.Error = dvError;
        if (dvProps.TryGetValue("promptTitle", out var dvPromptTitle))
            dv.PromptTitle = dvPromptTitle;
        if (dvProps.TryGetValue("prompt", out var dvPrompt))
            dv.Prompt = dvPrompt;

        // V6 — errorStyle: stop (default), warning, information.
        if (dvProps.TryGetValue("errorStyle", out var dvErrStyle))
        {
            dv.ErrorStyle = dvErrStyle.ToLowerInvariant() switch
            {
                "stop" => DataValidationErrorStyleValues.Stop,
                "warning" or "warn" => DataValidationErrorStyleValues.Warning,
                "information" or "info" => DataValidationErrorStyleValues.Information,
                _ => throw new ArgumentException(
                    $"Unknown errorStyle: {dvErrStyle}. Use: stop, warning, information")
            };
        }

        // V7 — showDropDown / inCellDropdown. OOXML `showDropDown`
        // has INVERTED semantics: true = HIDE the in-cell arrow.
        // Expose it as `inCellDropdown` (user-friendly sense) and
        // the raw `showDropDown` (OOXML sense).
        if (dvProps.TryGetValue("inCellDropdown", out var dvInCell))
            dv.ShowDropDown = !ParseHelpers.IsTruthy(dvInCell);
        else if (dvProps.TryGetValue("showDropDown", out var dvShowDd))
            dv.ShowDropDown = ParseHelpers.IsTruthy(dvShowDd);

        var wsEl = GetSheet(dvWorksheet);
        var dvs = wsEl.GetFirstChild<DataValidations>();
        // R27-3: stacking a second DV on a sqref that overlaps an existing
        // DV is silently invisible in Excel (first wins). Reject up-front
        // rather than persist a useless rule.
        if (dvs != null)
        {
            var newRanges = dvSqref.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            foreach (var existing in dvs.Elements<DataValidation>())
            {
                var existingSqref = existing.SequenceOfReferences?.InnerText ?? "";
                var existingRanges = existingSqref.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                foreach (var nr in newRanges)
                    foreach (var er in existingRanges)
                        if (RangesOverlap(nr, er))
                            throw new ArgumentException(
                                $"DataValidation sqref '{nr}' overlaps existing validation sqref '{er}'; Excel ignores stacked validations on the same cells. Remove the existing validation first or use a non-overlapping range.");
            }
        }
        if (dvs == null)
        {
            dvs = new DataValidations();
            var insertAfter = wsEl.GetFirstChild<Hyperlinks>() as OpenXmlElement
                ?? wsEl.Elements<ConditionalFormatting>().LastOrDefault() as OpenXmlElement
                ?? wsEl.GetFirstChild<SheetData>() as OpenXmlElement;
            if (insertAfter is Hyperlinks)
                insertAfter.InsertBeforeSelf(dvs);
            else if (insertAfter != null)
                insertAfter.InsertAfterSelf(dvs);
            else
                wsEl.AppendChild(dvs);
        }

        dvs.AppendChild(dv);
        dvs.Count = (uint)dvs.Elements<DataValidation>().Count();

        SaveWorksheet(dvWorksheet);
        var dvIndex = dvs.Elements<DataValidation>().ToList().IndexOf(dv) + 1;
        // CONSISTENCY(path-segment-naming): the path segment must match the
        // type name the caller used in `add` (`dataValidation`). The legacy
        // `/validation[N]` form remains accepted by Get / Set / Remove as an
        // alias for back-compat (R7-bt-6).
        return $"/{dvSheetName}/dataValidation[{dvIndex}]";
    }

    private string AddAutoFilter(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        var afSegments = parentPath.TrimStart('/').Split('/', 2);
        var afSheetName = afSegments[0];
        var afWorksheet = FindWorksheet(afSheetName)
            ?? throw new ArgumentException($"Sheet not found: {afSheetName}");

        // CONSISTENCY(tracking-rebind): the criteriaN.OP loop below iterates
        // properties via foreach over the static Dictionary<,> type, which
        // bypasses TrackingPropertyDictionary's comparer. Mark every
        // criteriaN.OP key (and `range`) as consumed up-front so they
        // don't surface as false unsupported_property warnings. Keys that
        // don't match either pattern fall through to the existing UNSUPPORTED
        // path naturally.
        if (properties is OfficeCli.Core.TrackingPropertyDictionary afTracking)
        {
            var consumed = properties.Keys
                .Where(k => string.Equals(k, "range", StringComparison.OrdinalIgnoreCase)
                    || Regex.IsMatch(k, @"^criteria\d+\.[A-Za-z]+$"))
                .ToList();
            afTracking.MarkAllConsumed(consumed);
        }

        var afRange = properties.GetValueOrDefault("range")
            ?? throw new ArgumentException("AutoFilter requires 'range' property (e.g. range=A1:F100)");

        // CONSISTENCY(cellref-validate): reject garbage refs (e.g. "BADREF")
        // so Excel doesn't silently open with an invalid <x:autoFilter ref="...">.
        if (!Regex.IsMatch(afRange.Trim(),
                @"^\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?$",
                RegexOptions.IgnoreCase))
            throw new ArgumentException(
                $"Invalid 'range' value: '{afRange}'. Expected a cell range like 'A1:F100' or 'A1'.");

        // CONSISTENCY(autofilter-table-dup): a Table already owns its own
        // <autoFilter> internally; layering a sheet-level <autoFilter> over
        // the same range produces the duplicate that Excel rejects with a
        // "found a problem" repair dialog. Mirror the T4 overlap check
        // used by AddTable.
        var afRangeUpper = afRange.ToUpperInvariant();
        foreach (var existingTdp in afWorksheet.TableDefinitionParts)
        {
            var existingTable = existingTdp.Table;
            if (existingTable?.Reference?.Value is string existingTableRef
                && RangesOverlap(afRangeUpper, existingTableRef.ToUpperInvariant()))
                throw new ArgumentException(
                    $"AutoFilter range '{afRangeUpper}' overlaps existing table " +
                    $"'{existingTable.Name?.Value ?? existingTable.DisplayName?.Value}' " +
                    $"({existingTableRef}); tables already include their own autoFilter.");
        }

        var wsElement = GetSheet(afWorksheet);
        var autoFilter = wsElement.GetFirstChild<AutoFilter>();
        if (autoFilter == null)
        {
            autoFilter = new AutoFilter();
            // AutoFilter goes after SheetData (after MergeCells if present)
            var mergeCellsEl = wsElement.GetFirstChild<MergeCells>();
            var sheetDataEl = wsElement.GetFirstChild<SheetData>();
            if (mergeCellsEl != null)
                mergeCellsEl.InsertAfterSelf(autoFilter);
            else if (sheetDataEl != null)
                sheetDataEl.InsertAfterSelf(autoFilter);
            else
                wsElement.AppendChild(autoFilter);
        }
        autoFilter.Reference = afRange.ToUpperInvariant();

        // AF1: per-column criteria. Syntax: criteriaN.OP=VAL where
        // N is 0-based column offset from the filter range's
        // leftmost column and OP is one of:
        //   equals, contains, gt, lt, top, blanks, nonBlanks
        // Each distinct N builds one <x:filterColumn colId="N">.
        // Previous criteria for the same N are replaced.
        var criteriaGroups = new Dictionary<uint, List<(string op, string val)>>();
        foreach (var (k, v) in properties)
        {
            var cm = Regex.Match(k, @"^criteria(\d+)\.([A-Za-z]+)$");
            if (!cm.Success) continue;
            var colId = uint.Parse(cm.Groups[1].Value);
            var op = cm.Groups[2].Value.ToLowerInvariant();
            if (!criteriaGroups.TryGetValue(colId, out var list))
                criteriaGroups[colId] = list = new List<(string, string)>();
            list.Add((op, v));
        }
        // Strip any prior filterColumn entries so a re-Add is idempotent
        foreach (var fc in autoFilter.Elements<FilterColumn>().ToList())
            fc.Remove();
        foreach (var (colId, entries) in criteriaGroups.OrderBy(kv => kv.Key))
        {
            var filterColumn = new FilterColumn { ColumnId = colId };
            // Dispatch by operator family. Top-N, Blanks, value-list,
            // and dynamicFilter build dedicated child elements;
            // text/number ops feed into <customFilters>.
            var customEntries = new List<(FilterOperatorValues fop, string val)>();
            bool customFilterAnd = false;
            bool handledDedicated = false;
            foreach (var (op, rawVal) in entries)
            {
                switch (op)
                {
                    case "equals":
                        customEntries.Add((FilterOperatorValues.Equal, rawVal));
                        break;
                    case "notequals":
                        customEntries.Add((FilterOperatorValues.NotEqual, rawVal));
                        break;
                    case "contains":
                    {
                        var wild = rawVal.Contains('*') ? rawVal : $"*{rawVal}*";
                        customEntries.Add((FilterOperatorValues.Equal, wild));
                        break;
                    }
                    case "doesnotcontain":
                    {
                        var wild = rawVal.Contains('*') ? rawVal : $"*{rawVal}*";
                        customEntries.Add((FilterOperatorValues.NotEqual, wild));
                        break;
                    }
                    case "beginswith":
                    {
                        var wild = rawVal.EndsWith("*") ? rawVal : $"{rawVal}*";
                        customEntries.Add((FilterOperatorValues.Equal, wild));
                        break;
                    }
                    case "endswith":
                    {
                        var wild = rawVal.StartsWith("*") ? rawVal : $"*{rawVal}";
                        customEntries.Add((FilterOperatorValues.Equal, wild));
                        break;
                    }
                    case "gt":
                        customEntries.Add((FilterOperatorValues.GreaterThan, rawVal));
                        break;
                    case "gte":
                        customEntries.Add((FilterOperatorValues.GreaterThanOrEqual, rawVal));
                        break;
                    case "lt":
                        customEntries.Add((FilterOperatorValues.LessThan, rawVal));
                        break;
                    case "lte":
                        customEntries.Add((FilterOperatorValues.LessThanOrEqual, rawVal));
                        break;
                    case "between":
                    case "notbetween":
                    {
                        var parts = rawVal.Split(',');
                        if (parts.Length != 2)
                            throw new ArgumentException(
                                $"criteria{colId}.{op} requires 'lo,hi', got: '{rawVal}'");
                        var lo = parts[0].Trim();
                        var hi = parts[1].Trim();
                        if (op == "between")
                        {
                            customEntries.Add((FilterOperatorValues.GreaterThanOrEqual, lo));
                            customEntries.Add((FilterOperatorValues.LessThanOrEqual, hi));
                            customFilterAnd = true;
                        }
                        else
                        {
                            // notBetween = lt lo OR gt hi (Excel default OR)
                            customEntries.Add((FilterOperatorValues.LessThan, lo));
                            customEntries.Add((FilterOperatorValues.GreaterThan, hi));
                        }
                        break;
                    }
                    case "top":
                    case "toppercent":
                    case "bottom":
                    case "bottompercent":
                    {
                        if (!double.TryParse(rawVal, System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture, out var topN))
                            throw new ArgumentException(
                                $"criteria{colId}.{op} requires a numeric value, got: '{rawVal}'");
                        filterColumn.Top10 = new Top10
                        {
                            Top = op == "top" || op == "toppercent",
                            Percent = op == "toppercent" || op == "bottompercent",
                            Val = topN
                        };
                        handledDedicated = true;
                        break;
                    }
                    case "blanks":
                        if (IsTruthy(rawVal))
                        {
                            filterColumn.Filters = new Filters { Blank = true };
                            handledDedicated = true;
                        }
                        break;
                    case "nonblanks":
                        if (IsTruthy(rawVal))
                        {
                            customEntries.Add((FilterOperatorValues.NotEqual, ""));
                        }
                        break;
                    case "values":
                    {
                        // Discrete value-list filter: comma-separated
                        // (split+trim empty; escape \, not supported).
                        var vals = rawVal.Split(',')
                            .Select(s => s.Trim())
                            .Where(s => s.Length > 0)
                            .ToList();
                        var filters = filterColumn.Filters ?? (filterColumn.Filters = new Filters());
                        foreach (var v in vals)
                            filters.AppendChild(new Filter { Val = v });
                        handledDedicated = true;
                        break;
                    }
                    case "dynamic":
                    {
                        var dyn = new DynamicFilter
                        {
                            Type = new EnumValue<DynamicFilterValues>(new DynamicFilterValues(rawVal))
                        };
                        filterColumn.DynamicFilter = dyn;
                        handledDedicated = true;
                        break;
                    }
                    default:
                        throw new ArgumentException(
                            $"Unsupported criteria operator: '{op}'. Valid: equals, notEquals, contains, doesNotContain, beginsWith, endsWith, gt, gte, lt, lte, between, notBetween, top, topPercent, bottom, bottomPercent, blanks, nonBlanks, values, dynamic.");
                }
            }
            if (customEntries.Count > 0 && !handledDedicated)
            {
                var cf = new CustomFilters();
                if (customFilterAnd)
                    cf.And = true;
                foreach (var (fop, val) in customEntries)
                    cf.AppendChild(new CustomFilter
                    {
                        Operator = fop,
                        Val = val
                    });
                filterColumn.CustomFilters = cf;
            }
            autoFilter.AppendChild(filterColumn);
        }

        SaveWorksheet(afWorksheet);
        return $"/{afSheetName}/autofilter";
    }

    private string AddTable(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        var tblSegments = parentPath.TrimStart('/').Split('/', 2);
        var tblSheetName = tblSegments[0];
        var tblWorksheet = FindWorksheet(tblSheetName)
            ?? throw new ArgumentException($"Sheet not found: {tblSheetName}");

        var rangeRef = (properties.GetValueOrDefault("ref") ?? properties.GetValueOrDefault("range")
            ?? throw new ArgumentException("Property 'ref' or 'range' is required for table")).ToUpperInvariant();

        // T4 — reject a new table whose ref overlaps any existing table on
        // the same sheet. Excel silently corrupts the file otherwise.
        foreach (var existingTdp in tblWorksheet.TableDefinitionParts)
        {
            var existing = existingTdp.Table;
            if (existing?.Reference?.Value is not string existingRef) continue;
            if (RangesOverlap(rangeRef, existingRef))
                throw new ArgumentException(
                    $"Table ref overlaps existing table '{existing.Name?.Value ?? existing.DisplayName?.Value}' ({existingRef})");
        }


        var existingTableIds = _doc.WorkbookPart!.WorksheetParts
            .SelectMany(wp => wp.TableDefinitionParts)
            .Select(tdp => tdp.Table?.Id?.Value ?? 0);
        var tableId = existingTableIds.Any() ? existingTableIds.Max() + 1 : 1;

        var userProvidedName = properties.ContainsKey("name");
        var tableName = SanitizeTableIdentifier(
            properties.GetValueOrDefault("name", $"Table{tableId}"),
            userProvided: userProvidedName);
        // displayName defaults to the (already-sanitized) tableName; if
        // name was user-provided it flows through verbatim so Excel
        // shows the same identifier the user asked for.
        var userProvidedDisplay = properties.ContainsKey("displayName");
        var displayName = SanitizeTableIdentifier(
            properties.GetValueOrDefault("displayName", tableName),
            userProvided: userProvidedDisplay || userProvidedName);

        // CONSISTENCY(table-name-unique): Excel requires both name and
        // displayName to be unique workbook-wide. A duplicate across
        // sheets surfaces a "found a problem" repair dialog. Walk every
        // WorksheetPart's tables, comparing case-insensitively.
        foreach (var existingTable in _doc.WorkbookPart!.WorksheetParts
            .SelectMany(wp => wp.TableDefinitionParts)
            .Select(tdp => tdp.Table)
            .Where(t => t != null)!)
        {
            if (string.Equals(existingTable!.Name?.Value, tableName, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    $"Table name '{tableName}' already exists in workbook; choose a different name.");
            if (string.Equals(existingTable.DisplayName?.Value, displayName, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    $"Table displayName '{displayName}' already exists in workbook; choose a different displayName.");
        }
        var styleName = properties.GetValueOrDefault("style", "TableStyleMedium2");
        // BUG-R9-B2: accept short aliases (medium2, light1, dark1, none) — schema
        // documents these but ValidateTableStyleName only accepted full names.
        styleName = NormalizeTableStyleName(styleName) ?? styleName;
        // T6 — validate style name against the built-in whitelist +
        // any workbook-level customStyles. Unknown names silently
        // fell through to Excel which would either ignore or
        // reject the file; prefer an explicit ArgumentException.
        ValidateTableStyleName(styleName);
        // T1 — accept `showHeader=false` alias alongside `headerRow=false`.
        var hasHeader = !(properties.TryGetValue("headerRow", out var hrVal) && !IsTruthy(hrVal))
                     && !(properties.TryGetValue("showHeader", out var shVal) && !IsTruthy(shVal));
        // CONSISTENCY(table-totalrow): accept `showTotals=true` alias
        // alongside `totalRow=true` (mirrors the `showHeader` alias
        // pattern above for users coming from Office API vocabulary).
        var hasTotalRow = (properties.TryGetValue("totalRow", out var trVal) && IsTruthy(trVal))
                       || (properties.TryGetValue("showTotals", out var stVal) && IsTruthy(stVal));

        var rangeParts = rangeRef.Split(':');
        var (startCol, startRow) = ParseCellReference(rangeParts[0]);
        var (endCol, endRow) = ParseCellReference(rangeParts[1]);
        var startColIdx = ColumnNameToIndex(startCol);
        var endColIdx = ColumnNameToIndex(endCol);
        var colCount = endColIdx - startColIdx + 1;

        // T5-ext: autoExpand=true probes the sheet for contiguous
        // non-empty rows immediately below the declared ref and grows
        // endRow to include them. Mirrors Excel's "Table expand when
        // you type below" behavior at Add time.
        if (properties.TryGetValue("autoExpand", out var autoExpandRaw) && IsTruthy(autoExpandRaw))
        {
            var sheetDataForProbe = GetSheet(tblWorksheet).GetFirstChild<SheetData>();
            if (sheetDataForProbe != null)
            {
                int probeRow = endRow + 1;
                while (true)
                {
                    var probe = sheetDataForProbe.Elements<Row>()
                        .FirstOrDefault(r => r.RowIndex?.Value == (uint)probeRow);
                    if (probe == null) break;
                    // non-empty = at least one cell in the column
                    // span carries a CellValue or InlineString.
                    bool anyNonEmpty = false;
                    for (int ci = 0; ci < colCount; ci++)
                    {
                        var cLetter = IndexToColumnName(startColIdx + ci);
                        var cRef = $"{cLetter}{probeRow}";
                        var probeCell = probe.Elements<Cell>()
                            .FirstOrDefault(c => c.CellReference?.Value == cRef);
                        if (probeCell == null) continue;
                        if (probeCell.CellValue != null || probeCell.InlineString != null)
                        {
                            anyNonEmpty = true;
                            break;
                        }
                    }
                    if (!anyNonEmpty) break;
                    endRow = probeRow;
                    probeRow++;
                }
                rangeRef = $"{startCol}{startRow}:{endCol}{endRow}";
            }
        }

        // i103: when headerRow=true (the default) the table ref must cover
        // at least 2 rows — header plus one data row. A header-only ref
        // (e.g. A1:C1) produces an <autoFilter> that Excel rejects with
        // "Removed Feature: AutoFilter from /xl/tables/tableN.xml part",
        // which cascades to drop the whole table on file open. Reject up
        // front with a clear message instead of letting Excel silently
        // strip the table. headerRow=false is allowed to be a single
        // (data-only) row.
        if (hasHeader && startRow == endRow)
        {
            throw new ArgumentException(
                $"table ref '{rangeRef}' has 1 row; tables must have at least 2 rows " +
                "(header + 1 data row). Pass headerRow=false for a data-only single-row table.");
        }

        // CONSISTENCY(table-totalrow): a:totalsRowShown MUST point at a row
        // OUTSIDE the data area. Previously we reused endRow as the totals
        // row, which overwrote whatever data lived on that last row. Expand
        // the ref by one row so the totals row is appended below the data
        // instead of stamping over it.
        if (hasTotalRow)
        {
            endRow += 1;
            rangeRef = $"{startCol}{startRow}:{endCol}{endRow}";
        }

        string[] colNames;
        if (properties.TryGetValue("columns", out var tblColsStr))
        {
            var userColNames = tblColsStr.Split(',').Select(c => c.Trim()).ToArray();
            // Pad with default names if fewer columns provided than range requires
            colNames = new string[colCount];
            for (int i = 0; i < colCount; i++)
                colNames[i] = i < userColNames.Length ? userColNames[i] : $"Column{i + 1}";

        }
        else
        {
            colNames = new string[colCount];
            if (hasHeader)
            {
                var tblSheetData = GetSheet(tblWorksheet).GetFirstChild<SheetData>();
                var headerRow = tblSheetData?.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == (uint)startRow);
                for (int i = 0; i < colCount; i++)
                {
                    var colLetter = IndexToColumnName(startColIdx + i);
                    var cellRefStr = $"{colLetter}{startRow}";
                    var headerCell = headerRow?.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == cellRefStr);
                    colNames[i] = (headerCell != null ? GetCellDisplayValue(headerCell) : null) ?? $"Column{i + 1}";
                    if (string.IsNullOrEmpty(colNames[i]))
                        colNames[i] = $"Column{i + 1}";
                    // Excel rejects a table whose header cell is typed
                    // as a number. Convert the cell to an inline string
                    // so the header reads as text, and tableColumn name
                    // (read above) still matches the cell's visible
                    // value exactly — Excel also requires that match.
                    if (headerCell != null && (headerCell.DataType == null || headerCell.DataType.Value == CellValues.Number))
                    {
                        var text = colNames[i];
                        headerCell.DataType = CellValues.InlineString;
                        headerCell.CellValue = null;
                        headerCell.InlineString = new InlineString(new Text(text));
                    }
                }
            }
            else
            {
                for (int i = 0; i < colCount; i++)
                    colNames[i] = $"Column{i + 1}";
            }
        }

        var tableDefPart = tblWorksheet.AddNewPart<TableDefinitionPart>();
        var table = new Table
        {
            Id = (uint)tableId,
            Name = tableName,
            DisplayName = displayName,
            Reference = rangeRef,
            TotalsRowShown = hasTotalRow
        };
        if (hasTotalRow)
            table.TotalsRowCount = 1;
        if (!hasHeader)
            table.HeaderRowCount = 0;

        table.AppendChild(new AutoFilter { Reference = rangeRef });

        // CONSISTENCY(autofilter-table-dup): Excel rejects a worksheet that
        // carries both a sheet-level <autoFilter> AND a <tableParts> reference
        // whose underlying table covers the same range — the table already
        // owns its own <autoFilter> for that range, and Excel surfaces a
        // "found a problem" repair dialog on the duplicate. Drop the sheet-
        // level filter whenever it overlaps the new table.
        var existingSheetFilter = tblWorksheet.Worksheet?.GetFirstChild<AutoFilter>();
        if (existingSheetFilter?.Reference?.Value is string existingFilterRef
            && RangesOverlap(rangeRef, existingFilterRef.ToUpperInvariant()))
        {
            existingSheetFilter.Remove();
        }

        // Dedupe duplicate column names (Excel also trips on those).
        var usedColNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < colCount; i++)
        {
            var baseName = colNames[i];
            var cn = baseName;
            var dedupIdx = 2;
            while (!usedColNames.Add(cn))
                cn = $"{baseName}{dedupIdx++}";
            colNames[i] = cn;
        }

        // CONSISTENCY(tablecolumn-header-match): after dedupe finalizes
        // colNames, force the header row cells to match. Excel rejects a
        // table whose <tableColumn name="X"> differs from the visible
        // text of its header cell. The implicit-discovery path above
        // already harmonized header cells while reading them; this pass
        // additionally covers (a) the explicit `columns=` path that
        // previously left header cells untouched, (b) padded `ColumnN`
        // names when fewer columns supplied than the range needs, and
        // (c) post-dedupe renames like X → X2.
        if (hasHeader)
        {
            var hdrSheetData = GetSheet(tblWorksheet).GetFirstChild<SheetData>()
                ?? GetSheet(tblWorksheet).AppendChild(new SheetData());
            var hdrRow = hdrSheetData.Elements<Row>()
                .FirstOrDefault(r => r.RowIndex?.Value == (uint)startRow);
            if (hdrRow == null)
            {
                hdrRow = new Row { RowIndex = (uint)startRow };
                var insertAfter = hdrSheetData.Elements<Row>()
                    .Where(r => r.RowIndex?.Value < (uint)startRow)
                    .LastOrDefault();
                if (insertAfter != null) insertAfter.InsertAfterSelf(hdrRow);
                else hdrSheetData.PrependChild(hdrRow);
            }
            for (int i = 0; i < colCount; i++)
            {
                var colLetter = IndexToColumnName(startColIdx + i);
                var cellRefStr = $"{colLetter}{startRow}";
                var headerCell = hdrRow.Elements<Cell>()
                    .FirstOrDefault(c => c.CellReference?.Value == cellRefStr);
                if (headerCell == null)
                {
                    headerCell = new Cell { CellReference = cellRefStr };
                    var insertBefore = hdrRow.Elements<Cell>()
                        .FirstOrDefault(c => ColumnNameToIndex(
                            System.Text.RegularExpressions.Regex.Match(
                                c.CellReference?.Value ?? "", @"^[A-Z]+").Value) > startColIdx + i);
                    if (insertBefore != null) insertBefore.InsertBeforeSelf(headerCell);
                    else hdrRow.AppendChild(headerCell);
                }
                // Stamp inline-string with the final column name. Skip when
                // the cell already shows exactly this text via shared/inline
                // strings, to leave shared-string indexes alone in the common
                // (already-matching) implicit-discovery case.
                var current = GetCellDisplayValue(headerCell);
                if (!string.Equals(current, colNames[i], StringComparison.Ordinal))
                {
                    headerCell.DataType = CellValues.InlineString;
                    headerCell.CellValue = null;
                    headerCell.CellFormula = null;
                    headerCell.InlineString = new InlineString(new Text(colNames[i]));
                }
            }
        }

        var tableColumns = new TableColumns { Count = (uint)colCount };
        for (int i = 0; i < colCount; i++)
            tableColumns.AppendChild(new TableColumn { Id = (uint)(i + 1), Name = colNames[i] });
        table.AppendChild(tableColumns);

        // T-ext: detect uniform formula pattern per column and emit
        // <x:calculatedColumnFormula> so Excel auto-fills the formula
        // into new rows appended to the table. Heuristic: if every data
        // row in a column carries a CellFormula whose relative form
        // (row numbers stripped) is identical, treat it as a calc'd
        // column and store the first row's formula.
        {
            var ccfSheetData = GetSheet(tblWorksheet).GetFirstChild<SheetData>();
            var dataStart = hasHeader ? startRow + 1 : startRow;
            var dataEnd = hasTotalRow ? endRow - 1 : endRow;
            if (ccfSheetData != null && dataEnd >= dataStart)
            {
                var tblColElems = tableColumns.Elements<TableColumn>().ToList();
                for (int ci = 0; ci < colCount; ci++)
                {
                    var colLetter = IndexToColumnName(startColIdx + ci);
                    string? firstFormula = null;
                    string? pattern = null;
                    bool uniform = true;
                    for (int r = dataStart; r <= dataEnd; r++)
                    {
                        var row = ccfSheetData.Elements<Row>()
                            .FirstOrDefault(rr => rr.RowIndex?.Value == (uint)r);
                        if (row == null) { uniform = false; break; }
                        var cellRefS = $"{colLetter}{r}";
                        var c = row.Elements<Cell>()
                            .FirstOrDefault(x => x.CellReference?.Value == cellRefS);
                        var f = c?.CellFormula?.Text;
                        if (string.IsNullOrEmpty(f)) { uniform = false; break; }
                        // Strip row numbers so =J2*K2 and =J3*K3 collapse to =J*K
                        var relF = System.Text.RegularExpressions.Regex.Replace(
                            f, @"\$?\d+", "");
                        if (pattern == null) { pattern = relF; firstFormula = f; }
                        else if (relF != pattern) { uniform = false; break; }
                    }
                    if (uniform && firstFormula != null)
                    {
                        tblColElems[ci].CalculatedColumnFormula =
                            new CalculatedColumnFormula(firstFormula);
                    }
                }
            }
        }

        // T7-ext: `columns.N.dxfId=<id>` stamps dataDxfId on the
        // target tableColumn (N is 1-based). The id must reference
        // an existing workbook differentialFormats entry; we do not
        // synthesize new dxfs here — users who want inline style
        // values should register a dxf first via `add dxf` (or the
        // underlying APIs) and then reference it.
        var tblColList = tableColumns.Elements<TableColumn>().ToList();
        foreach (var (rawKey, rawVal) in properties)
        {
            var m = Regex.Match(rawKey, @"^columns?\.(\d+)\.dxfId$",
                RegexOptions.IgnoreCase);
            if (!m.Success) continue;
            var n = int.Parse(m.Groups[1].Value);
            if (n < 1 || n > tblColList.Count) continue;
            if (!uint.TryParse(rawVal, out var dxfId))
                throw new ArgumentException(
                    $"columns.{n}.dxfId requires a numeric dxf id, got: '{rawVal}'");
            tblColList[n - 1].DataFormatId = dxfId;
        }

        // T2 — wire the banded rows/columns + first/last column
        // flags onto the TableStyleInfo. Each accepts `showX` or
        // its alias; default matches the old hard-coded values so
        // omitting them is identical to previous behavior.
        table.AppendChild(new TableStyleInfo
        {
            Name = styleName,
            ShowFirstColumn = (properties.TryGetValue("showFirstColumn", out var sfc)
                    || properties.TryGetValue("firstColumn", out sfc)
                    || properties.TryGetValue("firstCol", out sfc))
                ? IsTruthy(sfc) : false,
            ShowLastColumn = (properties.TryGetValue("showLastColumn", out var slc)
                    || properties.TryGetValue("lastColumn", out slc)
                    || properties.TryGetValue("lastCol", out slc))
                ? IsTruthy(slc) : false,
            // Accept showBandedRows / showRowStripes / bandedRows as aliases.
            // Set.Tables.cs already accepts the same set; mirror here.
            ShowRowStripes = (properties.TryGetValue("showBandedRows", out var sbr)
                    || properties.TryGetValue("showRowStripes", out sbr)
                    || properties.TryGetValue("bandedRows", out sbr))
                ? IsTruthy(sbr) : true,
            ShowColumnStripes = (properties.TryGetValue("showBandedColumns", out var sbc)
                    || properties.TryGetValue("showColumnStripes", out sbc)
                    || properties.TryGetValue("bandedColumns", out sbc)
                    || properties.TryGetValue("bandedCols", out sbc))
                ? IsTruthy(sbc) : false
        });

        // Generate total row content in SheetData when totalRow is enabled
        if (hasTotalRow)
        {
            var tblSheetData = GetSheet(tblWorksheet).GetFirstChild<SheetData>()
                ?? GetSheet(tblWorksheet).AppendChild(new SheetData());
            var totalRowIdx = (uint)endRow;
            var totalRow = tblSheetData.Elements<Row>()
                .FirstOrDefault(r => r.RowIndex?.Value == totalRowIdx);
            if (totalRow == null)
            {
                totalRow = new Row { RowIndex = totalRowIdx };
                // Insert in correct position
                var lastRow = tblSheetData.Elements<Row>()
                    .Where(r => r.RowIndex?.Value < totalRowIdx)
                    .LastOrDefault();
                if (lastRow != null)
                    lastRow.InsertAfterSelf(totalRow);
                else
                    tblSheetData.AppendChild(totalRow);
            }

            var tblCols = tableColumns.Elements<TableColumn>().ToList();
            // Per-column totalsRowFunction tokens: "none,sum,average"
            // → first col = label/none, rest = sum, average. If the
            // user didn't pass it, default to "none" on col0 + "sum"
            // on the rest (legacy behavior).
            string[] trfTokens = properties.TryGetValue("totalsRowFunction", out var trfRaw)
                ? trfRaw.Split(',').Select(s => s.Trim()).ToArray()
                : Array.Empty<string>();
            for (int ci = 0; ci < tblCols.Count; ci++)
            {
                var colLetter = IndexToColumnName(startColIdx + ci);
                var cellRefStr = $"{colLetter}{totalRowIdx}";
                var existingCell = totalRow.Elements<Cell>()
                    .FirstOrDefault(c => c.CellReference?.Value == cellRefStr);
                if (existingCell == null)
                {
                    existingCell = new Cell { CellReference = cellRefStr };
                    totalRow.AppendChild(existingCell);
                }

                var tokRaw = ci < trfTokens.Length ? trfTokens[ci].ToLowerInvariant() : "";
                var (trfEnum, subtotalCode) = MapTotalsRowFunction(tokRaw);

                if (ci == 0 && (tokRaw == "" || tokRaw == "none" || tokRaw == "label"))
                {
                    // First column: label "Total"
                    tblCols[ci].TotalsRowLabel = "Total";
                    existingCell.CellValue = new CellValue("Total");
                    existingCell.DataType = new EnumValue<CellValues>(CellValues.String);
                }
                else if (trfEnum == TotalsRowFunctionValues.None)
                {
                    // Skip — leave cell empty, no function set.
                }
                else
                {
                    // Default non-first column (no explicit token) = SUM
                    if (ci > 0 && tokRaw == "")
                    {
                        trfEnum = TotalsRowFunctionValues.Sum;
                        subtotalCode = 109;
                    }
                    tblCols[ci].TotalsRowFunction = trfEnum;
                    var dataStartRow = hasHeader ? startRow + 1 : startRow;
                    var dataEndRow = (int)totalRowIdx - 1;
                    var formulaRange = $"{colLetter}{dataStartRow}:{colLetter}{dataEndRow}";
                    existingCell.CellFormula = new CellFormula($"SUBTOTAL({subtotalCode},{formulaRange})");
                }
            }

            // T10: per-column custom totalsFormula override. Syntax:
            //   columns.N.totalsFormula="=SUM(Table1[Sales])/2"
            // where N is 1-based. This sets the column's
            // totalsRowFunction to "custom" + writes <calculatedColumnFormula>,
            // and replaces the SUBTOTAL cell formula with the user's.
            foreach (var (rawKey, rawVal) in properties)
            {
                var m = Regex.Match(rawKey, @"^columns?\.(\d+)\.totalsFormula$",
                    RegexOptions.IgnoreCase);
                if (!m.Success) continue;
                var n = int.Parse(m.Groups[1].Value);
                if (n < 1 || n > tblCols.Count) continue;
                var ci = n - 1;
                var colLetter = IndexToColumnName(startColIdx + ci);
                var cellRefStr = $"{colLetter}{totalRowIdx}";
                var existingCell = totalRow.Elements<Cell>()
                    .FirstOrDefault(c => c.CellReference?.Value == cellRefStr)
                    ?? totalRow.AppendChild(new Cell { CellReference = cellRefStr });

                var customFormula = rawVal.TrimStart('=');
                tblCols[ci].TotalsRowFunction = TotalsRowFunctionValues.Custom;
                tblCols[ci].TotalsRowLabel = null;
                tblCols[ci].TotalsRowFormula = new TotalsRowFormula(customFormula);
                existingCell.CellFormula = new CellFormula(OfficeCli.Core.PivotTableHelper.SanitizeXmlText(customFormula));
                existingCell.CellValue = null;
                existingCell.DataType = null;
            }
        }

        // CONSISTENCY(xlsx/table-autoexpand): persist the opt-in flag as
        // a custom-namespace attribute on <x:table> so eager auto-grow
        // survives reopen. Real Excel ignores unknown-namespace attrs.
        if (properties.TryGetValue("autoExpand", out var aeRaw) && IsTruthy(aeRaw))
            SetTableAutoExpandMarker(table, true);

        tableDefPart.Table = table;
        tableDefPart.Table.Save();

        var tblWs = GetSheet(tblWorksheet);
        var tableParts = tblWs.GetFirstChild<TableParts>();
        if (tableParts == null)
        {
            tableParts = new TableParts();
            tblWs.AppendChild(tableParts);
        }
        tableParts.AppendChild(new TablePart { Id = tblWorksheet.GetIdOfPart(tableDefPart) });
        tableParts.Count = (uint)tableParts.Elements<TablePart>().Count();
        SaveWorksheet(tblWorksheet);

        var tblIdx = tblWorksheet.TableDefinitionParts.ToList().IndexOf(tableDefPart) + 1;
        return $"/{tblSheetName}/table[{tblIdx}]";
    }

    private string AddPivotTable(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        var ptSegments = parentPath.TrimStart('/').Split('/', 2);
        var ptSheetName = ptSegments[0];
        var ptWorksheet = FindWorksheet(ptSheetName)
            ?? throw new ArgumentException($"Sheet not found: {ptSheetName}");

        // Source: "Sheet1!A1:D100" or "A1:D100" (same sheet)
        var sourceSpec = properties.GetValueOrDefault("source", "")
            ?? properties.GetValueOrDefault("src", "")
            ?? throw new ArgumentException("pivottable requires 'source' property (e.g. source=Sheet1!A1:D100)");
        if (string.IsNullOrEmpty(sourceSpec))
            throw new ArgumentException("pivottable requires 'source' property (e.g. source=Sheet1!A1:D100)");

        // R8-7: incidental whitespace around the source spec or its
        // components (" Sheet1 ! A1:D10 ") is a common paste-from-docs
        // artefact. Trim the whole string and both sides of the '!'
        // split so the downstream sheet/range lookup sees clean values.
        sourceSpec = sourceSpec.Trim();

        // R8-3: external workbook refs such as [other.xlsx]Sheet1!A1:D10
        // used to fall through to FindWorksheet and surface as the
        // misleading "Source sheet not found: [other.xlsx]Sheet1".
        // Detect the '[' prefix up front and throw a clear error so
        // users know the feature is not supported rather than blaming
        // a missing sheet.
        if (sourceSpec.StartsWith("["))
            throw new ArgumentException(
                "External workbook references are not supported in pivot source. "
                + "Use a local sheet name (e.g. Sheet1!A1:D10)");

        string sourceSheetName;
        string sourceRef;

        // B6 v2: try resolving structured-table refs (Table1[#All]) and
        // workbook/sheet-scoped defined names (SalesData, Sheet1!SalesData)
        // into an explicit (sheet, range) tuple BEFORE the literal-parse
        // path. Falls through to the literal parser for explicit
        // "Sheet1!A1:C5" specs and any form the resolver doesn't recognize.
        // See PivotTableHelper.Cache.cs ResolvePivotSourceSpec for coverage.
        var resolved = OfficeCli.Core.PivotTableHelper.ResolvePivotSourceSpec(
            _doc.WorkbookPart!, sourceSpec, defaultSheet: ptSheetName);
        if (resolved.HasValue)
        {
            sourceSheetName = resolved.Value.sheet;
            sourceRef = resolved.Value.rangeRef;
        }
        else if (sourceSpec.Contains('!'))
        {
            var srcParts = sourceSpec.Split('!', 2);
            sourceSheetName = srcParts[0].Trim().Trim('\'', '"').Trim();
            sourceRef = srcParts[1].Trim();
        }
        else
        {
            sourceSheetName = ptSheetName;
            sourceRef = sourceSpec;
        }

        var sourceWorksheet = FindWorksheet(sourceSheetName)
            ?? throw new ArgumentException($"Source sheet not found: {sourceSheetName}");

        var ptPosition = (properties.GetValueOrDefault("position", "")
            ?? properties.GetValueOrDefault("pos", ""))
            ?.Replace("$", ""); // CONSISTENCY(dollar-strip): parity with source ref handling
        if (string.IsNullOrEmpty(ptPosition))
        {
            // Auto-position: place after the source data range
            var rangeEnd = sourceRef.Split(':').Last();
            var colEndMatch = System.Text.RegularExpressions.Regex.Match(rangeEnd, @"([A-Za-z]+)");
            var nextCol = colEndMatch.Success ? IndexToColumnName(ColumnNameToIndex(colEndMatch.Value.ToUpperInvariant()) + 2) : "H";
            ptPosition = $"{nextCol}1";
        }

        // R26-1: validate that the pivot output fits within sheet dimensions
        // before writing any cache/pivot parts. A position near the sheet edge
        // can produce an end-location beyond XFD1048576, which causes a
        // partial-write: cache parts are already saved when the render stage
        // discovers the overflow and throws, leaving a corrupt zip.
        {
            const int ExcelMaxCol = 16384; // XFD
            const int ExcelMaxRow = 1048576;
            var srcRefParts = sourceRef.Replace("$", "").Split(':');
            if (srcRefParts.Length == 2)
            {
                var (srcStartCol, srcStartRow) = ParseCellReference(srcRefParts[0].Trim().ToUpperInvariant());
                var (srcEndCol, srcEndRow)     = ParseCellReference(srcRefParts[1].Trim().ToUpperInvariant());
                int nSourceCols = ColumnNameToIndex(srcEndCol) - ColumnNameToIndex(srcStartCol) + 1;
                int nDataRows   = srcEndRow - srcStartRow; // header excluded
                var (anchorColStr, anchorRow) = ParseCellReference(ptPosition.ToUpperInvariant());
                int anchorColIdx = ColumnNameToIndex(anchorColStr);
                // Conservative lower-bound: pivot needs at least nSourceCols columns
                // (row-label cols + value cols + grand-total col) and at least
                // nDataRows + 2 rows (header + data rows + grand-total row).
                int minEndColIdx = anchorColIdx + nSourceCols - 1;
                int minEndRow    = anchorRow + nDataRows + 1;
                if (minEndColIdx > ExcelMaxCol || minEndRow > ExcelMaxRow)
                {
                    throw new ArgumentException(
                        $"pivot at {ptPosition} does not fit: computed end col={minEndColIdx} row={minEndRow} exceeds sheet dimensions (max XFD1048576)");
                }

                // CONSISTENCY(pivot-output-overlap): two pivot tables whose
                // <x:location> rectangles overlap on the same sheet make
                // Excel surface a "found a problem" repair dialog because
                // the output cells fight for ownership. Mirror the T4
                // table-table overlap check using the conservative output
                // bounds computed above. Cross-sheet pivots are fine.
                var newPivotRange = $"{IndexToColumnName(anchorColIdx)}{anchorRow}:" +
                                    $"{IndexToColumnName(minEndColIdx)}{minEndRow}";
                foreach (var existingPivot in ptWorksheet.PivotTableParts
                    .Select(ptp => ptp.PivotTableDefinition)
                    .Where(d => d != null))
                {
                    var existingLoc = existingPivot!.Location?.Reference?.Value;
                    if (string.IsNullOrEmpty(existingLoc)) continue;
                    if (RangesOverlap(newPivotRange.ToUpperInvariant(), existingLoc.ToUpperInvariant()))
                        throw new ArgumentException(
                            $"Pivot output range overlaps existing pivot " +
                            $"'{existingPivot.Name?.Value}' at {existingLoc}; " +
                            $"choose a different anchor (--prop position=...).");
                }
            }
        }

        // CONSISTENCY(tracking-rebind): CreatePivotTable internally rebinds
        // `properties` to a fresh non-tracking dictionary via
        // NormalizePivotProperties, so all subsequent TryGetValue calls
        // would never reach our TrackingPropertyDictionary comparer. Mark
        // every input key whose normalized form is a known pivot property
        // as consumed up-front, so they don't surface as false
        // unsupported_property warnings. Keys the helper genuinely doesn't
        // know about are still flagged via WarnUnknownPivotProperties +
        // CollectUnknownPivotKeys (R12-1).
        if (properties is OfficeCli.Core.TrackingPropertyDictionary ptTracking)
        {
            var consumed = properties.Keys
                .Where(k => OfficeCli.Core.PivotTableHelper.IsKnownPivotProperty(k))
                .ToList();
            ptTracking.MarkAllConsumed(consumed);
        }

        var ptIdx = PivotTableHelper.CreatePivotTable(
            _doc.WorkbookPart!, ptWorksheet, sourceWorksheet,
            sourceSheetName, sourceRef, ptPosition, properties);

        return $"/{ptSheetName}/pivottable[{ptIdx}]";
    }

}
