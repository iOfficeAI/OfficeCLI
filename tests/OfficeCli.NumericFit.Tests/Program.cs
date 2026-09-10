// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using OfficeCli.Handlers;
using W = DocumentFormat.OpenXml.Wordprocessing;

const string NumericOverflowSubtype = "numeric_overflow";
const string GeneralPrecisionSubtype = "general_precision_loss";

var standardPath = Path.Combine(Path.GetTempPath(), $"officecli-numeric-fit-{Guid.NewGuid():N}.xlsx");
var date1904Path = Path.Combine(Path.GetTempPath(), $"officecli-numeric-fit-1904-{Guid.NewGuid():N}.xlsx");
var generalPath = Path.Combine(Path.GetTempPath(), $"officecli-general-precision-{Guid.NewGuid():N}.xlsx");
var stylelessPath = Path.Combine(Path.GetTempPath(), $"officecli-general-styleless-{Guid.NewGuid():N}.xlsx");
var chartDocxPath = Path.Combine(Path.GetTempPath(), $"officecli-tablechart-{Guid.NewGuid():N}.docx");
var chartEdgeDocxPath = Path.Combine(Path.GetTempPath(), $"officecli-tablechart-edge-{Guid.NewGuid():N}.docx");
try
{
    CreateStandardFixture(standardPath);
    VerifyStandardWorkbook(standardPath);

    CreateDate1904Fixture(date1904Path);
    VerifyDate1904Workbook(date1904Path);

    CreateGeneralPrecisionFixture(generalPath);
    VerifyGeneralPrecisionWorkbook(generalPath);

    CreateStylelessFixture(stylelessPath);
    VerifyStylelessWorkbook(stylelessPath);

    Console.WriteLine("XLSX numeric-fit issue tests passed.");

    CreateBlankDocx(chartDocxPath);
    VerifyDocxTableToChart(chartDocxPath);
    Console.WriteLine("TABLE->CHART DOCX sourceTable sugar tests passed.");

    CreateBlankDocx(chartEdgeDocxPath);
    VerifyDocxTableToChartEdgeCases(chartEdgeDocxPath);
    Console.WriteLine("TABLE->CHART DOCX sourceTable edge-case tests passed.");
}
finally
{
    if (File.Exists(standardPath)) File.Delete(standardPath);
    if (File.Exists(date1904Path)) File.Delete(date1904Path);
    if (File.Exists(generalPath)) File.Delete(generalPath);
    if (File.Exists(stylelessPath)) File.Delete(stylelessPath);
    if (File.Exists(chartDocxPath)) File.Delete(chartDocxPath);
    if (File.Exists(chartEdgeDocxPath)) File.Delete(chartEdgeDocxPath);
}

static void VerifyStandardWorkbook(string path)
{
    string[] expectedPaths =
    [
        "/FormulaView/A1", // Show Formulas does not hide an ordinary number
        "/Sheet1/A1",     // explicit narrow number
        "/Sheet1/B1",     // explicit narrow date serial
        "/Sheet1/F1",     // wrapText does not make a number spill
        "/Sheet1/M1",     // evaluated numeric formula
        "/Sheet1/N1",     // ISO t=d date
        "/Sheet1/P1",     // bracketed color numeric format
        "/Sheet1/T1",     // inherited column style
        "/Sheet1/U3",     // row style wins over column style
        "/Sheet1/W1",     // direct cell style wins over column shrinkToFit
        "/Sheet1/X4",     // direct cell style wins over row shrinkToFit
        "/Sheet1/AD1",    // active non-General section is still checked
    ];

    using var handler = new ExcelHandler(path, editable: false);
    var allIssues = handler.ViewAsIssues();
    var numericIssues = allIssues
        .Where(issue => issue.Subtype == NumericOverflowSubtype)
        .ToList();

    AssertPaths(expectedPaths, numericIssues.Select(issue => issue.Path), "default issues scan");

    foreach (var issue in numericIssues)
    {
        Assert(issue.Type == IssueType.Format, $"{issue.Path} should use the Format bucket");
        Assert(issue.Severity == IssueSeverity.Warning, $"{issue.Path} should be a warning");
        Assert(issue.Message.Contains("numeric overflow", StringComparison.Ordinal),
            $"{issue.Path} should describe the rendering defect");
        Assert(issue.Suggestion?.Contains("suggest.width=", StringComparison.Ordinal) == true,
            $"{issue.Path} should carry an actionable width suggestion");
    }
    Assert(numericIssues.Single(issue => issue.Path == "/Sheet1/A1").Message
            .Contains("'1,234,567.89'", StringComparison.Ordinal),
        "numeric finding should use the formatted display value");
    Assert(numericIssues.Single(issue => issue.Path == "/Sheet1/B1").Message
            .Contains("'2024-10-02'", StringComparison.Ordinal),
        "date finding should use the formatted display value");
    Assert(numericIssues.Single(issue => issue.Path == "/Sheet1/N1").Message
            .Contains("'2024-10-02'", StringComparison.Ordinal),
        "ISO date finding should use its absolute date");
    Assert(numericIssues.All(issue => issue.Path is not "/Sheet1/AE1"
            and not "/Sheet1/AF1" and not "/Sheet1/AG1"),
        "non-finite numeric values should not produce width findings");
    Assert(numericIssues.All(issue => issue.Path != "/Sheet1/AH1"),
        "negative elapsed values in a 1900 workbook need a non-width fix");

    var exactIssues = handler.ViewAsIssues(NumericOverflowSubtype);
    AssertPaths(expectedPaths, exactIssues.Select(issue => issue.Path), "exact subtype filter");

    var formatIssues = handler.ViewAsIssues("format")
        .Where(issue => issue.Subtype == NumericOverflowSubtype);
    AssertPaths(expectedPaths, formatIssues.Select(issue => issue.Path), "format bucket filter");

    // The scanner receives only the remaining capacity and stops at the first
    // matching Format finding; the broken defined name cannot consume it.
    var exactLimited = handler.ViewAsIssues(NumericOverflowSubtype, limit: 1);
    AssertPaths(["/Sheet1/A1"], exactLimited.Select(issue => issue.Path), "exact subtype limit");
    var formatLimited = handler.ViewAsIssues("format", limit: 1);
    AssertPaths(["/Sheet1/A1"], formatLimited.Select(issue => issue.Path), "format bucket limit");
}

static void VerifyDate1904Workbook(string path)
{
    using var handler = new ExcelHandler(path, editable: false);
    var issues = handler.ViewAsIssues(NumericOverflowSubtype);
    AssertPaths(
        [
            "/Sheet1/A1", "/Sheet1/C1", "/Sheet1/E1", "/Sheet1/F1",
            "/Sheet1/G1", "/Sheet1/H1", "/Sheet1/I1"
        ],
        issues.Select(issue => issue.Path),
        "1904 date-system scan");
    Assert(issues.Single(issue => issue.Path == "/Sheet1/A1").Message
            .Contains("'1904-01-01'", StringComparison.Ordinal),
        "a stored 1904 serial should use the workbook epoch");
    Assert(issues.All(issue => issue.Path is not "/Sheet1/B1" and not "/Sheet1/D1"),
        "date-formatted formulas in a 1904 workbook should be skipped conservatively");
    Assert(issues.Single(issue => issue.Path == "/Sheet1/F1").Message
            .Contains("'36:00:00'", StringComparison.Ordinal),
        "1904 workbooks must not epoch-shift positive elapsed hours");
    Assert(issues.Single(issue => issue.Path == "/Sheet1/G1").Message
            .Contains("'-36:00:00'", StringComparison.Ordinal),
        "1904 workbooks must not epoch-shift negative elapsed hours");
}

static void CreateStandardFixture(string path)
{
    using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
    stylesPart.Stylesheet = CreateStylesheet();

    var visiblePart = workbookPart.AddNewPart<WorksheetPart>();
    visiblePart.Worksheet = CreateVisibleWorksheet();
    var formulaViewPart = workbookPart.AddNewPart<WorksheetPart>();
    formulaViewPart.Worksheet = CreateFormulaViewWorksheet();
    var hiddenPart = workbookPart.AddNewPart<WorksheetPart>();
    hiddenPart.Worksheet = CreateHiddenWorksheet();

    workbookPart.Workbook.Append(
        new Sheets(
            SheetFor(workbookPart, visiblePart, 1, "Sheet1"),
            SheetFor(workbookPart, formulaViewPart, 2, "FormulaView"),
            SheetFor(workbookPart, hiddenPart, 3, "HiddenSheet", SheetStateValues.Hidden)),
        new DefinedNames(new DefinedName("#REF!") { Name = "BrokenName" }));
    workbookPart.Workbook.Save();
}

static void CreateDate1904Fixture(string path)
{
    using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook(new WorkbookProperties { Date1904 = true });
    var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
    stylesPart.Stylesheet = CreateStylesheet();
    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    worksheetPart.Worksheet = new Worksheet(
        new SheetFormatProperties { DefaultRowHeight = 15D, DefaultColumnWidth = 8.43D },
        new Columns(
            Col(1, 4), Col(2, 4), Col(3, 2), Col(4, 4), Col(5, 2),
            Col(6, 2), Col(7, 2), Col(8, 2), Col(9, 2)),
        new SheetData(new Row(
            NumberCell("A1", "0", 2),
            FormulaCell("B1", "DATE(2024,10,2)", 2),
            FormulaCell("C1", "1234567.89", 1),
            FormulaCell("D1", "A1", 2),
            NumberCell("E1", "0.5", 10),
            NumberCell("F1", "1.5", 11),
            NumberCell("G1", "-1.5", 11),
            NumberCell("H1", "0.000011574074", 12),
            NumberCell("I1", "1234567", 13))
        { RowIndex = 1 }));
    workbookPart.Workbook.Append(new Sheets(
        SheetFor(workbookPart, worksheetPart, 1, "Sheet1")));
    workbookPart.Workbook.Save();
}

static Sheet SheetFor(
    WorkbookPart workbookPart,
    WorksheetPart worksheetPart,
    uint id,
    string name,
    SheetStateValues? state = null)
{
    var sheet = new Sheet
    {
        Id = workbookPart.GetIdOfPart(worksheetPart),
        SheetId = id,
        Name = name
    };
    if (state != null) sheet.State = state.Value;
    return sheet;
}

static Stylesheet CreateStylesheet()
{
    var numberingFormats = new NumberingFormats(
        new NumberingFormat { NumberFormatId = 164, FormatCode = "#,##0.00" },
        new NumberingFormat { NumberFormatId = 165, FormatCode = "yyyy-mm-dd" },
        new NumberingFormat { NumberFormatId = 166, FormatCode = "[Red]General" },
        new NumberingFormat { NumberFormatId = 167, FormatCode = "[Red]#,##0.00" },
        new NumberingFormat { NumberFormatId = 168, FormatCode = "General;0.00" })
    { Count = 5 };

    var fonts = new Fonts(
        new Font(new FontSize { Val = 11D }, new FontName { Val = "Calibri" }))
    { Count = 1 };
    var fills = new Fills(
        new Fill(new PatternFill { PatternType = PatternValues.None }),
        new Fill(new PatternFill { PatternType = PatternValues.Gray125 }))
    { Count = 2 };
    var borders = new Borders(new Border()) { Count = 1 };
    var cellStyleFormats = new CellStyleFormats(
        new CellFormat(),
        new CellFormat
        {
            ApplyAlignment = true,
            Alignment = new Alignment { ShrinkToFit = true }
        })
    { Count = 2 };
    var cellFormats = new CellFormats(
        new CellFormat(),
        new CellFormat { NumberFormatId = 164, FontId = 0, ApplyNumberFormat = true },
        new CellFormat { NumberFormatId = 165, FontId = 0, ApplyNumberFormat = true },
        new CellFormat
        {
            NumberFormatId = 164,
            FontId = 0,
            ApplyNumberFormat = true,
            ApplyAlignment = true,
            Alignment = new Alignment { ShrinkToFit = true }
        },
        new CellFormat
        {
            NumberFormatId = 164,
            FontId = 0,
            ApplyNumberFormat = true,
            ApplyAlignment = true,
            Alignment = new Alignment { WrapText = true }
        },
        new CellFormat { NumberFormatId = 199, FontId = 0, ApplyNumberFormat = true },
        new CellFormat
        {
            NumberFormatId = 164,
            FontId = 0,
            FormatId = 1,
            ApplyNumberFormat = true
        },
        new CellFormat { NumberFormatId = 166, FontId = 0, ApplyNumberFormat = true },
        new CellFormat { NumberFormatId = 167, FontId = 0, ApplyNumberFormat = true },
        new CellFormat { NumberFormatId = 168, FontId = 0, ApplyNumberFormat = true },
        new CellFormat { NumberFormatId = 45, FontId = 0, ApplyNumberFormat = true },
        new CellFormat { NumberFormatId = 46, FontId = 0, ApplyNumberFormat = true },
        new CellFormat { NumberFormatId = 47, FontId = 0, ApplyNumberFormat = true },
        new CellFormat { NumberFormatId = 48, FontId = 0, ApplyNumberFormat = true })
    { Count = 14 };
    var cellStyles = new CellStyles(
        new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 },
        new CellStyle { Name = "ShrinkBase", FormatId = 1 })
    { Count = 2 };

    return new Stylesheet(
        numberingFormats,
        fonts,
        fills,
        borders,
        cellStyleFormats,
        cellFormats,
        cellStyles,
        new DifferentialFormats { Count = 0 },
        new TableStyles { Count = 0 });
}

static Worksheet CreateVisibleWorksheet()
{
    var columns = new Columns(
        Col(1, 2), Col(2, 4), Col(3, 24), Col(4, 2), Col(5, 2), Col(6, 2),
        Col(7, 8), Col(8, 8), Col(9, 2, hidden: true), Col(10, 2),
        Col(11, 2), Col(12, 2), Col(13, 2), Col(14, 4), Col(15, 2), Col(16, 2),
        Col(17, 7), Col(18, 7),
        Col(20, 2, styleIndex: 1), Col(21, 2, styleIndex: 3),
        Col(22, 2, styleIndex: 1), Col(23, 2, styleIndex: 3), Col(24, 2),
        Col(25, 2), Col(26, 2, styleIndex: 3), Col(27, 2, styleIndex: 1), Col(28, 2),
        Col(29, 2), Col(30, 2), Col(31, 2), Col(32, 2), Col(33, 2), Col(34, 2),
        Col(201, 7), Col(202, 7));

    var row1 = new Row(
        NumberCell("A1", "1234567.89", 1),
        NumberCell("B1", "45567", 2),
        NumberCell("C1", "1234567.89", 1),
        InlineTextCell("D1", "1234567.89"),
        NumberCell("E1", "1234567.89", 3),
        NumberCell("F1", "1234567.89", 4),
        NumberCell("G1", "1234567.89", 1),
        NumberCell("H1", "9876543.21", 1),
        NumberCell("I1", "1234567.89", 1),
        NumberCell("K1", "123456789", 0),
        NumberCell("L1", "1234567.89", 5),
        FormulaCell("M1", "1234567.89", 1),
        IsoDateCell("N1", "2024-10-02T00:00:00Z", 2),
        NumberCell("O1", "123456789", 7),
        NumberCell("P1", "-1234567.89", 8),
        NumberCell("T1", "1234567.89"),
        NumberCell("V1", "1234567.89", 3),
        NumberCell("W1", "1234567.89", 1),
        NumberCell("Z1", "1234567.89"),
        NumberCell("AA1", "1234567.89", 0),
        NumberCell("AB1", "1234567.89", 6),
        NumberCell("AC1", "123456789", 9),
        NumberCell("AD1", "-1234567.89", 9),
        NumberCell("AE1", "NaN", 1),
        NumberCell("AF1", "Infinity", 1),
        NumberCell("AG1", "-Infinity", 1),
        NumberCell("AH1", "-1.5", 11),
        NumberCell("GS1", "12345.67", 1),
        NumberCell("GT1", "98765.43", 1))
    { RowIndex = 1 };
    var row2 = new Row(NumberCell("J2", "1234567.89", 1))
    {
        RowIndex = 2,
        Hidden = true
    };
    var row3 = new Row(NumberCell("U3", "1234567.89"))
    {
        RowIndex = 3,
        CustomFormat = true,
        StyleIndex = 1
    };
    var row4 = new Row(NumberCell("X4", "1234567.89", 1))
    {
        RowIndex = 4,
        CustomFormat = true,
        StyleIndex = 3
    };
    var row5 = new Row(NumberCell("Y5", "1234567.89"))
    {
        RowIndex = 5,
        CustomFormat = true,
        StyleIndex = 3
    };
    var row5001 = new Row(
        NumberCell("Q5001", "12345.67", 1),
        NumberCell("R5001", "98765.43", 1))
    { RowIndex = 5001 };

    var sheetData = new SheetData(row1, row2, row3, row4, row5, row5001);
    var mergeCells = new MergeCells(
        new MergeCell { Reference = "G1:H1" },
        new MergeCell { Reference = "Q5001:R5001" },
        new MergeCell { Reference = "GS1:GT1" })
    { Count = 3 };
    return new Worksheet(
        new SheetFormatProperties { DefaultRowHeight = 15D, DefaultColumnWidth = 8.43D },
        columns,
        sheetData,
        mergeCells);
}

static Worksheet CreateFormulaViewWorksheet()
{
    return new Worksheet(
        new SheetViews(new SheetView { WorkbookViewId = 0, ShowFormulas = true }),
        new SheetFormatProperties { DefaultRowHeight = 15D, DefaultColumnWidth = 2D },
        new Columns(Col(1, 2), Col(2, 2)),
        new SheetData(new Row(
            NumberCell("A1", "1234567.89", 1),
            FormulaCell("B1", "1234567.89", 1))
        { RowIndex = 1 }));
}

static Worksheet CreateHiddenWorksheet()
{
    return new Worksheet(
        new SheetFormatProperties { DefaultRowHeight = 15D, DefaultColumnWidth = 2D },
        new Columns(Col(1, 2)),
        new SheetData(new Row(NumberCell("A1", "1234567.89", 1)) { RowIndex = 1 }));
}

static Column Col(uint index, double width, bool hidden = false, uint? styleIndex = null)
{
    var column = new Column
    {
        Min = index,
        Max = index,
        Width = width,
        CustomWidth = true,
        Hidden = hidden
    };
    if (styleIndex.HasValue) column.Style = styleIndex.Value;
    return column;
}

static Cell NumberCell(string reference, string value, uint? styleIndex = null)
{
    var cell = new Cell
    {
        CellReference = reference,
        CellValue = new CellValue(value)
    };
    if (styleIndex.HasValue) cell.StyleIndex = styleIndex.Value;
    return cell;
}

static Cell FormulaCell(string reference, string formula, uint styleIndex) => new()
{
    CellReference = reference,
    StyleIndex = styleIndex,
    CellFormula = new CellFormula(formula)
};

static Cell IsoDateCell(string reference, string value, uint styleIndex) => new()
{
    CellReference = reference,
    StyleIndex = styleIndex,
    DataType = CellValues.Date,
    CellValue = new CellValue(value)
};

static Cell InlineTextCell(string reference, string value) => new()
{
    CellReference = reference,
    StyleIndex = 1,
    DataType = CellValues.InlineString,
    InlineString = new InlineString(new Text(value))
};

static void AssertPaths(IEnumerable<string> expected, IEnumerable<string> actual, string scenario)
{
    var expectedList = expected.OrderBy(path => path, StringComparer.Ordinal).ToArray();
    var actualList = actual.OrderBy(path => path, StringComparer.Ordinal).ToArray();
    if (expectedList.SequenceEqual(actualList, StringComparer.Ordinal)) return;
    throw new InvalidOperationException(
        $"{scenario}: expected [{string.Join(", ", expectedList)}], got [{string.Join(", ", actualList)}]");
}

static void Assert(bool condition, string message)
{
    if (!condition) throw new InvalidOperationException(message);
}

// ── general_precision_loss ────────────────────────────────────────────────
// Excel's General display caps at 11 significant digits and falls back to
// scientific notation past that, REGARDLESS of column width. Verified against
// desktop Excel at width 20: 11 digits render in full, 12+ become 1.23457E+11.

static void CreateGeneralPrecisionFixture(string path)
{
    using var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
    var wbPart = doc.AddWorkbookPart();
    wbPart.Workbook = new Workbook();
    var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
    stylesPart.Stylesheet = new Stylesheet(
        new Fonts(new Font(new FontSize { Val = 11 })) { Count = 1 },
        new Fills(new Fill(new PatternFill { PatternType = PatternValues.None })) { Count = 1 },
        new Borders(new Border()) { Count = 1 },
        new CellFormats(
            new CellFormat(),                                        // 0: General
            new CellFormat { NumberFormatId = 1, ApplyNumberFormat = true }) // 1: "0"
        { Count = 2 });
    stylesPart.Stylesheet.Save();

    var wsPart = wbPart.AddNewPart<WorksheetPart>();
    var row = new Row { RowIndex = 1 };
    // Under the 11-digit cap: General shows these in full.
    row.Append(NumberCell("A1", "1234567890"));        // 10 digits
    row.Append(NumberCell("B1", "12345678901"));       // 11 digits — boundary, clean
    // Past the cap with a 12+ digit integer part: scientific, no width fixes it.
    row.Append(NumberCell("C1", "123456789012"));      // 12 digits
    row.Append(NumberCell("D1", "1234567890123"));     // 13 digits
    row.Append(NumberCell("E1", "123456789012345"));   // 15 digits — Excel's storage limit
    // Float noise below 1e11: General rounds to 11 significant digits and renders
    // 46551.3 positionally, so these must NOT be reported.
    row.Append(NumberCell("F1", "46551.299999999996"));
    row.Append(NumberCell("G1", "2887816940.375"));
    // Beyond 1e15 no number format can help; scientific is the only sane display,
    // so a finding would be noise.
    row.Append(NumberCell("H1", "5.556714093664954E+31"));
    // An EXPLICIT format is the other scan's territory, never this one.
    row.Append(NumberCell("I1", "1234567890123", 1));
    // Zero and small values are never candidates.
    row.Append(NumberCell("J1", "0"));
    row.Append(NumberCell("K1", "3.14"));
    wsPart.Worksheet = new Worksheet(new SheetData(row));
    wsPart.Worksheet.Save();

    wbPart.Workbook.AppendChild(new Sheets(new Sheet
    {
        Name = "Sheet1",
        SheetId = 1,
        Id = wbPart.GetIdOfPart(wsPart)
    }));
    wbPart.Workbook.Save();
}

static void VerifyGeneralPrecisionWorkbook(string path)
{
    string[] expectedPaths = ["/Sheet1/C1", "/Sheet1/D1", "/Sheet1/E1"];

    using var handler = new ExcelHandler(path, editable: false);
    var issues = handler.ViewAsIssues()
        .Where(issue => issue.Subtype == GeneralPrecisionSubtype)
        .ToList();
    AssertPaths(expectedPaths, issues.Select(issue => issue.Path), "general precision scan");

    foreach (var issue in issues)
    {
        Assert(issue.Type == IssueType.Format, $"{issue.Path} should use the Format bucket");
        Assert(issue.Severity == IssueSeverity.Warning, $"{issue.Path} should be a warning");
        Assert(issue.Message.Contains("General precision loss", StringComparison.Ordinal),
            $"{issue.Path} should name the defect");
        // The remedy is a number format, never a width — that distinction is
        // the whole reason this is a separate family from numeric_overflow.
        var suggestion = issue.Suggestion ?? "";
        Assert(suggestion.Contains("numberFormat", StringComparison.Ordinal),
            $"{issue.Path} should suggest a number format");
        Assert(!suggestion.Contains("suggest.width", StringComparison.Ordinal),
            $"{issue.Path} must not suggest a width");
    }

    // The two families are disjoint: numeric_overflow requires an explicit
    // format, this one requires General, so no cell is reported twice.
    var overflowPaths = handler.ViewAsIssues()
        .Where(issue => issue.Subtype == NumericOverflowSubtype)
        .Select(issue => issue.Path)
        .ToHashSet(StringComparer.Ordinal);
    foreach (var issue in issues)
        Assert(!overflowPaths.Contains(issue.Path),
            $"{issue.Path} reported by both numeric_overflow and general_precision_loss");

    // Exact-subtype and bucket requests reach the same findings.
    AssertPaths(expectedPaths,
        handler.ViewAsIssues(GeneralPrecisionSubtype).Select(issue => issue.Path),
        "general precision exact subtype");
    AssertPaths(expectedPaths,
        handler.ViewAsIssues("format")
            .Where(issue => issue.Subtype == GeneralPrecisionSubtype)
            .Select(issue => issue.Path),
        "general precision format bucket");

    Assert(handler.ViewAsIssues(GeneralPrecisionSubtype, limit: 1).Count == 1,
        "general precision scan should honor --limit");
}

// A workbook with no stylesheet cannot carry a number format, so every cell in
// it is General. Bailing out on a missing stylesheet (as the width-based scan
// must) would blind this check to exactly that case.
static void CreateStylelessFixture(string path)
{
    using var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
    var wbPart = doc.AddWorkbookPart();
    wbPart.Workbook = new Workbook();
    var wsPart = wbPart.AddNewPart<WorksheetPart>();
    var row = new Row { RowIndex = 1 };
    row.Append(NumberCell("A1", "1234567890123"));
    row.Append(NumberCell("B1", "42"));
    wsPart.Worksheet = new Worksheet(new SheetData(row));
    wsPart.Worksheet.Save();
    wbPart.Workbook.AppendChild(new Sheets(new Sheet
    {
        Name = "Sheet1",
        SheetId = 1,
        Id = wbPart.GetIdOfPart(wsPart)
    }));
    wbPart.Workbook.Save();
}

static void VerifyStylelessWorkbook(string path)
{
    using var handler = new ExcelHandler(path, editable: false);
    AssertPaths(["/Sheet1/A1"],
        handler.ViewAsIssues()
            .Where(issue => issue.Subtype == GeneralPrecisionSubtype)
            .Select(issue => issue.Path),
        "styleless workbook is all-General");
}

static void CreateBlankDocx(string path)
{
    using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var mainPart = doc.AddMainDocumentPart();
    mainPart.Document = new W.Document(new W.Body(new W.Paragraph()));
    mainPart.Document.Save();
}

static void VerifyDocxTableToChart(string path)
{
    // Create a table with a header row (Quarter) + one numeric column (Revenue),
    // then a chart whose categories/series come from the table via dataTable=.
    string chartPath;
    using (var handler = new WordHandler(path, editable: true))
    {
        // Table `data=` parse grid: ',' per cell, ';' per row (DelimitedText
        // falls back to `,`/`;` because the value is not a resolvable file src).
        handler.Add("/body", "table", null, new Dictionary<string, string>
        {
            ["data"] = "Quarter,Revenue;Q1,100;Q2,150;Q3,200;Q4,250",
            ["header"] = "true",
        });
        // Windowed: blank seed paragraph precedes the table, so /body/tbl[1] is it.
        chartPath = handler.Add("/body", "chart", null, new Dictionary<string, string>
        {
            ["sourceTable"] = "/body/tbl[1]",
            ["chartType"] = "column",
        });
        Assert(chartPath.StartsWith("/chart["), "sourceTable chart should resolve to a /chart[N] path, got " + chartPath);
    }

    using var doc = WordprocessingDocument.Open(path, false);
    // The chart part must be physically present (the dataTable sugar produced a
    // real chart, not a silent no-op).
    Assert(doc.MainDocumentPart!.ChartParts.Any(),
        "dataTable chart should create a ChartPart in the package");

    // Full grid must survive: the old `ri < rows[ri].Count` guard dropped every
    // data row past the first (only Q1/100 reached the chart).
    using (var handler = new WordHandler(path, editable: false))
    {
        var node = handler.Query("chart").First(n => n.Path == chartPath);
        Assert(node.Format.TryGetValue("series1", out var s1) && s1?.ToString() == "Revenue:100,150,200,250",
            "header table should chart all four quarters, got series1=" + (s1?.ToString() ?? "<none>"));
        Assert(node.Format.TryGetValue("categories", out var cats) && cats?.ToString() == "Q1,Q2,Q3,Q4",
            "header table categories should be Q1..Q4, got " + (cats?.ToString() ?? "<none>"));
    }
}

static void VerifyDocxTableToChartEdgeCases(string path)
{
    using (var handler = new WordHandler(path, editable: true))
    {
        // 1. Ragged table: a short row must not throw IndexOutOfRangeException.
        handler.Add("/body", "table", null, new Dictionary<string, string>
        {
            ["data"] = "Item,Sales;OnlyLabel;B,200",
        });
        var raggedChart = handler.Add("/body", "chart", null, new Dictionary<string, string>
        {
            ["sourceTable"] = "/body/tbl[1]",
            ["chartType"] = "column",
        });
        Assert(raggedChart.StartsWith("/chart["), "ragged-table chart should resolve, got " + raggedChart);

        // 2. Headerless label-first-column table: "Q1|100" rows are DATA, not a
        // header — col 0 becomes categories, the numeric column the series.
        handler.Add("/body", "table", null, new Dictionary<string, string>
        {
            ["data"] = "Q1,100;Q2,150;Q3,200",
        });
        var bareChart = handler.Add("/body", "chart", null, new Dictionary<string, string>
        {
            ["sourceTable"] = "/body/tbl[2]",
            ["chartType"] = "column",
        });
        Assert(bareChart.StartsWith("/chart["), "headerless chart should resolve, got " + bareChart);

        // 3. Stray non-numeric cell: the bad row drops whole (label included) so
        // categories can never drift out of alignment with values.
        handler.Add("/body", "table", null, new Dictionary<string, string>
        {
            ["data"] = "Item,Sales;A,10;B,N/A;C,30",
        });
        var gapChart = handler.Add("/body", "chart", null, new Dictionary<string, string>
        {
            ["sourceTable"] = "/body/tbl[3]",
            ["chartType"] = "column",
        });
        Assert(gapChart.StartsWith("/chart["), "gap-table chart should resolve, got " + gapChart);

        var charts = handler.Query("chart").ToDictionary(n => n.Path);
        var ragged = charts[raggedChart];
        Assert(ragged.Format.TryGetValue("series1", out var rs) && rs?.ToString() == "Sales:200",
            "ragged table should chart only the complete row, got series1=" + (rs?.ToString() ?? "<none>"));
        Assert(ragged.Format.TryGetValue("categories", out var rc) && rc?.ToString() == "B",
            "ragged table categories should be just B, got " + (rc?.ToString() ?? "<none>"));

        var bare = charts[bareChart];
        Assert(bare.Format.TryGetValue("series1", out var bs) && bs?.ToString() == "Series 1:100,150,200",
            "headerless table should keep all rows as data, got series1=" + (bs?.ToString() ?? "<none>"));
        Assert(bare.Format.TryGetValue("categories", out var bc) && bc?.ToString() == "Q1,Q2,Q3",
            "headerless col 0 should become categories, got " + (bc?.ToString() ?? "<none>"));

        var gap = charts[gapChart];
        Assert(gap.Format.TryGetValue("series1", out var gs) && gs?.ToString() == "Sales:10,30",
            "N/A row should drop its values, got series1=" + (gs?.ToString() ?? "<none>"));
        Assert(gap.Format.TryGetValue("categories", out var gc) && gc?.ToString() == "A,C",
            "N/A row should drop its label too (alignment), got " + (gc?.ToString() ?? "<none>"));
    }

    using var doc = WordprocessingDocument.Open(path, false);
    Assert(doc.MainDocumentPart!.ChartParts.Count() == 3,
        "edge-case docx should contain three ChartParts");
}
