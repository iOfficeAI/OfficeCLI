// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Core;

/// <summary>
/// Keeps a pptx / docx chart's embedded workbook in step with a data Set.
///
/// A chart that came from PowerPoint or Word stores its data twice: the
/// series reference cells of an embedded workbook (<c:numRef>/<c:strRef> with
/// a <c:f> formula) and a cache of those cells' values inside the chart part.
/// The shared chart setter writes new series data as LITERALS (numLit/strLit/
/// a bare c:v name) — right for an xlsx chart, whose references point at the
/// live sheet, but on a chart with an embedded workbook it silently cut the
/// series off its data: the slide showed the new values while "Edit Data"
/// opened the workbook with the old ones (issue #452).
///
/// Around such a Set, <see cref="Capture"/> records every series reference;
/// <see cref="Restore"/> then turns each element the Set rewrote as a literal
/// back into a reference to the same cells (range resized to the new point
/// count), with the literal as its cache, and writes the same values into the
/// embedded workbook. Charts without an embedded workbook, and references
/// that are not a plain single-row / single-column range, keep the literal
/// behaviour unchanged.
/// </summary>
public static class ChartEmbeddedDataSync
{
    public sealed class Snapshot
    {
        internal readonly Dictionary<(int Series, string Kind), string> Formulas = new();
        internal int SeriesCount;
    }

    private static readonly string[] Kinds = { "tx", "cat", "val", "xVal", "yVal" };

    public static Snapshot? Capture(ChartPart chartPart)
    {
        if (FindEmbeddedWorkbook(chartPart) == null) return null;
        var snap = new Snapshot();
        var series = AllSeries(chartPart);
        snap.SeriesCount = series.Count;
        for (int i = 0; i < series.Count; i++)
            foreach (var kind in Kinds)
                if (Slot(series[i], kind) is { } slot
                    && slot.Descendants<C.Formula>().FirstOrDefault()?.Text is { Length: > 0 } f)
                    snap.Formulas[(i, kind)] = f;
        return snap.Formulas.Count == 0 ? null : snap;
    }

    public static void Restore(ChartPart chartPart, Snapshot? snap)
    {
        if (snap == null) return;
        var embedded = FindEmbeddedWorkbook(chartPart);
        if (embedded == null) return;

        var series = AllSeries(chartPart);
        bool changedSeries = series.Count != snap.SeriesCount;
        if (changedSeries)
        {
            // Only literals need new backing; preserve the current cell references.
            using var source = embedded.GetStream(FileMode.Open, FileAccess.Read);
            using var workbook = SpreadsheetDocument.Open(source, false);
            if (workbook.WorkbookPart is not { } workbookPart) return;
            var sheet = workbookPart.Workbook?.Sheets?.Elements<Sheet>().FirstOrDefault();
            if (sheet?.Id?.Value is not { } sheetId || sheet.Name?.Value is not { } sheetName
                || workbookPart.GetPartById(sheetId) is not WorksheetPart worksheet
                || worksheet.Worksheet is not { } sheetRoot) return;
            int column = sheetRoot.Descendants<Cell>()
                .Select(cell => ColIndex(new string((cell.CellReference?.Value ?? "").TakeWhile(char.IsLetter).ToArray())))
                .DefaultIfEmpty(0).Max() + 1;
            snap = new Snapshot { SeriesCount = series.Count };
            for (int i = 0; i < series.Count; i++)
                foreach (var kind in Kinds)
                    if (Slot(series[i], kind) is { } slot && !slot.Descendants<C.Formula>().Any()
                        && ReadLiteral(slot, kind).Values is { Count: > 0 })
                        snap.Formulas[(i, kind)] = $"{ModernFunctionQualifier.QuoteSheetNameForRef(sheetName)}!${ColName(column++)}$1";
        }
        var writes = new List<(string Sheet, int Col, int Row, string Value, bool IsNumber)>();
        foreach (var ((si, kind), oldFormula) in snap.Formulas)
        {
            if (si >= series.Count || Slot(series[si], kind) is not { } slot) continue;
            if (slot.Descendants<C.Formula>().Any()) continue; // still a reference — untouched

            var (values, isNumber, formatCode) = ReadLiteral(slot, kind);
            if (values == null) continue;
            if (!TryParseRange(oldFormula, values.Count, out var sheet, out var cells, out var newFormula))
                continue; // not a plain vector range: keep the literal

            slot.RemoveAllChildren();
            if (isNumber)
            {
                var cache = new C.NumberingCache(new C.FormatCode(formatCode ?? "General"),
                    new C.PointCount { Val = (uint)values.Count });
                for (int i = 0; i < values.Count; i++)
                    cache.AppendChild(new C.NumericPoint(new C.NumericValue(values[i])) { Index = (uint)i });
                slot.AppendChild(new C.NumberReference(new C.Formula(newFormula), cache));
            }
            else
            {
                var cache = new C.StringCache(new C.PointCount { Val = (uint)values.Count });
                for (int i = 0; i < values.Count; i++)
                    cache.AppendChild(new C.StringPoint(new C.NumericValue(values[i])) { Index = (uint)i });
                slot.AppendChild(new C.StringReference(new C.Formula(newFormula), cache));
            }
            for (int i = 0; i < values.Count; i++)
                writes.Add((sheet, cells[i].Col, cells[i].Row, values[i], isNumber));
        }
        if (writes.Count == 0) return;

        chartPart.ChartSpace!.Save();
        WriteCells(embedded, writes);
    }

    // ---------------------------------------------------------------- chart side

    private static List<OpenXmlCompositeElement> AllSeries(ChartPart chartPart) =>
        chartPart.ChartSpace?.GetFirstChild<C.Chart>()?.GetFirstChild<C.PlotArea>()?
            .Descendants<OpenXmlCompositeElement>().Where(e => e.LocalName == "ser").ToList()
        ?? new List<OpenXmlCompositeElement>();

    private static OpenXmlCompositeElement? Slot(OpenXmlCompositeElement ser, string kind) =>
        ser.Elements<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == kind);

    private static (List<string>? Values, bool IsNumber, string? FormatCode) ReadLiteral(
        OpenXmlCompositeElement slot, string kind)
    {
        if (kind == "tx")
        {
            // A literal series name is a bare <c:v> under <c:tx>.
            var v = slot.GetFirstChild<C.NumericValue>()?.Text;
            return v == null ? (null, false, null) : (new List<string> { v }, false, null);
        }
        if (slot.GetFirstChild<C.NumberLiteral>() is { } numLit)
            return (Points(numLit.Elements<C.NumericPoint>().Select(p => (p.Index?.Value ?? 0u, p.NumericValue?.Text ?? "")),
                        numLit.PointCount?.Val?.Value),
                    true, numLit.FormatCode?.Text);
        if (slot.GetFirstChild<C.StringLiteral>() is { } strLit)
            return (Points(strLit.Elements<C.StringPoint>().Select(p => (p.Index?.Value ?? 0u, p.NumericValue?.Text ?? "")),
                        strLit.PointCount?.Val?.Value),
                    false, null);
        return (null, false, null);
    }

    private static List<string> Points(IEnumerable<(uint Idx, string V)> pts, uint? count)
    {
        var list = pts.ToList();
        var n = (int)(count ?? (list.Count == 0 ? 0u : list.Max(p => p.Idx) + 1));
        var values = Enumerable.Repeat("", n).ToList();
        foreach (var (idx, v) in list)
            if (idx < n) values[(int)idx] = v;
        return values;
    }

    // "Sheet1!$B$2:$B$4" / "'My Sheet'!$A$2:$A$4" / "Sheet1!$B$1"
    private static readonly Regex RangeRx = new(
        @"^(?<sheet>'(?:[^']|'')+'|[^!']+)!\$?(?<c1>[A-Za-z]{1,3})\$?(?<r1>\d+)(?::\$?(?<c2>[A-Za-z]{1,3})\$?(?<r2>\d+))?$",
        RegexOptions.Compiled);

    private static bool TryParseRange(string formula, int count, out string sheet,
        out List<(int Col, int Row)> cells, out string newFormula)
    {
        sheet = ""; cells = new(); newFormula = formula;
        var m = RangeRx.Match(formula.Trim());
        if (!m.Success || count <= 0) return false;
        var sheetToken = m.Groups["sheet"].Value;
        sheet = sheetToken.StartsWith('\'') ? sheetToken[1..^1].Replace("''", "'") : sheetToken;
        int c1 = ColIndex(m.Groups["c1"].Value), r1 = int.Parse(m.Groups["r1"].Value, CultureInfo.InvariantCulture);
        int c2 = m.Groups["c2"].Success ? ColIndex(m.Groups["c2"].Value) : c1;
        int r2 = m.Groups["r2"].Success ? int.Parse(m.Groups["r2"].Value, CultureInfo.InvariantCulture) : r1;
        bool vertical;
        if (c1 == c2 && r1 == r2) vertical = true;       // single cell: grow downward
        else if (c1 == c2) vertical = true;
        else if (r1 == r2) vertical = false;
        else return false;                               // 2-D block: not a series vector
        for (int i = 0; i < count; i++)
            cells.Add(vertical ? (c1, r1 + i) : (c1 + i, r1));
        var (ec, er) = cells[^1];
        var start = $"${ColName(c1)}${r1}";
        var end = $"${ColName(ec)}${er}";
        newFormula = $"{sheetToken}!{start}" + (count == 1 && !m.Groups["c2"].Success ? "" : $":{end}");
        return true;
    }

    private static int ColIndex(string col)
    {
        int n = 0;
        foreach (var ch in col.ToUpperInvariant()) n = n * 26 + (ch - 'A' + 1);
        return n;
    }

    private static string ColName(int idx)
    {
        var s = "";
        while (idx > 0) { idx--; s = (char)('A' + idx % 26) + s; idx /= 26; }
        return s;
    }

    // ---------------------------------------------------------- workbook side

    private static EmbeddedPackagePart? FindEmbeddedWorkbook(ChartPart chartPart)
    {
        var relId = chartPart.ChartSpace?.GetFirstChild<C.ExternalData>()?.Id?.Value;
        if (string.IsNullOrEmpty(relId)) return null;
        try
        {
            return chartPart.GetPartById(relId) is EmbeddedPackagePart p
                   && p.ContentType.Contains("spreadsheetml", StringComparison.OrdinalIgnoreCase)
                ? p : null;
        }
        catch (ArgumentOutOfRangeException) { return null; }
    }

    private static void WriteCells(EmbeddedPackagePart part,
        List<(string Sheet, int Col, int Row, string Value, bool IsNumber)> writes)
    {
        using var ms = new MemoryStream();
        using (var src = part.GetStream(FileMode.Open, FileAccess.Read)) src.CopyTo(ms);
        ms.Position = 0;
        using (var wb = SpreadsheetDocument.Open(ms, true))
        {
            var wbPart = wb.WorkbookPart;
            if (wbPart?.Workbook?.Sheets == null) return;
            foreach (var group in writes.GroupBy(w => w.Sheet, StringComparer.OrdinalIgnoreCase))
            {
                var sheet = wbPart.Workbook.Sheets.Elements<Sheet>()
                    .FirstOrDefault(s => string.Equals(s.Name?.Value, group.Key, StringComparison.OrdinalIgnoreCase));
                if (sheet?.Id?.Value is not { } sid || wbPart.GetPartById(sid) is not WorksheetPart wsPart) continue;
                var sheetData = wsPart.Worksheet?.GetFirstChild<SheetData>();
                if (sheetData == null) continue;
                foreach (var w in group) SetCell(sheetData, w.Col, w.Row, w.Value, w.IsNumber);
                wsPart.Worksheet!.Save();
            }
        }
        ms.Position = 0;
        part.FeedData(ms);
    }

    private static void SetCell(SheetData sheetData, int col, int row, string value, bool isNumber)
    {
        var rowEl = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == (uint)row);
        if (rowEl == null)
        {
            rowEl = new Row { RowIndex = (uint)row };
            var after = sheetData.Elements<Row>().LastOrDefault(r => (r.RowIndex?.Value ?? 0) < row);
            if (after != null) after.InsertAfterSelf(rowEl); else sheetData.PrependChild(rowEl);
        }
        var cellRef = ColName(col) + row.ToString(CultureInfo.InvariantCulture);
        var cell = rowEl.Elements<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
        if (cell == null)
        {
            cell = new Cell { CellReference = cellRef };
            var after = rowEl.Elements<Cell>().LastOrDefault(c =>
                c.CellReference?.Value is { } r && ColIndex(new string(r.TakeWhile(char.IsLetter).ToArray())) < col);
            if (after != null) after.InsertAfterSelf(cell); else rowEl.PrependChild(cell);
        }
        cell.RemoveAllChildren();
        if (isNumber && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
        {
            cell.DataType = null;
            cell.AppendChild(new CellValue(value));
        }
        else
        {
            cell.DataType = CellValues.InlineString;
            cell.AppendChild(new InlineString(new Text(value)));
        }
    }
}
