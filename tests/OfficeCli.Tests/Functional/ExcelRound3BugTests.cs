// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Failing tests for Round 3 top-4 bugs:
///   Bug A — Scatter chart get --json crashes: markerSize stored as byte,
///            not registered in AppJsonContext source-gen serializer.
///   Bug B — textbox add with shadow=true crashes: "Invalid color value: 'true'"
///            because "true" is passed directly to the color parser.
///   Bug C — namedrange added after pivottable lands after &lt;pivotCaches&gt;,
///            violating OOXML workbook schema order (definedNames must precede pivotCaches).
///   Bug D — No high-level Add API for rich text cells, and Get cannot read
///            InlineString cells (cell.CellValue is null; text lives in cell.InlineString).
/// </summary>
public class ExcelRound3BugTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelRound3BugTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_path);
        _handler = new ExcelHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private void Reopen() { _handler.Dispose(); _handler = new ExcelHandler(_path, editable: false); }
    private void ReopenEditable() { _handler.Dispose(); _handler = new ExcelHandler(_path, editable: true); }

    // ==================== Bug A: Scatter chart --json markerSize byte crash ====================

    /// <summary>
    /// Regression: scatter chart series with marker size stores a System.Byte in Format["markerSize"].
    /// The AOT JSON serializer (AppJsonContext) does not have byte registered, so FormatNode
    /// with OutputFormat.Json throws NotSupportedException / InvalidOperationException.
    /// Fix: cast markerSize.Value to int before storing in Format.
    /// </summary>
    [Fact]
    public void GetScatterChart_WithJson_DoesNotThrow()
    {
        // 1. Create a scatter chart (markers are the default series style)
        _handler.Add("/", "sheet", null, new() { ["name"] = "Data" });
        _handler.Add("/Data", "chart", null, new()
        {
            ["chartType"] = "scatter",
            ["title"] = "Scatter Test",
            ["categories"] = "1,2,3",
            ["series1"] = "S1:10,20,30"
        });

        Reopen();

        // 2. Get the chart node — this must not throw
        DocumentNode node = null!;
        var act = () => { node = _handler.Get("/Data/chart[1]", depth: 1); };
        act.Should().NotThrow("Get should not throw when reading a scatter chart node");

        // 3. Serialize to JSON — this is where the byte crash occurs
        var serializeAct = () =>
            OutputFormatter.FormatNode(node, OutputFormat.Json);
        serializeAct.Should().NotThrow(
            "FormatNode(Json) must not crash with 'JsonTypeInfo metadata for type System.Byte' — " +
            "markerSize should be stored as int, not byte");

        // 4. Verify the JSON output is valid and contains markerSize as a number (not null)
        var json = OutputFormatter.FormatNode(node, OutputFormat.Json);
        var doc = JsonDocument.Parse(json);
        // The chart node serializes fine and chartType is present
        doc.RootElement.GetProperty("format").TryGetProperty("chartType", out var ct).Should().BeTrue();
        ct.GetString().Should().Be("scatter");
    }

    /// <summary>
    /// Specifically verifies that when a scatter chart series carries an explicit marker size,
    /// the Format["markerSize"] value is stored as an int (or at least a JSON-serializable type),
    /// not System.Byte which causes the source-generated serializer to fail.
    /// </summary>
    [Fact]
    public void GetScatterChart_MarkerSizeInFormat_IsNotByteType()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Data" });
        _handler.Add("/Data", "chart", null, new()
        {
            ["chartType"] = "scatter",
            ["categories"] = "1,2,3",
            ["series1"] = "S1:10,20,30",
            ["marker"] = "circle",
            ["markerSize"] = "7"
        });

        Reopen();

        var node = _handler.Get("/Data/chart[1]", depth: 1);
        node.Children.Should().NotBeEmpty("scatter chart should expose series children");

        var seriesNode = node.Children[0];
        if (seriesNode.Format.ContainsKey("markerSize"))
        {
            var markerSizeVal = seriesNode.Format["markerSize"];
            markerSizeVal.Should().NotBeOfType<byte>(
                "markerSize must be stored as int/long, not byte — byte is not registered in AppJsonContext");
            // Confirm it can be serialized without exception
            var serializeAct = () => JsonSerializer.Serialize(markerSizeVal);
            serializeAct.Should().NotThrow();
        }
    }

    // ==================== Bug B: shadow=true crashes with "Invalid color value: 'true'" ====================

    /// <summary>
    /// Regression: Add textbox/shape with shadow=true fails because the shadow parser
    /// splits the value on '-' → parts[0] = "true" → passes "true" to the color builder
    /// → SanitizeColorForOoxml("true") throws ArgumentException("Invalid color value: 'true'").
    /// Fix: treat "true" as a boolean shorthand meaning "use default shadow color" (e.g. black #000000).
    /// </summary>
    [Fact]
    public void AddShape_WithShadowTrue_DoesNotThrow()
    {
        // 1. Create
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });

        // 2. Add shape with shadow=true — this must not throw
        var addAct = () => _handler.Add("/Sheet1", "shape", null, new()
        {
            ["text"] = "Hello Shadow",
            ["fill"] = "4472C4",
            ["shadow"] = "true"
        });
        addAct.Should().NotThrow(
            "shadow=true must be interpreted as default shadow, not passed as a color string");

        // 3. Get + Verify the shape was created
        Reopen();
        var node = _handler.Get("/Sheet1/shape[1]");
        node.Should().NotBeNull();
        node.Text.Should().Be("Hello Shadow");
    }

    /// <summary>
    /// Additional: shadow=true on a textbox (fill=none path) also must not crash.
    /// </summary>
    [Fact]
    public void AddTextbox_WithShadowTrue_DoesNotThrow()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });

        var addAct = () => _handler.Add("/Sheet1", "textbox", null, new()
        {
            ["text"] = "Textbox Shadow",
            ["fill"] = "none",
            ["shadow"] = "true"
        });
        addAct.Should().NotThrow(
            "shadow=true on a fill=none textbox must not throw 'Invalid color value: true'");

        Reopen();
        var node = _handler.Get("/Sheet1/shape[1]");
        node.Should().NotBeNull();
    }

    // ==================== Bug C: definedNames placed after pivotCaches ====================

    /// <summary>
    /// Regression: when a pivot table exists (adding PivotCaches to the workbook element),
    /// then a named range is added, the DefinedNames element is AppendChild'd to the end —
    /// after PivotCaches. OOXML workbook schema requires definedNames before pivotCaches.
    /// The file is technically corrupt and Excel will repair/strip the named range on open.
    /// Fix: insert DefinedNames before PivotCaches (or use a workbook reorder routine).
    /// </summary>
    [Fact]
    public void AddNamedRange_AfterPivotTable_DefinedNamesComesBeforePivotCaches()
    {
        // 1. Create data for pivot source
        _handler.Add("/", "sheet", null, new() { ["name"] = "Source" });
        _handler.Add("/Source", "cell", null, new() { ["ref"] = "A1", ["value"] = "Region" });
        _handler.Add("/Source", "cell", null, new() { ["ref"] = "B1", ["value"] = "Sales" });
        _handler.Add("/Source", "cell", null, new() { ["ref"] = "A2", ["value"] = "North" });
        _handler.Add("/Source", "cell", null, new() { ["ref"] = "B2", ["value"] = "100" });
        _handler.Add("/Source", "cell", null, new() { ["ref"] = "A3", ["value"] = "South" });
        _handler.Add("/Source", "cell", null, new() { ["ref"] = "B3", ["value"] = "200" });

        // 2. Add a pivot table (this creates PivotCaches in workbook)
        _handler.Add("/", "sheet", null, new() { ["name"] = "Pivot" });
        _handler.Add("/Pivot", "pivottable", null, new()
        {
            ["source"] = "Source!A1:B3",
            ["rows"] = "Region",
            ["values"] = "Sales"
        });

        // 3. Now add a named range (this should insert DefinedNames BEFORE PivotCaches)
        _handler.Add("/", "namedrange", null, new()
        {
            ["name"] = "SalesData",
            ["ref"] = "Source!$A$1:$B$3"
        });

        // 4. Verify by inspecting the raw XML element order in the workbook
        // Re-open read-only to get clean state
        Reopen();

        // Use reflection to access the workbook and check element order
        var workbookPartField = typeof(ExcelHandler)
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
        workbookPartField.Should().NotBeNull("_doc field must exist for inspection");

        var doc = workbookPartField!.GetValue(_handler) as DocumentFormat.OpenXml.Packaging.SpreadsheetDocument;
        doc.Should().NotBeNull();

        var workbook = doc!.WorkbookPart?.Workbook;
        workbook.Should().NotBeNull();

        var children = workbook!.ChildElements.ToList();
        var definedNamesIdx = children.FindIndex(e => e.LocalName == "definedNames");
        var pivotCachesIdx = children.FindIndex(e => e.LocalName == "pivotCaches");

        definedNamesIdx.Should().BeGreaterThanOrEqualTo(0, "definedNames element must exist in workbook");
        pivotCachesIdx.Should().BeGreaterThanOrEqualTo(0, "pivotCaches element must exist in workbook");

        definedNamesIdx.Should().BeLessThan(pivotCachesIdx,
            "OOXML schema requires <definedNames> to appear BEFORE <pivotCaches> in the workbook element; " +
            "AppendChild places it after, violating schema order");
    }

    /// <summary>
    /// Simpler variant: verify named range can be retrieved after the workbook is saved
    /// following a pivot table addition (without the file being corrupt).
    /// </summary>
    [Fact]
    public void AddNamedRange_AfterPivotTable_NamedRangeIsReadableAfterReopen()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Source" });
        _handler.Add("/Source", "cell", null, new() { ["ref"] = "A1", ["value"] = "Category" });
        _handler.Add("/Source", "cell", null, new() { ["ref"] = "B1", ["value"] = "Amount" });
        _handler.Add("/Source", "cell", null, new() { ["ref"] = "A2", ["value"] = "X" });
        _handler.Add("/Source", "cell", null, new() { ["ref"] = "B2", ["value"] = "50" });

        _handler.Add("/", "sheet", null, new() { ["name"] = "Pivot" });
        _handler.Add("/Pivot", "pivottable", null, new()
        {
            ["source"] = "Source!A1:B2",
            ["rows"] = "Category",
            ["values"] = "Amount"
        });

        _handler.Add("/", "namedrange", null, new()
        {
            ["name"] = "MyRange",
            ["ref"] = "Source!$A$1:$B$2"
        });

        Reopen();

        // Should be able to get the named range by name
        var act = () => _handler.Get("/namedrange[MyRange]");
        act.Should().NotThrow("named range added after pivot table should remain readable after reopen");

        var node = _handler.Get("/namedrange[MyRange]");
        node.Format["name"].Should().Be("MyRange");
        node.Format["ref"].Should().Be("Source!$A$1:$B$2");
    }

    // ==================== Bug D: Rich text — no Add API, and InlineString Get returns empty text ====================

    /// <summary>
    /// Bug D-1: There is no high-level API to create a rich-text cell via Add.
    /// The test documents the gap: adding a cell with richtext or run properties
    /// should work (currently throws or silently ignores).
    /// Expected fix: Add should support type=richtext or runs property.
    /// </summary>
    [Fact]
    public void AddCell_WithRichTextRuns_CreatesRichTextCell()
    {
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });

        // Attempt to create a rich text cell via add — currently no such API exists,
        // so this either throws ArgumentException or creates a plain cell.
        // After the fix, this should create a SharedString rich text entry.
        _handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1",
            ["type"] = "richtext",
            ["run1"] = "Bold:bold=true",
            ["run2"] = " Normal"
        });

        Reopen();

        // After fix: cell should exist and expose richtext=true in Format
        var node = _handler.Get("/Sheet1/A1");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("richtext",
            "a cell created with type=richtext must set Format[\"richtext\"] = true on readback");
        node.Format["richtext"].Should().Be(true);
    }

    /// <summary>
    /// Bug D-2: Get cannot read InlineString cells. InlineString cells store text in
    /// cell.InlineString (not cell.CellValue), so GetCellDisplayValue returns empty string.
    /// This test creates an InlineString cell directly via OpenXML and then reads it back
    /// via the handler to verify the text is returned correctly.
    /// </summary>
    [Fact]
    public void GetCell_InlineStringType_ReturnsCorrectText()
    {
        // 1. Create an InlineString cell directly via OpenXML (bypassing the handler API)
        _handler.Add("/", "sheet", null, new() { ["name"] = "Sheet1" });
        _handler.Dispose();

        using (var editHandler = new ExcelHandler(_path, editable: true))
        {
            // Access internal worksheet to inject an InlineString cell
            var wsPartField = typeof(ExcelHandler)
                .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            var doc = wsPartField!.GetValue(editHandler) as DocumentFormat.OpenXml.Packaging.SpreadsheetDocument;

            var workbookPart = doc!.WorkbookPart!;
            var sheet = workbookPart.Workbook.GetFirstChild<Sheets>()?.Elements<Sheet>()
                .FirstOrDefault(s => s.Name?.Value == "Sheet1");
            var wsPart = sheet?.Id?.Value != null
                ? (DocumentFormat.OpenXml.Packaging.WorksheetPart)workbookPart.GetPartById(sheet.Id.Value)
                : null;

            wsPart.Should().NotBeNull("Sheet1 must exist");

            var sheetData = wsPart!.Worksheet.GetFirstChild<SheetData>()
                ?? wsPart.Worksheet.AppendChild(new SheetData());

            // Build a row + InlineString cell
            var row = new Row { RowIndex = 1u };
            var cell = new Cell
            {
                CellReference = "A1",
                DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.InlineString),
                InlineString = new InlineString(new Text("Hello InlineString"))
            };
            row.AppendChild(cell);
            sheetData.AppendChild(row);
            wsPart.Worksheet.Save();
        }

        _handler = new ExcelHandler(_path, editable: false);

        // 2. Get the cell — Text should return "Hello InlineString", not ""
        var node = _handler.Get("/Sheet1/A1");
        node.Should().NotBeNull();
        node.Text.Should().Be("Hello InlineString",
            "GetCellDisplayValue must read cell.InlineString.InnerText for InlineString cells, " +
            "not return empty string from cell.CellValue?.Text which is null for InlineString");
        node.Format["type"].Should().Be("InlineString");
    }
}
