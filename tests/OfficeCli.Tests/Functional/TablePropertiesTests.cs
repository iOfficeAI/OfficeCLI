// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using OfficeCli.Handlers;
using Xunit;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Functional tests for newly added table properties across PPTX, DOCX, and XLSX.
/// </summary>
public class PptxTablePropertiesTests : IDisposable
{
    private readonly string _path;
    private PowerPointHandler _handler;

    public PptxTablePropertiesTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_path);
        _handler = new PowerPointHandler(_path, editable: true);
        _handler.Add("/", "slide", null, new() { ["layout"] = "blank" });
        _handler.Add("/slide[1]", "table", null, new() { ["rows"] = "3", ["cols"] = "3", ["style"] = "medium2" });
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private PowerPointHandler Reopen()
    {
        _handler.Dispose();
        _handler = new PowerPointHandler(_path, editable: true);
        return _handler;
    }

    // ==================== TableLook flags ====================

    [Fact]
    public void Set_FirstRow_IsReadBack()
    {
        _handler.Set("/slide[1]/table[1]", new() { ["firstRow"] = "false" });
        var node = _handler.Get("/slide[1]/table[1]");
        node.Format["firstRow"].Should().Be(false);
    }

    [Fact]
    public void Set_LastRow_IsReadBack()
    {
        _handler.Set("/slide[1]/table[1]", new() { ["lastRow"] = "true" });
        var node = _handler.Get("/slide[1]/table[1]");
        node.Format["lastRow"].Should().Be(true);
    }

    [Fact]
    public void Set_FirstCol_IsReadBack()
    {
        _handler.Set("/slide[1]/table[1]", new() { ["firstCol"] = "true" });
        var node = _handler.Get("/slide[1]/table[1]");
        node.Format["firstCol"].Should().Be(true);
    }

    [Fact]
    public void Set_LastCol_IsReadBack()
    {
        _handler.Set("/slide[1]/table[1]", new() { ["lastCol"] = "true" });
        var node = _handler.Get("/slide[1]/table[1]");
        node.Format["lastCol"].Should().Be(true);
    }

    [Fact]
    public void Set_BandedRows_False_IsReadBack()
    {
        _handler.Set("/slide[1]/table[1]", new() { ["bandedRows"] = "false" });
        var node = _handler.Get("/slide[1]/table[1]");
        node.Format["bandedRows"].Should().Be(false);
    }

    [Fact]
    public void Set_BandedCols_True_IsReadBack()
    {
        _handler.Set("/slide[1]/table[1]", new() { ["bandedCols"] = "true" });
        var node = _handler.Get("/slide[1]/table[1]");
        node.Format["bandedCols"].Should().Be(true);
    }

    [Fact]
    public void Set_TableLookFlags_Persist_SurvivesReopen()
    {
        _handler.Set("/slide[1]/table[1]", new()
        {
            ["firstRow"] = "true", ["lastRow"] = "true",
            ["firstCol"] = "true", ["lastCol"] = "true",
            ["bandedRows"] = "true", ["bandedCols"] = "true"
        });
        Reopen();
        var node = _handler.Get("/slide[1]/table[1]");
        node.Format["firstRow"].Should().Be(true);
        node.Format["lastRow"].Should().Be(true);
        node.Format["firstCol"].Should().Be(true);
        node.Format["lastCol"].Should().Be(true);
        node.Format["bandedRows"].Should().Be(true);
        node.Format["bandedCols"].Should().Be(true);
    }

    // ==================== Column widths ====================

    [Fact]
    public void Set_ColWidths_UpdatesGridColumns()
    {
        _handler.Set("/slide[1]/table[1]", new() { ["colWidths"] = "3cm,5cm,3cm" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var table = doc.PresentationPart!.SlideParts.First().Slide.Descendants<Drawing.Table>().First();
        var gridCols = table.TableGrid!.Elements<Drawing.GridColumn>().ToList();
        gridCols.Should().HaveCount(3);
        gridCols[0].Width!.Value.Should().BeInRange(1070000, 1090000); // 3cm ≈ 1080000
        gridCols[1].Width!.Value.Should().BeInRange(1790000, 1810000); // 5cm ≈ 1800000
        gridCols[2].Width!.Value.Should().BeInRange(1070000, 1090000);
        // Reopen handler for Dispose
        _handler = new PowerPointHandler(_path, editable: true);
    }

    // ==================== Cell padding/margin ====================

    [Fact]
    public void Set_CellMargin_SingleValue_AppliesAllSides()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["margin"] = "0.5cm" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var tc = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First();
        var tcPr = tc.TableCellProperties!;
        tcPr.LeftMargin!.Value.Should().BeInRange(175000, 185000);
        tcPr.RightMargin!.Value.Should().BeInRange(175000, 185000);
        tcPr.TopMargin!.Value.Should().BeInRange(175000, 185000);
        tcPr.BottomMargin!.Value.Should().BeInRange(175000, 185000);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_CellMargin_FourValues_AppliesIndividually()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["padding"] = "0.2cm,0.3cm,0.4cm,0.5cm" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var tc = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First();
        var tcPr = tc.TableCellProperties!;
        tcPr.LeftMargin!.Value.Should().BeInRange(70000, 75000);   // 0.2cm
        tcPr.TopMargin!.Value.Should().BeInRange(105000, 115000);  // 0.3cm
        tcPr.RightMargin!.Value.Should().BeInRange(140000, 150000); // 0.4cm
        tcPr.BottomMargin!.Value.Should().BeInRange(175000, 185000); // 0.5cm
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_CellMarginLeft_AppliesOnlyLeft()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["margin.left"] = "1cm" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var tc = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First();
        tc.TableCellProperties!.LeftMargin!.Value.Should().BeInRange(355000, 365000);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    // ==================== Text direction ====================

    [Fact]
    public void Set_TextDirection_Vertical_SetsOnCell()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["textDirection"] = "vertical" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var tc = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First();
        tc.TableCellProperties!.Vertical!.Value.Should().Be(Drawing.TextVerticalValues.Vertical270);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_TextDirection_Horizontal_SetsOnCell()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["textDirection"] = "vertical" });
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["textDirection"] = "horizontal" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var tc = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First();
        tc.TableCellProperties!.Vertical!.Value.Should().Be(Drawing.TextVerticalValues.Horizontal);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    // ==================== Word wrap ====================

    [Fact]
    public void Set_WordWrap_False_SetsNoWrap()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["wordWrap"] = "false" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var tc = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First();
        var bodyProps = tc.TextBody!.GetFirstChild<Drawing.BodyProperties>()!;
        bodyProps.Wrap!.Value.Should().Be(Drawing.TextWrappingValues.None);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    // ==================== Line spacing ====================

    [Fact]
    public void Set_LineSpacing_Multiplier_SetsSpacingPercent()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "Hello", ["lineSpacing"] = "1.5x" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var tc = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First();
        var para = tc.TextBody!.Elements<Drawing.Paragraph>().First();
        var ls = para.ParagraphProperties!.GetFirstChild<Drawing.LineSpacing>()!;
        ls.GetFirstChild<Drawing.SpacingPercent>()!.Val!.Value.Should().Be(150000);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_SpaceBefore_SetsPoints()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "Hello", ["spaceBefore"] = "12pt" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var tc = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First();
        var para = tc.TextBody!.Elements<Drawing.Paragraph>().First();
        var sb = para.ParagraphProperties!.GetFirstChild<Drawing.SpaceBefore>()!;
        sb.GetFirstChild<Drawing.SpacingPoints>()!.Val!.Value.Should().Be(1200);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    [Fact]
    public void Set_SpaceAfter_SetsPoints()
    {
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "Hello", ["spaceAfter"] = "6pt" });
        _handler.Dispose();
        using var doc = PresentationDocument.Open(_path, false);
        var tc = doc.PresentationPart!.SlideParts.First().Slide
            .Descendants<Drawing.TableCell>().First();
        var para = tc.TextBody!.Elements<Drawing.Paragraph>().First();
        var sa = para.ParagraphProperties!.GetFirstChild<Drawing.SpaceAfter>()!;
        sa.GetFirstChild<Drawing.SpacingPoints>()!.Val!.Value.Should().Be(600);
        _handler = new PowerPointHandler(_path, editable: true);
    }
}

// ==================== DOCX Table Properties Tests ====================

public class DocxTablePropertiesTests : IDisposable
{
    private readonly string _path;
    private WordHandler _handler;

    public DocxTablePropertiesTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        BlankDocCreator.Create(_path);
        _handler = new WordHandler(_path, editable: true);
        _handler.Add("/body", "table", null, new() { ["rows"] = "3", ["cols"] = "3", ["style"] = "TableGrid" });
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private WordHandler Reopen()
    {
        _handler.Dispose();
        _handler = new WordHandler(_path, editable: true);
        return _handler;
    }

    // ==================== TableLook flags ====================

    [Fact]
    public void Set_FirstRow_SetsTableLookAttribute()
    {
        _handler.Set("/body/tbl[1]", new() { ["firstRow"] = "true" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var tblLook = doc.MainDocumentPart!.Document!.Body!
            .Descendants<DocumentFormat.OpenXml.Wordprocessing.TableLook>().First();
        tblLook.FirstRow!.Value.Should().BeTrue();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_LastRow_SetsTableLookAttribute()
    {
        _handler.Set("/body/tbl[1]", new() { ["lastRow"] = "true" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var tblLook = doc.MainDocumentPart!.Document!.Body!
            .Descendants<DocumentFormat.OpenXml.Wordprocessing.TableLook>().First();
        tblLook.LastRow!.Value.Should().BeTrue();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_FirstCol_SetsTableLookAttribute()
    {
        _handler.Set("/body/tbl[1]", new() { ["firstCol"] = "true" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var tblLook = doc.MainDocumentPart!.Document!.Body!
            .Descendants<DocumentFormat.OpenXml.Wordprocessing.TableLook>().First();
        tblLook.FirstColumn!.Value.Should().BeTrue();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_BandedRows_True_ClearsNoHBand()
    {
        _handler.Set("/body/tbl[1]", new() { ["bandedRows"] = "true" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var tblLook = doc.MainDocumentPart!.Document!.Body!
            .Descendants<DocumentFormat.OpenXml.Wordprocessing.TableLook>().First();
        tblLook.NoHorizontalBand!.Value.Should().BeFalse();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_BandedCols_True_ClearsNoVBand()
    {
        _handler.Set("/body/tbl[1]", new() { ["bandedCols"] = "true" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var tblLook = doc.MainDocumentPart!.Document!.Body!
            .Descendants<DocumentFormat.OpenXml.Wordprocessing.TableLook>().First();
        tblLook.NoVerticalBand!.Value.Should().BeFalse();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_TableLookFlags_Persist_SurvivesReopen()
    {
        _handler.Set("/body/tbl[1]", new()
        {
            ["firstRow"] = "true", ["lastRow"] = "true",
            ["firstCol"] = "true", ["lastCol"] = "true",
            ["bandedRows"] = "true", ["bandedCols"] = "true"
        });
        Reopen();
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var tblLook = doc.MainDocumentPart!.Document!.Body!
            .Descendants<DocumentFormat.OpenXml.Wordprocessing.TableLook>().First();
        tblLook.FirstRow!.Value.Should().BeTrue();
        tblLook.LastRow!.Value.Should().BeTrue();
        tblLook.FirstColumn!.Value.Should().BeTrue();
        tblLook.LastColumn!.Value.Should().BeTrue();
        tblLook.NoHorizontalBand!.Value.Should().BeFalse();
        tblLook.NoVerticalBand!.Value.Should().BeFalse();
        _handler = new WordHandler(_path, editable: true);
    }

    // ==================== Caption & Description ====================

    [Fact]
    public void Set_Caption_SetsTableCaption()
    {
        _handler.Set("/body/tbl[1]", new() { ["caption"] = "Sales Data" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var caption = doc.MainDocumentPart!.Document!.Body!
            .Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCaption>().FirstOrDefault();
        caption.Should().NotBeNull();
        caption!.Val!.Value.Should().Be("Sales Data");
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_Description_SetsTableDescription()
    {
        _handler.Set("/body/tbl[1]", new() { ["description"] = "Quarterly revenue breakdown" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var desc = doc.MainDocumentPart!.Document!.Body!
            .Descendants<DocumentFormat.OpenXml.Wordprocessing.TableDescription>().FirstOrDefault();
        desc.Should().NotBeNull();
        desc!.Val!.Value.Should().Be("Quarterly revenue breakdown");
        _handler = new WordHandler(_path, editable: true);
    }

    // ==================== FitText ====================

    [Fact]
    public void Set_FitText_True_AddsFitTextElement()
    {
        _handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Hello World" });
        _handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["fitText"] = "true" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var fitText = doc.MainDocumentPart!.Document!.Body!
            .Descendants<DocumentFormat.OpenXml.Wordprocessing.FitText>().FirstOrDefault();
        fitText.Should().NotBeNull();
        _handler = new WordHandler(_path, editable: true);
    }

    [Fact]
    public void Set_FitText_False_RemovesFitTextElement()
    {
        _handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Hello World" });
        _handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["fitText"] = "true" });
        _handler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["fitText"] = "false" });
        _handler.Dispose();
        using var doc = WordprocessingDocument.Open(_path, false);
        var fitText = doc.MainDocumentPart!.Document!.Body!
            .Descendants<DocumentFormat.OpenXml.Wordprocessing.FitText>().FirstOrDefault();
        fitText.Should().BeNull();
        _handler = new WordHandler(_path, editable: true);
    }
}

// ==================== XLSX Table Properties Tests ====================

public class ExcelTablePropertiesTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelTablePropertiesTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_path);
        _handler = new ExcelHandler(_path, editable: true);
        _handler.Set("/Sheet1/A1", new() { ["value"] = "Name" });
        _handler.Set("/Sheet1/B1", new() { ["value"] = "Revenue" });
        _handler.Set("/Sheet1/C1", new() { ["value"] = "Quantity" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "Widget" });
        _handler.Set("/Sheet1/B2", new() { ["value"] = "100" });
        _handler.Set("/Sheet1/C2", new() { ["value"] = "10" });
        _handler.Set("/Sheet1/A3", new() { ["value"] = "Gadget" });
        _handler.Set("/Sheet1/B3", new() { ["value"] = "200" });
        _handler.Set("/Sheet1/C3", new() { ["value"] = "20" });
        _handler.Add("/Sheet1", "table", null, new() { ["ref"] = "A1:C3", ["name"] = "SalesData" });
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private ExcelHandler Reopen()
    {
        _handler.Dispose();
        _handler = new ExcelHandler(_path, editable: true);
        return _handler;
    }

    // ==================== Style options ====================

    [Fact]
    public void Set_ShowRowStripes_False_UpdatesTableStyleInfo()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["showRowStripes"] = "false" });
        var node = _handler.Get("/Sheet1/table[1]");
        node.Format["showRowStripes"].Should().Be(false);
    }

    [Fact]
    public void Set_ShowColumnStripes_True_UpdatesTableStyleInfo()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["showColumnStripes"] = "true" });
        var node = _handler.Get("/Sheet1/table[1]");
        node.Format["showColumnStripes"].Should().Be(true);
    }

    [Fact]
    public void Set_ShowFirstColumn_True_UpdatesTableStyleInfo()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["showFirstColumn"] = "true" });
        var node = _handler.Get("/Sheet1/table[1]");
        node.Format["showFirstColumn"].Should().Be(true);
    }

    [Fact]
    public void Set_ShowLastColumn_True_UpdatesTableStyleInfo()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["showLastColumn"] = "true" });
        var node = _handler.Get("/Sheet1/table[1]");
        node.Format["showLastColumn"].Should().Be(true);
    }

    [Fact]
    public void Set_StyleOptions_Persist_SurvivesReopen()
    {
        _handler.Set("/Sheet1/table[1]", new()
        {
            ["showRowStripes"] = "false",
            ["showColumnStripes"] = "true",
            ["showFirstColumn"] = "true",
            ["showLastColumn"] = "true"
        });
        Reopen();
        var node = _handler.Get("/Sheet1/table[1]");
        node.Format["showRowStripes"].Should().Be(false);
        node.Format["showColumnStripes"].Should().Be(true);
        node.Format["showFirstColumn"].Should().Be(true);
        node.Format["showLastColumn"].Should().Be(true);
    }

    // ==================== Column-level properties ====================

    [Fact]
    public void Set_ColumnName_UpdatesColumnDefinition()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["col[1].name"] = "FullName" });
        _handler.Dispose();
        using var doc = SpreadsheetDocument.Open(_path, false);
        var table = doc.WorkbookPart!.WorksheetParts.First()
            .TableDefinitionParts.First().Table!;
        var cols = table.GetFirstChild<TableColumns>()!.Elements<TableColumn>().ToList();
        cols[0].Name!.Value.Should().Be("FullName");
        _handler = new ExcelHandler(_path, editable: true);
    }

    [Fact]
    public void Set_ColumnTotalFunction_Sum_SetsOnColumn()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["col[2].totalFunction"] = "sum" });
        _handler.Dispose();
        using var doc = SpreadsheetDocument.Open(_path, false);
        var table = doc.WorkbookPart!.WorksheetParts.First()
            .TableDefinitionParts.First().Table!;
        var cols = table.GetFirstChild<TableColumns>()!.Elements<TableColumn>().ToList();
        cols[1].TotalsRowFunction!.Value.Should().Be(TotalsRowFunctionValues.Sum);
        _handler = new ExcelHandler(_path, editable: true);
    }

    [Fact]
    public void Set_ColumnTotalFunction_Average_SetsOnColumn()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["col[3].totalFunction"] = "average" });
        _handler.Dispose();
        using var doc = SpreadsheetDocument.Open(_path, false);
        var table = doc.WorkbookPart!.WorksheetParts.First()
            .TableDefinitionParts.First().Table!;
        var cols = table.GetFirstChild<TableColumns>()!.Elements<TableColumn>().ToList();
        cols[2].TotalsRowFunction!.Value.Should().Be(TotalsRowFunctionValues.Average);
        _handler = new ExcelHandler(_path, editable: true);
    }

    [Fact]
    public void Set_ColumnTotalLabel_SetsOnColumn()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["col[1].totalLabel"] = "Total" });
        _handler.Dispose();
        using var doc = SpreadsheetDocument.Open(_path, false);
        var table = doc.WorkbookPart!.WorksheetParts.First()
            .TableDefinitionParts.First().Table!;
        var cols = table.GetFirstChild<TableColumns>()!.Elements<TableColumn>().ToList();
        cols[0].TotalsRowLabel!.Value.Should().Be("Total");
        _handler = new ExcelHandler(_path, editable: true);
    }

    [Fact]
    public void Set_ColumnFormula_SetsCalculatedColumnFormula()
    {
        _handler.Set("/Sheet1/table[1]", new() { ["col[3].formula"] = "[Revenue]*[Quantity]" });
        _handler.Dispose();
        using var doc = SpreadsheetDocument.Open(_path, false);
        var table = doc.WorkbookPart!.WorksheetParts.First()
            .TableDefinitionParts.First().Table!;
        var cols = table.GetFirstChild<TableColumns>()!.Elements<TableColumn>().ToList();
        cols[2].CalculatedColumnFormula.Should().NotBeNull();
        cols[2].CalculatedColumnFormula!.Text.Should().Be("[Revenue]*[Quantity]");
        _handler = new ExcelHandler(_path, editable: true);
    }
}
