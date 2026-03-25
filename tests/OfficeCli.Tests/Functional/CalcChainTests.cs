// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests that the calculation chain part is auto-deleted after cell/formula mutations.
/// This prevents stale calc chain references that can cause Excel repair prompts.
/// </summary>
public class CalcChainTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public CalcChainTests()
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

    private ExcelHandler Reopen()
    {
        _handler.Dispose();
        _handler = new ExcelHandler(_path, editable: true);
        return _handler;
    }

    /// <summary>
    /// Inject a fake CalculationChainPart into the workbook to simulate an existing calc chain.
    /// </summary>
    private void InjectCalcChain()
    {
        _handler.Dispose();
        using (var doc = SpreadsheetDocument.Open(_path, true))
        {
            var wbPart = doc.WorkbookPart!;
            if (wbPart.CalculationChainPart == null)
            {
                var ccPart = wbPart.AddNewPart<CalculationChainPart>();
                ccPart.CalculationChain = new CalculationChain(
                    new CalculationCell { CellReference = "A1", SheetId = 1 }
                );
                ccPart.CalculationChain.Save();
            }
        }
        _handler = new ExcelHandler(_path, editable: true);
    }

    private bool HasCalcChain()
    {
        _handler.Dispose();
        bool result;
        using (var doc = SpreadsheetDocument.Open(_path, false))
        {
            result = doc.WorkbookPart?.CalculationChainPart != null;
        }
        _handler = new ExcelHandler(_path, editable: true);
        return result;
    }

    [Fact]
    public void Set_CellValue_DeletesCalcChain()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });
        InjectCalcChain();
        HasCalcChain().Should().BeTrue("calc chain was injected");

        _handler.Set("/Sheet1/A1", new() { ["value"] = "20" });

        HasCalcChain().Should().BeFalse("calc chain should be deleted after cell value mutation");
    }

    [Fact]
    public void Set_CellFormula_DeletesCalcChain()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });
        InjectCalcChain();

        _handler.Set("/Sheet1/A1", new() { ["formula"] = "=1+1" });

        HasCalcChain().Should().BeFalse("calc chain should be deleted after formula mutation");
    }

    [Fact]
    public void Add_CellWithFormula_DeletesCalcChain()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "5" });
        InjectCalcChain();

        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["formula"] = "=A1*2" });

        HasCalcChain().Should().BeFalse("calc chain should be deleted after adding cell with formula");
    }

    [Fact]
    public void Remove_Cell_DeletesCalcChain()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });
        InjectCalcChain();

        _handler.Remove("/Sheet1/A1");

        HasCalcChain().Should().BeFalse("calc chain should be deleted after cell removal");
    }

    [Fact]
    public void Remove_Row_DeletesCalcChain()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "20" });
        InjectCalcChain();

        _handler.Remove("/Sheet1/row[1]");

        HasCalcChain().Should().BeFalse("calc chain should be deleted after row removal");
    }

    [Fact]
    public void Remove_Column_DeletesCalcChain()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "20" });
        InjectCalcChain();

        _handler.Remove("/Sheet1/col[A]");

        HasCalcChain().Should().BeFalse("calc chain should be deleted after column removal");
    }

    [Fact]
    public void NoCalcChain_SetDoesNotThrow()
    {
        // When there is no calc chain, Set should still work fine
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });

        var act = () => _handler.Set("/Sheet1/A1", new() { ["value"] = "20" });
        act.Should().NotThrow();
    }
}
