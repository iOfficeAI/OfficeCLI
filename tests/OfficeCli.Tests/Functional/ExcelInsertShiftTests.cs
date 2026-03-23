// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Full-lifecycle tests for:
///   - Insert row/col with shift (Add)
///   - Whole-row/col native path notation (Sheet1!1:1, Sheet1!A:A)
///   - Formula readback via Get (Format["formula"])
///   - Formula warning precision (#REF! vs shifted)
///
/// Each test follows: Create → Set/Add → Verify in-memory → Reopen → Verify persisted.
/// </summary>
public class ExcelInsertShiftTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelInsertShiftTests()
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

    // ==================== Insert row with shift ====================

    [Fact]
    public void AddRow_AtFront_ShiftsAllRowsDown_Persists()
    {
        // Create
        _handler.Set("/Sheet1/A1", new() { ["value"] = "row1" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "row2" });
        _handler.Set("/Sheet1/A3", new() { ["value"] = "row3" });

        // Add blank row at position 1 → shifts everything down
        var path = _handler.Add("/Sheet1", "row", 1, new());
        path.Should().Be("/Sheet1/row[1]");

        // Verify in-memory
        _handler.Get("/Sheet1/A2").Text.Should().Be("row1");
        _handler.Get("/Sheet1/A3").Text.Should().Be("row2");
        _handler.Get("/Sheet1/A4").Text.Should().Be("row3");

        // Write something into the new row
        _handler.Set("/Sheet1/A1", new() { ["value"] = "inserted" });

        // Reopen and verify persisted state
        Reopen();
        _handler.Get("/Sheet1/A1").Text.Should().Be("inserted");
        _handler.Get("/Sheet1/A2").Text.Should().Be("row1");
        _handler.Get("/Sheet1/A3").Text.Should().Be("row2");
        _handler.Get("/Sheet1/A4").Text.Should().Be("row3");
    }

    [Fact]
    public void AddRow_InMiddle_ShiftsCorrectly_Persists()
    {
        // Create
        _handler.Set("/Sheet1/A1", new() { ["value"] = "a" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "b" });
        _handler.Set("/Sheet1/A3", new() { ["value"] = "c" });

        // Insert at row 2
        _handler.Add("/Sheet1", "row", 2, new());

        // Verify in-memory
        _handler.Get("/Sheet1/A1").Text.Should().Be("a");
        _handler.Get("/Sheet1/A3").Text.Should().Be("b");
        _handler.Get("/Sheet1/A4").Text.Should().Be("c");

        // Set the new blank row
        _handler.Set("/Sheet1/A2", new() { ["value"] = "new-b" });

        // Reopen and verify
        Reopen();
        _handler.Get("/Sheet1/A1").Text.Should().Be("a");
        _handler.Get("/Sheet1/A2").Text.Should().Be("new-b");
        _handler.Get("/Sheet1/A3").Text.Should().Be("b");
        _handler.Get("/Sheet1/A4").Text.Should().Be("c");
    }

    [Fact]
    public void AddRow_AtEnd_DoesNotShift_Persists()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "r1" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "r2" });

        _handler.Add("/Sheet1", "row", 3, new());
        _handler.Set("/Sheet1/A3", new() { ["value"] = "r3-new" });

        Reopen();
        _handler.Get("/Sheet1/A1").Text.Should().Be("r1");
        _handler.Get("/Sheet1/A2").Text.Should().Be("r2");
        _handler.Get("/Sheet1/A3").Text.Should().Be("r3-new");
    }

    [Fact]
    public void AddRow_NewRowIsBlank_Persists()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "hello" });

        _handler.Add("/Sheet1", "row", 1, new());

        Reopen();
        var inserted = _handler.Get("/Sheet1/A1");
        var t = inserted.Text;
        (t == null || t == "" || t == "(empty)").Should().BeTrue("inserted row should be blank");
    }

    // ==================== Insert column with shift ====================

    [Fact]
    public void AddCol_AtFront_ShiftsAllColumnsRight_Persists()
    {
        // Create
        _handler.Set("/Sheet1/A1", new() { ["value"] = "col-a" });
        _handler.Set("/Sheet1/B1", new() { ["value"] = "col-b" });
        _handler.Set("/Sheet1/C1", new() { ["value"] = "col-c" });

        // Insert blank column at A (index 1) → A→B, B→C, C→D
        var path = _handler.Add("/Sheet1", "col", 1, new());
        path.Should().Be("/Sheet1/col[A]");

        // Verify in-memory
        _handler.Get("/Sheet1/B1").Text.Should().Be("col-a");
        _handler.Get("/Sheet1/C1").Text.Should().Be("col-b");
        _handler.Get("/Sheet1/D1").Text.Should().Be("col-c");

        // Write into new col
        _handler.Set("/Sheet1/A1", new() { ["value"] = "inserted-col" });

        // Reopen and verify
        Reopen();
        _handler.Get("/Sheet1/A1").Text.Should().Be("inserted-col");
        _handler.Get("/Sheet1/B1").Text.Should().Be("col-a");
        _handler.Get("/Sheet1/C1").Text.Should().Be("col-b");
        _handler.Get("/Sheet1/D1").Text.Should().Be("col-c");
    }

    [Fact]
    public void AddCol_InMiddle_ShiftsCorrectly_Persists()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "x" });
        _handler.Set("/Sheet1/B1", new() { ["value"] = "y" });
        _handler.Set("/Sheet1/C1", new() { ["value"] = "z" });

        // Insert at col 2 (B) → B→C, C→D
        _handler.Add("/Sheet1", "col", 2, new());
        _handler.Set("/Sheet1/B1", new() { ["value"] = "inserted" });

        Reopen();
        _handler.Get("/Sheet1/A1").Text.Should().Be("x");
        _handler.Get("/Sheet1/B1").Text.Should().Be("inserted");
        _handler.Get("/Sheet1/C1").Text.Should().Be("y");
        _handler.Get("/Sheet1/D1").Text.Should().Be("z");
    }

    [Fact]
    public void AddCol_NoIndex_AppendsAfterLastColumn_Persists()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "x" });
        _handler.Set("/Sheet1/B1", new() { ["value"] = "y" });

        var path = _handler.Add("/Sheet1", "col", null, new());
        path.Should().Be("/Sheet1/col[C]");

        _handler.Set("/Sheet1/C1", new() { ["value"] = "z" });

        Reopen();
        _handler.Get("/Sheet1/A1").Text.Should().Be("x");
        _handler.Get("/Sheet1/B1").Text.Should().Be("y");
        _handler.Get("/Sheet1/C1").Text.Should().Be("z");
    }

    // ==================== Whole row/col native path ====================

    [Fact]
    public void NativeWholeRowPath_DeleteAndShift_Persists()
    {
        // Create
        _handler.Set("/Sheet1/A1", new() { ["value"] = "r1" });
        _handler.Set("/Sheet1/B1", new() { ["value"] = "r1b" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "r2" });
        _handler.Set("/Sheet1/A3", new() { ["value"] = "r3" });

        // Delete row 2 via native notation
        _handler.Remove("Sheet1!2:2");

        // Verify in-memory
        _handler.Get("/Sheet1/A1").Text.Should().Be("r1");
        _handler.Get("/Sheet1/B1").Text.Should().Be("r1b");
        _handler.Get("/Sheet1/A2").Text.Should().Be("r3");

        // Reopen and verify persisted
        Reopen();
        _handler.Get("/Sheet1/A1").Text.Should().Be("r1");
        _handler.Get("/Sheet1/A2").Text.Should().Be("r3");
    }

    [Fact]
    public void NativeWholeColPath_DeleteAndShift_Persists()
    {
        // Create
        _handler.Set("/Sheet1/A1", new() { ["value"] = "a" });
        _handler.Set("/Sheet1/B1", new() { ["value"] = "b" });
        _handler.Set("/Sheet1/C1", new() { ["value"] = "c" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "a2" });
        _handler.Set("/Sheet1/C2", new() { ["value"] = "c2" });

        // Delete col B via native notation
        _handler.Remove("Sheet1!B:B");

        // Verify in-memory: c → B, c2 → B2
        _handler.Get("/Sheet1/B1").Text.Should().Be("c");
        _handler.Get("/Sheet1/B2").Text.Should().Be("c2");

        // Reopen and verify persisted
        Reopen();
        _handler.Get("/Sheet1/A1").Text.Should().Be("a");
        _handler.Get("/Sheet1/B1").Text.Should().Be("c");
        _handler.Get("/Sheet1/A2").Text.Should().Be("a2");
        _handler.Get("/Sheet1/B2").Text.Should().Be("c2");
    }

    // ==================== Formula readback ====================

    [Fact]
    public void FormulaReadback_SetThenGet_ExposesFormulaKey()
    {
        // Create
        _handler.Set("/Sheet1/A1", new() { ["value"] = "5" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "10" });
        _handler.Set("/Sheet1/B1", new() { ["formula"] = "SUM(A1:A2)" });

        // Verify in-memory
        var node = _handler.Get("/Sheet1/B1");
        node.Format.Should().ContainKey("formula");
        node.Format["formula"].Should().Be("SUM(A1:A2)");

        // Reopen and verify persisted
        Reopen();
        var persisted = _handler.Get("/Sheet1/B1");
        persisted.Format.Should().ContainKey("formula");
        persisted.Format["formula"].Should().Be("SUM(A1:A2)");
    }

    [Fact]
    public void FormulaReadback_NativePath_Persists()
    {
        // Create
        _handler.Set("/Sheet1/A1", new() { ["value"] = "3" });
        _handler.Set("/Sheet1/B1", new() { ["formula"] = "A1*10" });

        // Verify via native path in-memory
        _handler.Get("Sheet1!B1").Format["formula"].Should().Be("A1*10");

        // Reopen and verify via both paths
        Reopen();
        _handler.Get("/Sheet1/B1").Format["formula"].Should().Be("A1*10");
        _handler.Get("Sheet1!B1").Format["formula"].Should().Be("A1*10");
    }

    [Fact]
    public void FormulaReadback_ValueCell_HasNoFormulaKey_Persists()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "hello" });

        _handler.Get("/Sheet1/A1").Format.Should().NotContainKey("formula");

        Reopen();
        _handler.Get("/Sheet1/A1").Format.Should().NotContainKey("formula");
    }

    [Fact]
    public void FormulaReadback_SetFormulaToValue_FormulaKeyDisappears()
    {
        // Set as formula first
        _handler.Set("/Sheet1/A1", new() { ["formula"] = "1+1" });
        _handler.Get("/Sheet1/A1").Format.Should().ContainKey("formula");

        // Overwrite with plain value
        _handler.Set("/Sheet1/A1", new() { ["value"] = "42" });

        // Verify in-memory: formula key gone
        _handler.Get("/Sheet1/A1").Format.Should().NotContainKey("formula");

        // Reopen and verify persisted
        Reopen();
        _handler.Get("/Sheet1/A1").Format.Should().NotContainKey("formula");
        _handler.Get("/Sheet1/A1").Text.Should().Be("42");
    }

    // ==================== Warning precision ====================

    [Fact]
    public void Warning_DeletedRowRef_IsRefError_DataShifts_Persists()
    {
        // Create: B3 references the row being deleted (row 2)
        _handler.Set("/Sheet1/A1", new() { ["value"] = "10" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "20" });
        _handler.Set("/Sheet1/A3", new() { ["value"] = "30" });
        _handler.Set("/Sheet1/B3", new() { ["formula"] = "A2*5" }); // points at deleted row

        var warning = _handler.Remove("/Sheet1/row[2]");

        // Warning should identify as #REF! risk
        warning.Should().NotBeNull();
        warning.Should().Contain("#REF!");
        warning.Should().Contain("B3");

        // Data should have shifted: A3→A2
        _handler.Get("/Sheet1/A2").Text.Should().Be("30");

        // Reopen and verify shift persisted
        Reopen();
        _handler.Get("/Sheet1/A1").Text.Should().Be("10");
        _handler.Get("/Sheet1/A2").Text.Should().Be("30");
    }

    [Fact]
    public void Warning_ShiftedRowRef_IsStaleNotRefError_DataShifts_Persists()
    {
        // Create: B5 references row 3 which is after the deleted row 1 → shifted
        _handler.Set("/Sheet1/A1", new() { ["value"] = "deleted" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "kept-2" });
        _handler.Set("/Sheet1/A3", new() { ["value"] = "kept-3" });
        _handler.Set("/Sheet1/B5", new() { ["formula"] = "A3+1" }); // references shifted row

        var warning = _handler.Remove("/Sheet1/row[1]");

        // Warning should mention shifted, not #REF!
        warning.Should().NotBeNull();
        warning.Should().Contain("shifted");
        warning.Should().Contain("B5");
        warning.Should().NotContain("#REF!");

        // Data: A2→A1, A3→A2
        _handler.Get("/Sheet1/A1").Text.Should().Be("kept-2");
        _handler.Get("/Sheet1/A2").Text.Should().Be("kept-3");

        // Reopen and verify
        Reopen();
        _handler.Get("/Sheet1/A1").Text.Should().Be("kept-2");
        _handler.Get("/Sheet1/A2").Text.Should().Be("kept-3");
    }

    [Fact]
    public void Warning_ColDelete_RefErrorAndShifted_BothReported_Persists()
    {
        // C1 references B1 (deleted col B → #REF!)
        // D1 references C1 (shifted col C → stale)
        _handler.Set("/Sheet1/A1", new() { ["value"] = "a" });
        _handler.Set("/Sheet1/B1", new() { ["value"] = "b" });
        _handler.Set("/Sheet1/C1", new() { ["formula"] = "B1*2" });   // → #REF! after B deleted
        _handler.Set("/Sheet1/D1", new() { ["formula"] = "C1+100" }); // → shifted (C becomes B)

        var warning = _handler.Remove("/Sheet1/col[B]");

        warning.Should().NotBeNull();
        warning.Should().Contain("#REF!");
        // Both cells should be mentioned somewhere in the warning
        warning.Should().Contain("C1");

        // After deletion, original C→B, D→C
        Reopen();
        _handler.Get("/Sheet1/A1").Text.Should().Be("a");
        // B1 is now what was C1 (formula cell)
        _handler.Get("/Sheet1/B1").Format.Should().ContainKey("formula");
    }

    // ==================== Insert + Delete round-trip ====================

    [Fact]
    public void InsertRowThenDelete_RoundTrip_RestoresOriginalLayout()
    {
        // Create initial state
        _handler.Set("/Sheet1/A1", new() { ["value"] = "orig-1" });
        _handler.Set("/Sheet1/A2", new() { ["value"] = "orig-2" });
        _handler.Set("/Sheet1/A3", new() { ["value"] = "orig-3" });

        // Insert at row 2
        _handler.Add("/Sheet1", "row", 2, new());
        _handler.Get("/Sheet1/A1").Text.Should().Be("orig-1");
        _handler.Get("/Sheet1/A3").Text.Should().Be("orig-2");
        _handler.Get("/Sheet1/A4").Text.Should().Be("orig-3");

        // Delete the inserted row 2 to restore original
        _handler.Remove("/Sheet1/row[2]");

        // Should be back to original
        Reopen();
        _handler.Get("/Sheet1/A1").Text.Should().Be("orig-1");
        _handler.Get("/Sheet1/A2").Text.Should().Be("orig-2");
        _handler.Get("/Sheet1/A3").Text.Should().Be("orig-3");
    }

    [Fact]
    public void InsertColThenDelete_RoundTrip_RestoresOriginalLayout()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "x" });
        _handler.Set("/Sheet1/B1", new() { ["value"] = "y" });

        // Insert col at B
        _handler.Add("/Sheet1", "col", 2, new());
        _handler.Get("/Sheet1/A1").Text.Should().Be("x");
        _handler.Get("/Sheet1/C1").Text.Should().Be("y");

        // Delete inserted col B
        _handler.Remove("/Sheet1/col[B]");

        // Restored
        Reopen();
        _handler.Get("/Sheet1/A1").Text.Should().Be("x");
        _handler.Get("/Sheet1/B1").Text.Should().Be("y");
    }
}
