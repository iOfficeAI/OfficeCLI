// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Core;

/// <summary>
/// Resolves the effective formula text of a cell, expanding shared-formula
/// children.
///
/// <para>Excel stores a formula filled across a range once: the master cell
/// carries <c>&lt;f t="shared" ref="B1:B3" si="0"&gt;IF(A1&gt;1,A1*2,"")&lt;/f&gt;</c>
/// and every other cell in <c>ref</c> carries only <c>&lt;f t="shared" si="0"/&gt;</c>.
/// A child's formula is the master's text with relative references displaced
/// by the child's offset from the master — the same rule as copy/paste, which
/// is why the expansion reuses <see cref="FormulaRefShifter.ApplyCopyDelta"/>.
/// Reading <c>CellFormula.Text</c> on a child yields an empty string, which the
/// evaluator and the readbacks used to take at face value: the child was
/// reported as a formula-less cell, evaluated to blank, and any cell depending
/// on it was recomputed from that blank.</para>
///
/// <para>The XML is never rewritten — children keep their shared form on disk,
/// exactly as Excel leaves them. Only the text handed to the evaluator and the
/// readbacks is expanded.</para>
///
/// <para>Known limitation: cross-sheet relative references inside a shared
/// master (<c>Sheet2!A1</c>) are not displaced for children — ApplyCopyDelta
/// shifts same-sheet references only. Same-sheet references, which is what
/// shared formulas overwhelmingly contain, are handled fully.</para>
/// </summary>
internal static class SharedFormulaResolver
{
    private static readonly Regex CellRefPattern = new(@"^\$?([A-Za-z]{1,3})\$?(\d+)$", RegexOptions.Compiled);

    // Masters per sheetData, keyed by si. Built lazily on the first child lookup
    // and rebuilt whenever a cached master turns out to be detached (its row was
    // removed) so a stale entry can never be served.
    private static readonly ConditionalWeakTable<SheetData, Dictionary<uint, Cell>> MasterIndex = new();

    /// <summary>
    /// The formula text the cell effectively holds: its own <c>&lt;f&gt;</c> text,
    /// or — for a shared-formula child with no text — the master's text displaced
    /// to this cell. Returns null when the cell has no <c>&lt;f&gt;</c> at all, and
    /// the raw (empty) text when the child's master cannot be found.
    /// </summary>
    public static string? ResolveText(Cell cell, SheetData? sheetData = null)
    {
        var f = cell.CellFormula;
        if (f == null) return null;
        var text = f.Text;
        if (!string.IsNullOrEmpty(text)) return text;
        if (f.FormulaType?.Value != CellFormulaValues.Shared || f.SharedIndex?.Value is not { } si)
            return text;

        sheetData ??= cell.Parent?.Parent as SheetData;
        if (sheetData == null) return text;

        var master = FindMaster(sheetData, si);
        if (master == null || ReferenceEquals(master, cell)) return text;
        return Expand(master, cell) ?? text;
    }

    /// <summary>True when the cell is a shared-formula child (shared type, no text of its own).</summary>
    public static bool IsSharedChild(Cell cell)
        => cell.CellFormula is { } f
            && string.IsNullOrEmpty(f.Text)
            && f.FormulaType?.Value == CellFormulaValues.Shared
            && f.SharedIndex != null;

    private static Cell? FindMaster(SheetData sheetData, uint si)
    {
        var index = MasterIndex.GetValue(sheetData, BuildMasterIndex);
        if (index.TryGetValue(si, out var master) && master.Parent != null)
            return master;
        // Missing or detached: the sheet changed since the index was built.
        index = BuildMasterIndex(sheetData);
        MasterIndex.AddOrUpdate(sheetData, index);
        return index.TryGetValue(si, out master) ? master : null;
    }

    private static Dictionary<uint, Cell> BuildMasterIndex(SheetData sheetData)
    {
        var index = new Dictionary<uint, Cell>();
        foreach (var row in sheetData.Elements<Row>())
            foreach (var c in row.Elements<Cell>())
            {
                var cf = c.CellFormula;
                if (cf?.FormulaType?.Value == CellFormulaValues.Shared
                    && cf.SharedIndex?.Value is { } si
                    && cf.Reference != null
                    && !string.IsNullOrEmpty(cf.Text))
                    index.TryAdd(si, c);
            }
        return index;
    }

    private static string? Expand(Cell master, Cell child)
    {
        var masterText = master.CellFormula?.Text;
        if (string.IsNullOrEmpty(masterText)) return null;
        var mm = CellRefPattern.Match(master.CellReference?.Value ?? "");
        var cm = CellRefPattern.Match(child.CellReference?.Value ?? "");
        if (!mm.Success || !cm.Success) return null;
        var deltaCol = ColumnLettersToIndex(cm.Groups[1].Value) - ColumnLettersToIndex(mm.Groups[1].Value);
        var deltaRow = int.Parse(cm.Groups[2].Value) - int.Parse(mm.Groups[2].Value);
        // Both sheet names empty: unqualified refs resolve to "" and match "",
        // so they shift; sheet-qualified refs never match and stay as written.
        return FormulaRefShifter.ApplyCopyDelta(masterText, "", "", deltaCol, deltaRow);
    }

    private static int ColumnLettersToIndex(string letters)
    {
        var n = 0;
        foreach (var ch in letters.ToUpperInvariant()) n = n * 26 + (ch - 'A' + 1);
        return n;
    }
}
