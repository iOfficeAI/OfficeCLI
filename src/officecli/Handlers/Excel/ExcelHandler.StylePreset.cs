// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ===== Range style presets (set '/Sheet1/range[B2:T6]' --prop stylepreset=…, 03 §6) =====
    //
    // One-command styling for free-form ranges. Each preset is a curated bundle
    // of EXISTING cell style props (fill, font.color, bold, size, italic,
    // border.all, alignment.horizontal) — expanded through the very same
    // ApplyCellProperties path a hand-written `set` uses, so nothing a preset
    // requests can silently render as something else ("expand, don't invent" —
    // mirrors WordHandler.StylePreset from PR #376). Accent-keyed colors
    // (ACCENT1 / ACCENT1_TINT) resolve from the workbook theme part at apply
    // time with a fixed Office-default fallback, so a theme-less file still
    // gets sensible colors. Division of labor with ListObjects: `add --type
    // table --prop style=mediumN` styles a real table (banding that scales
    // with the table); stylepreset styles a free range — that's why table_banded
    // refuses ranges taller than 1000 rows and points at add table instead.

    /// <summary>A resolved preset: the props applied to every cell, plus an
    /// optional per-row override bundle (table_banded's zebra stripe).</summary>
    private sealed record StylePreset(Dictionary<string, string> Props, Dictionary<string, string>? OddRowProps = null);

    /// <summary>
    /// Registry of built-in xlsx range presets. All values are existing cell
    /// style-prop values, with two accent tokens: ACCENT1 (theme accent1 or
    /// the Office default 4472C4) and ACCENT1_TINT (accent1 blended 85% toward
    /// white — a light band/background tone). Colors mix deterministically so
    /// the equivalence test can recompute the exact hex from the readback
    /// theme accent.
    /// </summary>
    private static readonly Dictionary<string, StylePreset> StylePresets =
        new(StringComparer.OrdinalIgnoreCase)
        {
            // Header band: accent fill, white bold, centered.
            ["table_header"] = new StylePreset(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["fill"] = "ACCENT1",
                ["font.color"] = "FFFFFF",
                ["bold"] = "true",
                ["alignment.horizontal"] = "center",
            }),

            // Zebra rows: odd range-rows carry the light accent band, even rows
            // untouched (explicit cellProps from the same call still apply).
            ["table_banded"] = new StylePreset(
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
                OddRowProps: new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["fill"] = "ACCENT1_TINT",
                }),

            // KPI card: light accent panel, accent bold number, centered.
            ["kpi_card"] = new StylePreset(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["fill"] = "ACCENT1_TINT",
                ["font.color"] = "ACCENT1",
                ["bold"] = "true",
                ["size"] = "20",
                ["alignment.horizontal"] = "center",
            }),

            // Excel's canonical "Good" / "Bad" semantic pair (fixed, not themed —
            // good/bad must stay green/red whatever the workbook theme is).
            ["metric_positive"] = new StylePreset(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["font.color"] = "006100",
                ["fill"] = "C6EFCE",
            }),
            ["metric_negative"] = new StylePreset(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["font.color"] = "9C0006",
                ["fill"] = "FFC7CE",
            }),

            // Muted note text.
            ["note_gray"] = new StylePreset(new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["font.color"] = "808080",
                ["italic"] = "true",
                ["fill"] = "F2F2F2",
            }),
        };

    /// <summary>Office default theme accent1, used when the workbook has no
    /// theme part or an unreadable accent slot.</summary>
    private const string FallbackAccent1 = "4472C4";

    /// <summary>
    /// Apply a style preset to every cell of the range (atomic: restore on
    /// failure). Called from SetRange after explicit cellProps so the preset's
    /// bundle wins conflicts deterministically. Returns the unsupported list
    /// collected from the first cell only, mirroring SetRange's convention.
    /// </summary>
    private List<string> ApplyStylePresetRange(WorksheetPart worksheet, string rangeRef, string presetName,
        Dictionary<string, string> explicitProps)
    {
        if (!StylePresets.TryGetValue(presetName, out var preset))
            throw new Core.CliException($"Unknown stylepreset: '{presetName}'.")
            {
                Code = "invalid_value",
                ValidValues = StylePresets.Keys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase).ToArray(),
                Suggestion = "ListObject tables take built-in styles instead: add <file> /Sheet1 --type table --prop style=medium2",
            };

        var parts = rangeRef.Split(':');
        if (parts.Length != 2)
            throw new Core.CliException($"stylepreset needs a rectangular range path like /Sheet1/range[B2:T6] (got '{rangeRef}').")
            { Code = "invalid_argument" };
        var (startCol, startRow) = ParseCellReference(parts[0]);
        var (endCol, endRow) = ParseCellReference(parts[1]);
        var startColIdx = ColumnNameToIndex(startCol);
        var endColIdx = ColumnNameToIndex(endCol);

        var rowCount = endRow - startRow + 1;
        if (rowCount > 1000)
            throw new Core.CliException(
                $"table_banded expands to one fill per row and {rangeRef} spans {rowCount} rows (cap 1000).")
            {
                Code = "invalid_argument",
                Suggestion = "use add <file> /Sheet1 --type table --prop style=medium2 — a ListObject bands rows natively and scales with the table",
            };

        var accent1 = ResolveThemeAccent1();
        var ws = GetSheet(worksheet);
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            ws.Append(sheetData);
        }
        var sheetDataBackup = (SheetData)sheetData.CloneNode(true);
        var unsupported = new List<string>();
        try
        {
            for (var row = startRow; row <= endRow; row++)
            {
                // table_banded: odd range-rows (1st, 3rd, …) carry the stripe.
                var bundle = preset.OddRowProps != null
                    ? ((row - startRow) % 2 == 0 ? preset.OddRowProps : null)
                    : preset.Props;
                if (bundle == null) continue;
                var props = new Dictionary<string, string>(explicitProps, StringComparer.OrdinalIgnoreCase);
                foreach (var (key, value) in bundle)
                    props[key] = ResolveAccentTokens(value, accent1);
                for (var colIdx = startColIdx; colIdx <= endColIdx; colIdx++)
                {
                    var cellRef = $"{IndexToColumnName(colIdx)}{row}";
                    var cell = FindOrCreateCell(sheetData, cellRef);
                    var cellUnsupported = ApplyCellProperties(cell, worksheet, props);
                    PruneEmptyCell(cell);
                    if (row == startRow && colIdx == startColIdx)
                        unsupported.AddRange(cellUnsupported);
                }
            }
        }
        catch
        {
            ws.ReplaceChild(sheetDataBackup, sheetData);
            InvalidateRowIndex();
            throw;
        }
        return unsupported;
    }

    /// <summary>Replace the accent tokens in a preset value with the resolved
    /// theme hex. Anything without a token passes through untouched.</summary>
    private static string ResolveAccentTokens(string value, string accent1) => value switch
    {
        "ACCENT1" => accent1,
        "ACCENT1_TINT" => BlendTowardWhite(accent1, 0.15),
        _ => value,
    };

    /// <summary>Theme accent1 hex (RRGGBB) or the Office fallback.</summary>
    private string ResolveThemeAccent1()
        => Core.ThemeHandler.TryReadColorSlot(_doc.WorkbookPart?.ThemePart, "accent1") is { Length: > 0 } hex
            ? hex.TrimStart('#').ToUpperInvariant()
            : FallbackAccent1;

    /// <summary>
    /// Blend a hex color toward white. <paramref name="colorFraction"/> is the
    /// share of the original color (0.15 = light tint used for banding/card
    /// fills); per-channel and rounding are fixed so the result is
    /// reproducible byte-for-byte by the equivalence test.
    /// </summary>
    internal static string BlendTowardWhite(string hex, double colorFraction)
    {
        var v = hex.TrimStart('#');
        if (v.Length != 6 || !int.TryParse(v, System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture, out var rgb))
            return hex;
        var r = (rgb >> 16) & 0xFF;
        var g = (rgb >> 8) & 0xFF;
        var b = rgb & 0xFF;
        static int Blend(int c, double f) => (int)Math.Round(c * f + 255 * (1 - f), MidpointRounding.AwayFromZero);
        return $"{Blend(r, colorFraction):X2}{Blend(g, colorFraction):X2}{Blend(b, colorFraction):X2}";
    }
}
