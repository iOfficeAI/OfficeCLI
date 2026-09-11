// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    /// <summary>
    /// One-command style application for Word paragraphs/table rows (the `preset`
    /// prop). Each preset is a curated bundle of RENDER-SAFE paragraph/run
    /// properties — solid shading, paragraph borders, keepNext/keepLines,
    /// run color/size/bold/italic/caps — written through the exact same
    /// <see cref="ApplyParagraphLevelProperty"/> / <see cref="ApplyRunFormatting"/>
    /// paths a normal per-element `set` uses, so nothing the preset requests can
    /// silently render as the wrong thing.
    ///
    /// The six presets mirror the six "design building blocks" taught in the
    /// <c>officecli-docx-design</c> skill (callout card, dark accent card, kicker,
    /// section color-band, pull quote, table header band), so an agent can apply
    /// a designed look with one command instead of hand-writing
    /// <c>shd=... pbdr=... keepLines=...</c> every time.
    ///
    /// Only solid fills are used. Gradient fills can degrade to a single colour
    /// under some renderers, so presets deliberately avoid them to honour the
    /// "what you ask for is what renders" contract (mirrors the PPTX
    /// PowerPointHandler.StylePreset design).
    /// </summary>
    private void ApplyParagraphPreset(Paragraph para, ParagraphProperties pProps, string presetName)
    {
        var bundle = TryGetPresetProps(presetName)
            ?? throw new ArgumentException(
                $"Unknown preset: '{presetName}'. " +
                $"Available: {string.Join(", ", PresetProps.Keys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase))}");

        foreach (var (key, value) in bundle)
        {
            // Paragraph-level keys route through the normal paragraph Set path.
            if (ApplyParagraphLevelProperty(pProps, key, value, LastSetWarnings))
                continue;

            // Run-level keys apply to each existing run (never create empty runs;
            // a paragraph with no runs stays empty — the caller supplies text).
            bool appliedToAnyRun = false;
            foreach (var run in para.Descendants<Run>())
            {
                if (ApplyRunFormatting(run.RunProperties ?? run.PrependChild(new RunProperties()), key, value))
                    appliedToAnyRun = true;
            }
            // Mirror paragraph-mark rPr so the run props survive a cursor blink /
            // next-typed-text (same reason the mark path exists in Add/Set).
            if (appliedToAnyRun)
                ApplyRunFormatting(EnsureMarkRunProperties(pProps), key, value);
        }
    }

    /// <summary>
    /// Apply a table-row preset: shade every cell (tcPr) and format the row's
    /// runs (e.g. table-header = white bold on the accent band). Mirrors the
    /// paragraph preset's render-safe contract.
    /// </summary>
    private void ApplyTableRowPreset(TableRow row, string presetName)
    {
        var bundle = TryGetPresetProps(presetName)
            ?? throw new ArgumentException(
                $"Unknown preset: '{presetName}'. " +
                $"Available: {string.Join(", ", PresetProps.Keys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase))}");

        foreach (var (key, value) in bundle)
        {
            if (key.Equals("shd", StringComparison.OrdinalIgnoreCase))
            {
                // Row-level shading → each cell's tcPr (shading lives on tcPr).
                foreach (var cell in row.Elements<TableCell>())
                {
                    var tcPr = cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties());
                    tcPr.RemoveAllChildren<Shading>();
                    tcPr.AppendChild(ParseShadingValue(value));
                }
                continue;
            }
            // Run-level keys (color/bold/...) apply to every run in the row.
            foreach (var run in row.Descendants<Run>())
                ApplyRunFormatting(run.RunProperties ?? run.PrependChild(new RunProperties()), key, value);
        }
    }

    /// <summary>Resolve a preset name to its render-safe prop bundle, or null.</summary>
    private static Dictionary<string, string>? TryGetPresetProps(string name)
    {
        if (string.IsNullOrEmpty(name)) return null;
        return PresetProps.TryGetValue(name, out var bundle) ? bundle : null;
    }

    /// <summary>
    /// Registry of built-in style presets. Keys are case-insensitive; the value
    /// is the exact prop map handed to the normal paragraph/run Set paths. All
    /// property KEYS are the documented paragraph/run set-props (shd, pbdr.*,
    /// keepNext, keepLines, spaceBefore/After, align, color, size, bold, italic,
    /// caps, font) — never raw XML. Palette follows the docx-design skill's
    /// accent-first scheme (E6A23C amber accent, 1F4E79 navy, light warm fills).
    /// </summary>
    private static readonly Dictionary<string, Dictionary<string, string>> PresetProps =
        new(StringComparer.OrdinalIgnoreCase)
        {
            // Callout card — light warm fill, thin amber border, stays whole.
            ["card"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["shd"] = "solid;FFF4E5",
                ["pbdr.all"] = "single;1pt;E6A23C",
                ["keepLines"] = "true",
                ["spaceAfter"] = "120",
            },

            // Dark accent card — navy fill, white bold text, 2pt navy border.
            ["card-dark"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["shd"] = "solid;1F4E79",
                ["color"] = "FFFFFF",
                ["bold"] = "true",
                ["pbdr.all"] = "single;2pt;1F4E79",
                ["keepLines"] = "true",
            },

            // Kicker — amber small caps lead line, glued to the following title.
            ["kicker"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["color"] = "E6A23C",
                ["caps"] = "true",
                ["size"] = "11",
                ["spaceAfter"] = "0",
                ["keepNext"] = "true",
            },

            // Section color-band — full-width amber rule.
            ["band"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["shd"] = "solid;E6A23C",
                ["spaceBefore"] = "240",
                ["spaceAfter"] = "0",
            },

            // Pull quote — centered italic navy serif, thick left rule.
            ["quote"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["italic"] = "true",
                ["color"] = "1F4E79",
                ["size"] = "16",
                ["pbdr.left"] = "single;12pt;1F4E79",
                ["align"] = "center",
            },

            // Table header band — navy fill + white bold text (row-level).
            ["table-header"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["shd"] = "solid;1F4E79",
                ["color"] = "FFFFFF",
                ["bold"] = "true",
            },
        };
}