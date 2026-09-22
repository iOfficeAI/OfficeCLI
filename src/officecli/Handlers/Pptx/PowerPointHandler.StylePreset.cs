// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    /// <summary>
    /// One-command style application for PPTX shapes (the slide-level `preset`
    /// prop). Each preset is a curated bundle of RENDER-SAFE shape properties —
    /// solid fill, roundRect geometry + corner-radius adj, soft outer shadow,
    /// text color / font / size / bold — written through the exact same
    /// ApplyShapePropsCore path a normal per-shape `set` uses, so nothing the
    /// preset requests can silently render as the wrong thing.
    ///
    /// Colors, pairing and mood are sourced verbatim from the MIT-licensed,
    /// open-source "NextSlide" style preset collection (github.com/0xRafie/
    /// nextslide, STYLE_PRESETS.md). Fonts are intentionally mapped onto
    /// typefaces that ship with Office (Arab/Calibri/Georgia/…) instead of the
    /// original Google-Font names, because a font the viewer's machine does not
    /// have falls back to the theme default and loses the intended look —
    /// "broad compatibility" beats an exact pixel match for a reusable preset.
    ///
    /// Only solid fills are used. Gradient fills (and gradient text fills) can
    /// degrade to a single colour under some PowerPoint renderers, so presets
    /// deliberately avoid them to honour the "what you ask for is what renders"
    /// contract.
    /// </summary>
    private void ApplyStylePreset(SlidePart slidePart, string presetName, string? targets)
    {
        var props = TryGetPresetProps(presetName)
            ?? throw new ArgumentException(
                $"Unknown preset: '{presetName}'. " +
                $"Available: {string.Join(", ", PresetProps.Keys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase))}");

        var shapes = ResolveAlignTargets(slidePart, targets);
        if (shapes.Count == 0) return;

        foreach (var shape in shapes)
            ApplyShapePropsCore(slidePart, shape, props);
    }

    /// <summary>Resolve a preset name to its render-safe prop bundle, or null.</summary>
    private static Dictionary<string, string>? TryGetPresetProps(string name)
    {
        if (string.IsNullOrEmpty(name)) return null;
        return PresetProps.TryGetValue(name, out var bundle) ? bundle : null;
    }

    /// <summary>
    /// Registry of built-in style presets. Keys are case-insensitive; the value
    /// is the exact prop map handed to ApplyShapePropsCore for every target
    /// shape (which reads it without mutation, so one instance can be shared
    /// across all shapes on the slide). All property KEYS are the documented
    /// shape set-props (fill, color, geometry, adj, shadow, line, font, size,
    /// bold) — never raw XML.
    /// </summary>
    private static readonly Dictionary<string, Dictionary<string, string>> PresetProps =
        new(StringComparer.OrdinalIgnoreCase)
        {
            // NextSlide "Minimal Clean" — light card, thin primary border, soft slate shadow.
            ["minimal"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["fill"] = "F8FAFC",
                ["color"] = "1A1A2E",
                ["line"] = "3B82F6",
                ["geometry"] = "roundRect",
                ["adj"] = "adj:val 5000",
                ["shadow"] = "64748B",
                ["font"] = "Arial",
                ["size"] = "18",
            },

            // NextSlide "Corporate Pro" — white card, royal-blue hairline, navy ink.
            ["corporate"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["fill"] = "FFFFFF",
                ["color"] = "1E293B",
                ["line"] = "2563EB",
                ["geometry"] = "roundRect",
                ["adj"] = "adj:val 4000",
                ["shadow"] = "475569",
                ["font"] = "Calibri",
                ["size"] = "18",
            },

            // NextSlide "Pitch Deck" — deep-slate tech card, electric-green accent.
            ["pitch"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["fill"] = "1E293B",
                ["color"] = "F8FAFC",
                ["bold"] = "true",
                ["line"] = "22C55E",
                ["geometry"] = "roundRect",
                ["adj"] = "adj:val 8000",
                ["shadow"] = "000000",
                ["font"] = "Arial",
                ["size"] = "18",
            },

            // NextSlide "Bold Geometric" — solid red statement card, white ink.
            ["bold"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["fill"] = "DC2626",
                ["color"] = "FFFFFF",
                ["bold"] = "true",
                ["geometry"] = "roundRect",
                ["adj"] = "adj:val 4000",
                ["shadow"] = "78716C",
                ["font"] = "Arial Black",
                ["size"] = "18",
            },

            // NextSlide "Editorial" — warm ivory card, burgundy serif, gold hairline.
            ["editorial"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["fill"] = "FAF9F7",
                ["color"] = "991B1B",
                ["line"] = "B45309",
                ["geometry"] = "roundRect",
                ["adj"] = "adj:val 2000",
                ["shadow"] = "A3A3A3",
                ["font"] = "Georgia",
                ["size"] = "18",
            },

            // NextSlide "Teal Serenity" — teal card, mint ink, calm.
            ["teal"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["fill"] = "14B8A6",
                ["color"] = "F0FDFA",
                ["bold"] = "true",
                ["geometry"] = "roundRect",
                ["adj"] = "adj:val 6000",
                ["shadow"] = "0F766E",
                ["font"] = "Segoe UI",
                ["size"] = "18",
            },

            // NextSlide "Playful Pop" — vivid purple, very rounded, friendly.
            ["playful"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["fill"] = "7C3AED",
                ["color"] = "FFFFFF",
                ["bold"] = "true",
                ["geometry"] = "roundRect",
                ["adj"] = "adj:val 12000",
                ["shadow"] = "A78BFA",
                ["font"] = "Verdana",
                ["size"] = "18",
            },
        };
}