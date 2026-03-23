// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    /// <summary>
    /// Apply outer shadow effect to ShapeProperties.
    /// Format: "COLOR" or "COLOR-BLUR-ANGLE-DIST" or "COLOR-BLUR-ANGLE-DIST-OPACITY"
    ///   COLOR: hex (e.g. 000000)
    ///   BLUR: blur radius in points, default 4
    ///   ANGLE: direction in degrees, default 45
    ///   DIST: distance in points, default 3
    ///   OPACITY: 0-100 percent, default 40
    /// Examples: "000000", "000000-6-315-4-50", "none"
    /// </summary>
    private static void ApplyShadow(ShapeProperties spPr, string value)
    {
        var effectList = EnsureEffectList(spPr);
        effectList.RemoveAllChildren<Drawing.OuterShadow>();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) || value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            if (!effectList.HasChildren) spPr.RemoveChild(effectList);
            return;
        }

        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("Shadow value cannot be empty. Use 'none' to remove shadow.");

        InsertEffectInOrder(effectList, BuildOuterShadow(value));
    }

    /// <summary>
    /// Apply glow effect to ShapeProperties.
    /// Format: "COLOR" or "COLOR-RADIUS" or "COLOR-RADIUS-OPACITY"
    ///   COLOR: hex (e.g. 0070FF)
    ///   RADIUS: glow radius in points, default 8
    ///   OPACITY: 0-100 percent, default 75
    /// Examples: "0070FF", "FF0000-10", "00B0F0-6-60", "none"
    /// </summary>
    private static void ApplyGlow(ShapeProperties spPr, string value)
    {
        var effectList = EnsureEffectList(spPr);
        effectList.RemoveAllChildren<Drawing.Glow>();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) || value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            if (!effectList.HasChildren) spPr.RemoveChild(effectList);
            return;
        }

        InsertEffectInOrder(effectList, BuildGlow(value));
    }

    /// <summary>
    /// Check if a shape has no fill (transparent background).
    /// </summary>
    private static bool IsNoFillShape(ShapeProperties spPr)
    {
        return spPr.GetFirstChild<Drawing.NoFill>() != null;
    }

    /// <summary>
    /// Build an OuterShadow element from the shadow value string.
    /// </summary>
    private static Drawing.OuterShadow BuildOuterShadow(string value)
        => OfficeCli.Core.DrawingEffectsHelper.BuildOuterShadow(value, BuildColorElement);

    private static Drawing.Glow BuildGlow(string value)
        => OfficeCli.Core.DrawingEffectsHelper.BuildGlow(value, BuildColorElement);

    /// <summary>
    /// Get or create EffectList in correct schema position within RunProperties.
    /// CT_TextCharacterProperties schema order: ln → fill → effectLst → highlight → uLnTx/uLn → uFillTx/uFill → latin → ea → cs → sym → hlinkClick → hlinkMouseOver → extLst
    /// </summary>
    private static void InsertFillInRunProperties(Drawing.RunProperties rPr, DocumentFormat.OpenXml.OpenXmlElement fillElement)
        => OfficeCli.Core.DrawingEffectsHelper.InsertFillInRunProperties(rPr, fillElement);

    private static void ApplyTextShadow(Drawing.Run run, string value)
        => OfficeCli.Core.DrawingEffectsHelper.ApplyTextEffect<Drawing.OuterShadow>(run, value, () => BuildOuterShadow(value));

    private static void ApplyTextGlow(Drawing.Run run, string value)
        => OfficeCli.Core.DrawingEffectsHelper.ApplyTextEffect<Drawing.Glow>(run, value, () => BuildGlow(value));

    /// <summary>
    /// Apply reflection effect to ShapeProperties.
    /// Format: "TYPE" where TYPE is one of:
    ///   tight / small  — tight reflection, touching (stA=52000 endA=300 endPos=55000)
    ///   half           — half reflection (stA=52000 endA=300 endPos=90000)
    ///   full           — full reflection (stA=52000 endA=300 endPos=100000)
    ///   true           — alias for half
    ///   none / false   — remove reflection
    /// </summary>
    private static void ApplyReflection(ShapeProperties spPr, string value)
    {
        var effectList = EnsureEffectList(spPr);
        effectList.RemoveAllChildren<Drawing.Reflection>();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) || value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            if (!effectList.HasChildren) spPr.RemoveChild(effectList);
            return;
        }

        // endPos controls how much of the shape is reflected
        int endPos = value.ToLowerInvariant() switch
        {
            "tight" or "small" => 55000,
            "true" or "half"   => 90000,
            "full"             => 100000,
            _ => int.TryParse(value, out var pct) ? (int)Math.Min((long)pct * 1000, 100000) : 90000
        };

        var reflection = new Drawing.Reflection
        {
            BlurRadius      = 6350,
            StartOpacity    = 52000,
            StartPosition   = 0,
            EndAlpha        = 300,
            EndPosition     = endPos,
            Distance        = 0,
            Direction       = 5400000,  // 90° — downward
            VerticalRatio   = -100000,  // flip vertically
            Alignment       = Drawing.RectangleAlignmentValues.BottomLeft,
            RotateWithShape = false
        };
        InsertEffectInOrder(effectList, reflection);
    }

    /// <summary>
    /// Apply soft edge effect to ShapeProperties.
    /// Value: radius in points (e.g. "5") or "none" to remove.
    /// </summary>
    private static void ApplySoftEdge(ShapeProperties spPr, string value)
    {
        var effectList = EnsureEffectList(spPr);
        effectList.RemoveAllChildren<Drawing.SoftEdge>();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) || value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            if (!effectList.HasChildren) spPr.RemoveChild(effectList);
            return;
        }

        if (!double.TryParse(value, System.Globalization.CultureInfo.InvariantCulture, out var radiusPt) || double.IsNaN(radiusPt) || double.IsInfinity(radiusPt))
            throw new ArgumentException($"Invalid 'softedge' value '{value}'. Expected a finite numeric radius in points.");
        InsertEffectInOrder(effectList, new Drawing.SoftEdge { Radius = (long)(radiusPt * 12700) });
    }

    private static void ApplyTextReflection(Drawing.Run run, string value)
        => OfficeCli.Core.DrawingEffectsHelper.ApplyTextEffect<Drawing.Reflection>(run, value,
            () => OfficeCli.Core.DrawingEffectsHelper.BuildReflection(value));

    private static void ApplyTextSoftEdge(Drawing.Run run, string value)
        => OfficeCli.Core.DrawingEffectsHelper.ApplyTextEffect<Drawing.SoftEdge>(run, value,
            () => OfficeCli.Core.DrawingEffectsHelper.BuildSoftEdge(value));

    /// <summary>
    /// Apply 3D rotation (scene3d) to ShapeProperties.
    /// Format: "rotX,rotY,rotZ" in degrees (e.g. "45,30,0")
    /// </summary>
    private static void Apply3DRotation(ShapeProperties spPr, string value)
    {
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            var existing = spPr.GetFirstChild<Drawing.Scene3DType>();
            if (existing != null) spPr.RemoveChild(existing);
            return;
        }

        var parts = value.Split(',');
        if (parts.Length < 3)
            throw new ArgumentException($"Invalid '3drotation' value: '{value}'. Expected 3 components as 'rotX,rotY,rotZ' (e.g. '45,30,0').");
        if (!double.TryParse(parts[0].Trim(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var rotX) || double.IsNaN(rotX) || double.IsInfinity(rotX))
            throw new ArgumentException($"Invalid '3drotation' value: '{value}'. Expected finite degrees as 'rotX,rotY,rotZ' (e.g. '45,30,0').");
        if (!double.TryParse(parts[1].Trim(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var ry) || double.IsNaN(ry) || double.IsInfinity(ry))
            throw new ArgumentException($"Invalid '3drotation' rotY value: '{parts[1].Trim()}'. Expected a finite number.");
        if (!double.TryParse(parts[2].Trim(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var rz) || double.IsNaN(rz) || double.IsInfinity(rz))
            throw new ArgumentException($"Invalid '3drotation' rotZ value: '{parts[2].Trim()}'. Expected a finite number.");
        var rotY = ry;
        var rotZ = rz;

        var scene3d = EnsureScene3D(spPr);
        var camera = scene3d.Camera!;
        camera.Rotation = new Drawing.Rotation
        {
            Latitude = NormalizeDegrees60k(rotX),
            Longitude = NormalizeDegrees60k(rotY),
            Revolution = NormalizeDegrees60k(rotZ)
        };
    }

    /// <summary>
    /// Normalize degrees to OOXML 60000ths-of-a-degree range [0, 21600000).
    /// Accepts negative values (e.g. -20° → 340° → 20400000).
    /// </summary>
    private static int NormalizeDegrees60k(double degrees)
    {
        var val = (int)(degrees * 60000);
        const int full = 360 * 60000; // 21600000
        val %= full;
        if (val < 0) val += full;
        return val;
    }

    /// <summary>
    /// Apply a single 3D rotation axis.
    /// </summary>
    private static void Apply3DRotationAxis(ShapeProperties spPr, string axis, string value)
    {
        var scene3d = EnsureScene3D(spPr);
        var camera = scene3d.Camera!;
        var rot = camera.Rotation ?? (camera.Rotation = new Drawing.Rotation());
        if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var degVal) || double.IsNaN(degVal) || double.IsInfinity(degVal))
            throw new ArgumentException($"Invalid '3drotation.{axis}' value: '{value}'. Expected a finite number in degrees.");
        var deg = NormalizeDegrees60k(degVal);

        switch (axis)
        {
            case "x": rot.Latitude = deg; break;
            case "y": rot.Longitude = deg; break;
            case "z": rot.Revolution = deg; break;
        }
    }

    /// <summary>
    /// Apply bevel to ShapeProperties (top or bottom).
    /// Format: "preset" or "preset-width-height" (width/height in points)
    /// Presets: circle, relaxedInset, cross, coolSlant, angle, softRound, convex,
    ///          slope, divot, riblet, hardEdge, artDeco
    /// Examples: "circle", "circle-6-6", "none"
    /// </summary>
    private static void ApplyBevel(ShapeProperties spPr, string value, bool top)
    {
        var sp3d = spPr.GetFirstChild<Drawing.Shape3DType>();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            if (sp3d != null)
            {
                if (top) { sp3d.BevelTop = null; }
                else { sp3d.BevelBottom = null; }
                if (sp3d.BevelTop == null && sp3d.BevelBottom == null &&
                    (sp3d.ExtrusionHeight == null || sp3d.ExtrusionHeight.Value == 0))
                    spPr.RemoveChild(sp3d);
            }
            return;
        }

        sp3d ??= EnsureShape3D(spPr);
        // Normalize alternative separator: "preset;width;height" → "preset-width-height"
        value = value.Replace(';', '-');
        var bevelParts = value.Split('-');
        var preset = ParseBevelPreset(bevelParts[0].Trim());
        long w = 76200L, h;
        if (bevelParts.Length > 1)
        {
            if (!double.TryParse(bevelParts[1].Trim(), System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var wPt) || double.IsNaN(wPt) || double.IsInfinity(wPt))
                throw new ArgumentException($"Invalid bevel width: '{bevelParts[1]}'. Expected a finite number in points. Format: PRESET[-WIDTH[-HEIGHT]]");
            w = (long)(wPt * 12700);
        }
        if (bevelParts.Length > 2)
        {
            if (!double.TryParse(bevelParts[2].Trim(), System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var hPt) || double.IsNaN(hPt) || double.IsInfinity(hPt))
                throw new ArgumentException($"Invalid bevel height: '{bevelParts[2]}'. Expected a finite number in points. Format: PRESET[-WIDTH[-HEIGHT]]");
            h = (long)(hPt * 12700);
        }
        else h = w;

        if (top)
        {
            sp3d.BevelTop = new Drawing.BevelTop { Width = w, Height = h, Preset = preset };
        }
        else
        {
            sp3d.BevelBottom = new Drawing.BevelBottom { Width = w, Height = h, Preset = preset };
        }
    }

    /// <summary>
    /// Apply 3D extrusion depth in points.
    /// </summary>
    private static void Apply3DDepth(ShapeProperties spPr, string value)
    {
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) || value == "0")
        {
            var sp3d = spPr.GetFirstChild<Drawing.Shape3DType>();
            if (sp3d != null) { sp3d.ExtrusionHeight = 0; }
            return;
        }

        var sp3dEl = EnsureShape3D(spPr);
        if (!double.TryParse(value, System.Globalization.CultureInfo.InvariantCulture, out var depthPt) || double.IsNaN(depthPt) || double.IsInfinity(depthPt))
            throw new ArgumentException($"Invalid '3ddepth' value '{value}'. Expected a finite numeric depth in points.");
        sp3dEl.ExtrusionHeight = (long)(depthPt * 12700);
    }

    /// <summary>
    /// Apply 3D material preset.
    /// </summary>
    private static void Apply3DMaterial(ShapeProperties spPr, string value)
    {
        var sp3d = EnsureShape3D(spPr);
        sp3d.PresetMaterial = ParseMaterial(value);
    }

    /// <summary>
    /// Apply light rig preset to scene3d.
    /// </summary>
    private static void ApplyLightRig(ShapeProperties spPr, string value)
    {
        var scene3d = EnsureScene3D(spPr);
        scene3d.LightRig!.Rig = ParseLightRig(value);
    }

    // --- Helper methods ---

    /// <summary>
    /// Schema order for CT_EffectList children:
    /// blur → fillOverlay → glow → innerShdw → outerShdw → prstShdw → reflection → softEdge
    /// </summary>
    private static readonly Type[] EffectListChildOrder =
    [
        typeof(Drawing.Blur),
        typeof(Drawing.FillOverlay),
        typeof(Drawing.Glow),
        typeof(Drawing.InnerShadow),
        typeof(Drawing.OuterShadow),
        typeof(Drawing.PresetShadow),
        typeof(Drawing.Reflection),
        typeof(Drawing.SoftEdge),
    ];

    /// <summary>
    /// Insert an effect element into EffectList at the correct schema position.
    /// </summary>
    private static void InsertEffectInOrder(Drawing.EffectList effectList, DocumentFormat.OpenXml.OpenXmlElement element)
    {
        var targetIdx = Array.IndexOf(EffectListChildOrder, element.GetType());
        // Find the first existing child that should come after this element
        foreach (var child in effectList.ChildElements)
        {
            var childIdx = Array.IndexOf(EffectListChildOrder, child.GetType());
            if (childIdx > targetIdx)
            {
                effectList.InsertBefore(element, child);
                return;
            }
        }
        effectList.AppendChild(element);
    }

    /// <summary>
    /// Get or create EffectList in correct schema position.
    /// Schema order: fill → ln → effectLst → scene3d → sp3d → extLst
    /// </summary>
    private static Drawing.EffectList EnsureEffectList(ShapeProperties spPr)
    {
        var effectList = spPr.GetFirstChild<Drawing.EffectList>();
        if (effectList != null) return effectList;

        effectList = new Drawing.EffectList();
        // Insert before scene3d/sp3d/extLst if they exist
        var insertBefore = (DocumentFormat.OpenXml.OpenXmlElement?)spPr.GetFirstChild<Drawing.Scene3DType>()
            ?? (DocumentFormat.OpenXml.OpenXmlElement?)spPr.GetFirstChild<Drawing.Shape3DType>()
            ?? spPr.GetFirstChild<Drawing.ShapePropertiesExtensionList>();
        if (insertBefore != null)
            spPr.InsertBefore(effectList, insertBefore);
        else
            spPr.AppendChild(effectList);
        return effectList;
    }

    /// <summary>
    /// Get or create Outline in correct schema position.
    /// Schema order: fill → ln → effectLst → scene3d → sp3d → extLst
    /// </summary>
    private static Drawing.Outline EnsureOutline(ShapeProperties spPr)
    {
        var outline = spPr.GetFirstChild<Drawing.Outline>();
        if (outline != null) return outline;

        outline = new Drawing.Outline();
        // Insert before effectLst/scene3d/sp3d/extLst if they exist
        var insertBefore = (DocumentFormat.OpenXml.OpenXmlElement?)spPr.GetFirstChild<Drawing.EffectList>()
            ?? (DocumentFormat.OpenXml.OpenXmlElement?)spPr.GetFirstChild<Drawing.Scene3DType>()
            ?? (DocumentFormat.OpenXml.OpenXmlElement?)spPr.GetFirstChild<Drawing.Shape3DType>()
            ?? spPr.GetFirstChild<Drawing.ShapePropertiesExtensionList>();
        if (insertBefore != null)
            spPr.InsertBefore(outline, insertBefore);
        else
            spPr.AppendChild(outline);
        return outline;
    }

    private static Drawing.Scene3DType EnsureScene3D(ShapeProperties spPr)
    {
        var scene3d = spPr.GetFirstChild<Drawing.Scene3DType>();
        if (scene3d != null) return scene3d;

        scene3d = new Drawing.Scene3DType(
            new Drawing.Camera { Preset = Drawing.PresetCameraValues.OrthographicFront },
            new Drawing.LightRig { Rig = Drawing.LightRigValues.ThreePoints, Direction = Drawing.LightRigDirectionValues.Top }
        );
        // Schema order: effectLst → scene3d → sp3d → extLst
        // Insert before sp3d if it exists, otherwise append
        var sp3d = spPr.GetFirstChild<Drawing.Shape3DType>();
        if (sp3d != null)
            spPr.InsertBefore(scene3d, sp3d);
        else
            spPr.AppendChild(scene3d);
        return scene3d;
    }

    private static Drawing.Shape3DType EnsureShape3D(ShapeProperties spPr)
    {
        var sp3d = spPr.GetFirstChild<Drawing.Shape3DType>();
        if (sp3d != null) return sp3d;

        sp3d = new Drawing.Shape3DType();
        // Schema order: scene3d → sp3d → extLst
        // Insert before extLst if it exists, otherwise append
        var extLst = spPr.GetFirstChild<Drawing.ShapePropertiesExtensionList>();
        if (extLst != null)
            spPr.InsertBefore(sp3d, extLst);
        else
            spPr.AppendChild(sp3d);
        return sp3d;
    }

    private static Drawing.BevelPresetValues ParseBevelPreset(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "circle" => Drawing.BevelPresetValues.Circle,
            "relaxedinset" => Drawing.BevelPresetValues.RelaxedInset,
            "cross" => Drawing.BevelPresetValues.Cross,
            "coolslant" => Drawing.BevelPresetValues.CoolSlant,
            "angle" => Drawing.BevelPresetValues.Angle,
            "softround" => Drawing.BevelPresetValues.SoftRound,
            "convex" => Drawing.BevelPresetValues.Convex,
            "slope" => Drawing.BevelPresetValues.Slope,
            "divot" => Drawing.BevelPresetValues.Divot,
            "riblet" => Drawing.BevelPresetValues.Riblet,
            "hardedge" => Drawing.BevelPresetValues.HardEdge,
            "artdeco" => Drawing.BevelPresetValues.ArtDeco,
            _ => throw new ArgumentException($"Invalid bevel preset: '{value}'. Valid values: circle, relaxedinset, cross, coolslant, angle, softround, convex, slope, divot, riblet, hardedge, artdeco.")
        };
    }

    private static T WarnAndDefault<T>(string value, T defaultVal, string paramName, string validValues)
    {
        Console.Error.WriteLine($"Warning: unrecognized {paramName} '{value}', using default. Valid values: {validValues}");
        return defaultVal;
    }

    private static Drawing.PresetMaterialTypeValues ParseMaterial(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "warmmatte" => Drawing.PresetMaterialTypeValues.WarmMatte,
            "plastic" => Drawing.PresetMaterialTypeValues.Plastic,
            "metal" => Drawing.PresetMaterialTypeValues.Metal,
            "dkedge" or "darkedge" => Drawing.PresetMaterialTypeValues.DarkEdge,
            "softedge" => Drawing.PresetMaterialTypeValues.SoftEdge,
            "flat" => Drawing.PresetMaterialTypeValues.Flat,
            "wire" or "wireframe" => Drawing.PresetMaterialTypeValues.LegacyWireframe,
            "powder" => Drawing.PresetMaterialTypeValues.Powder,
            "translucentpowder" => Drawing.PresetMaterialTypeValues.TranslucentPowder,
            "clear" => Drawing.PresetMaterialTypeValues.Clear,
            "softmetal" => Drawing.PresetMaterialTypeValues.SoftMetal,
            "matte" => Drawing.PresetMaterialTypeValues.Matte,
            _ => throw new ArgumentException($"Invalid material value: '{value}'. Valid values: warmmatte, plastic, metal, darkedge, flat, wire, powder, translucentpowder, clear, softmetal, matte.")
        };
    }

    private static Drawing.LightRigValues ParseLightRig(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "threept" or "3pt" => Drawing.LightRigValues.ThreePoints,
            "balanced" => Drawing.LightRigValues.Balanced,
            "soft" => Drawing.LightRigValues.Soft,
            "harsh" => Drawing.LightRigValues.Harsh,
            "flood" => Drawing.LightRigValues.Flood,
            "contrasting" => Drawing.LightRigValues.Contrasting,
            "morning" => Drawing.LightRigValues.Morning,
            "sunrise" => Drawing.LightRigValues.Sunrise,
            "sunset" => Drawing.LightRigValues.Sunset,
            "chilly" => Drawing.LightRigValues.Chilly,
            "freezing" => Drawing.LightRigValues.Freezing,
            "flat" => Drawing.LightRigValues.Flat,
            "twopt" or "2pt" => Drawing.LightRigValues.TwoPoints,
            "glow" => Drawing.LightRigValues.Glow,
            "brightroom" => Drawing.LightRigValues.BrightRoom,
            _ => throw new ArgumentException($"Invalid lighting value: '{value}'. Valid values: threept, balanced, soft, harsh, flood, contrasting, morning, sunrise, sunset, chilly, freezing, flat, twopt, glow, brightroom.")
        };
    }

    /// <summary>
    /// Format a bevel element as "preset-width-height" string for reading back.
    /// </summary>
    internal static string FormatBevel(Drawing.BevelType bevel)
    {
        var preset = bevel.Preset?.HasValue == true ? bevel.Preset.InnerText : "circle";
        var w = bevel.Width?.HasValue == true ? $"{bevel.Width.Value / 12700.0:0.##}" : "6";
        var h = bevel.Height?.HasValue == true ? $"{bevel.Height.Value / 12700.0:0.##}" : "6";
        return $"{preset}-{w}-{h}";
    }
}
