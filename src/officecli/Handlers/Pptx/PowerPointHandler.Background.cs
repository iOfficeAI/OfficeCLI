// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Slide Background ====================

    /// <summary>
    /// Apply a background to a slide.
    ///
    /// Supported values for the "background" property:
    ///   RRGGBB               solid color        e.g. "FF0000"
    ///   none / transparent   remove background
    ///   C1-C2                gradient           e.g. "FF0000-0000FF"
    ///   C1-C2-angle          gradient + angle   e.g. "FF0000-0000FF-45"
    ///   C1-C2-C3             3-stop gradient    e.g. "FF0000-FFFF00-0000FF"
    ///   image:path           image fill         e.g. "image:/tmp/bg.png"
    /// </summary>
    private static void ApplySlideBackground(SlidePart slidePart, string value)
    {
        // Normalize alternative gradient format: "LINEAR;C1;C2;angle" → "C1-C2-angle"
        value = NormalizeGradientValue(value);

        var slide = GetSlide(slidePart);
        var cSld = slide.CommonSlideData
            ?? throw new InvalidOperationException("Slide has no CommonSlideData");

        // Remove any existing background element
        cSld.Background?.Remove();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("transparent", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("clear", StringComparison.OrdinalIgnoreCase))
            return;

        var bg = new Background();
        var bgPr = new BackgroundProperties();

        if (value.StartsWith("image:", StringComparison.OrdinalIgnoreCase))
        {
            var imagePath = value[6..].Trim();
            ApplyBackgroundImageFill(bgPr, slidePart, imagePath);
        }
        else if (value.StartsWith("radial:", StringComparison.OrdinalIgnoreCase) ||
                 value.StartsWith("path:", StringComparison.OrdinalIgnoreCase))
        {
            // Validate that radial:/path: prefix has valid color data
            if (!IsGradientColorString(value))
                throw new ArgumentException(
                    $"Invalid gradient specification: '{value}'. " +
                    "Radial/path gradients require at least 2 hex colors, e.g. 'radial:FF0000-0000FF'");
            bgPr.Append(BuildGradientFill(value));
        }
        else if (IsGradientColorString(value))
        {
            bgPr.Append(BuildGradientFill(value));
        }
        else
        {
            bgPr.Append(BuildSolidFill(value));
        }

        bg.Append(bgPr);

        // Insert before ShapeTree — schema order: p:bg → p:spTree
        var shapeTree = cSld.ShapeTree;
        if (shapeTree != null)
            cSld.InsertBefore(bg, shapeTree);
        else
            cSld.PrependChild(bg);
    }

    private static void ApplyBackgroundImageFill(
        BackgroundProperties bgPr, SlidePart slidePart, string imagePath)
    {
        if (!File.Exists(imagePath))
            throw new ArgumentException($"Image file not found: {imagePath}");

        var ext = Path.GetExtension(imagePath).ToLowerInvariant();
        var partType = ext switch
        {
            ".png"          => ImagePartType.Png,
            ".jpg" or ".jpeg" => ImagePartType.Jpeg,
            ".gif"          => ImagePartType.Gif,
            ".bmp"          => ImagePartType.Bmp,
            ".tif" or ".tiff" => ImagePartType.Tiff,
            ".emf"          => ImagePartType.Emf,
            ".wmf"          => ImagePartType.Wmf,
            _ => throw new ArgumentException(
                $"Unsupported image format: {ext}. Supported: png, jpg, gif, bmp, tif, emf, wmf")
        };

        var imagePart = slidePart.AddImagePart(partType);
        using (var stream = File.OpenRead(imagePath))
            imagePart.FeedData(stream);
        var relId = slidePart.GetIdOfPart(imagePart);

        var blipFill = new Drawing.BlipFill();
        blipFill.Append(new Drawing.Blip { Embed = relId });
        blipFill.Append(new Drawing.Stretch(new Drawing.FillRectangle()));
        bgPr.Append(blipFill);
    }

    // ==================== Read back ====================

    /// <summary>
    /// Populate Format["background"] on a slide DocumentNode.
    /// Values mirror the input format: hex for solid, "C1-C2[-angle]" for gradient, "image" for blip.
    /// </summary>
    private static void ReadSlideBackground(Slide slide, DocumentNode node)
    {
        var bgPr = slide.CommonSlideData?.Background?.BackgroundProperties;
        if (bgPr == null) return;

        var solidFill = bgPr.GetFirstChild<Drawing.SolidFill>();
        var gradFill  = bgPr.GetFirstChild<Drawing.GradientFill>();
        var blipFill  = bgPr.GetFirstChild<Drawing.BlipFill>();

        if (solidFill != null)
        {
            var bgColor = ReadColorFromFill(solidFill);
            if (bgColor != null) node.Format["background"] = bgColor;
        }
        else if (gradFill != null)
        {
            var stops = gradFill.GradientStopList?.Elements<Drawing.GradientStop>()
                .Select(gs => gs.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value ?? "?")
                .ToList();
            if (stops?.Count > 0)
            {
                var pathGrad = gradFill.GetFirstChild<Drawing.PathGradientFill>();
                if (pathGrad != null)
                {
                    var fillRect = pathGrad.GetFirstChild<Drawing.FillToRectangle>();
                    var focus = "center";
                    if (fillRect != null)
                    {
                        var fl = fillRect.Left?.Value ?? 50000;
                        var ft = fillRect.Top?.Value ?? 50000;
                        focus = (fl, ft) switch
                        {
                            (0, 0) => "tl",
                            ( >= 100000, 0) => "tr",
                            (0, >= 100000) => "bl",
                            ( >= 100000, >= 100000) => "br",
                            _ => "center"
                        };
                    }
                    node.Format["background"] = $"radial:{string.Join("-", stops)}-{focus}";
                }
                else
                {
                    var gradStr = string.Join("-", stops);
                    var linear = gradFill.GetFirstChild<Drawing.LinearGradientFill>();
                    if (linear?.Angle?.HasValue == true)
                        gradStr += $"-{linear.Angle.Value / 60000}";
                    node.Format["background"] = gradStr;
                }
            }
        }
        else if (blipFill != null)
        {
            node.Format["background"] = "image";
        }
    }

    // ==================== Helpers ====================

    /// <summary>
    /// Normalize alternative gradient formats to the canonical "-" separated form.
    /// Handles: "LINEAR;C1;C2;angle" → "C1-C2-angle", "RADIAL;C1;C2" → "radial:C1-C2"
    /// </summary>
    private static string NormalizeGradientValue(string value)
    {
        // Detect semicolon-separated format: TYPE;C1;C2[;angle/focus]
        if (!value.Contains(';')) return value;

        var parts = value.Split(';');
        if (parts.Length < 3) return value;

        var type = parts[0].Trim().ToUpperInvariant();
        var colorAndParams = parts.Skip(1).Select(p => p.Trim()).ToArray();

        return type switch
        {
            "LINEAR" => string.Join("-", colorAndParams),
            "RADIAL" => "radial:" + string.Join("-", colorAndParams),
            "PATH" => "path:" + string.Join("-", colorAndParams),
            _ => value // unknown type, leave as-is
        };
    }

    /// <summary>
    /// Returns true if value looks like a gradient color string ("RRGGBB-RRGGBB[-angle]").
    /// </summary>
    private static bool IsGradientColorString(string value)
    {
        // Handle radial:/path: prefix — must have color data after prefix
        var v = value;
        if (v.StartsWith("radial:", StringComparison.OrdinalIgnoreCase))
        {
            var after = v[7..];
            return after.Length > 0 && after.Split('-').Any(p => IsHexColorString(p));
        }
        if (v.StartsWith("path:", StringComparison.OrdinalIgnoreCase))
        {
            var after = v[5..];
            return after.Length > 0 && after.Split('-').Any(p => IsHexColorString(p));
        }

        var parts = v.Split('-');
        return parts.Length >= 2 && IsHexColorString(parts[0].TrimStart('#'));
    }

    private static bool IsHexColorString(string s)
    {
        s = s.TrimStart('#');
        return (s.Length == 6 || s.Length == 8) &&
               s.All(c => char.IsAsciiHexDigit(c));
    }

    /// <summary>
    /// Build a GradientFill element from a color string.
    /// Shared by both shape gradient and slide background gradient.
    ///
    /// Linear:  "C1-C2", "C1-C2-angle", "C1-C2-C3[-angle]"
    /// Radial:  "radial:C1-C2", "radial:C1-C2-tl" (focus: tl/tr/bl/br/center)
    /// Path:    "path:C1-C2", "path:C1-C2-tl"
    /// </summary>
    internal static Drawing.GradientFill BuildGradientFill(string value)
    {
        // Check for radial/path prefix
        string? gradientType = null;
        string colorSpec = value;

        if (value.StartsWith("radial:", StringComparison.OrdinalIgnoreCase))
        {
            gradientType = "radial";
            colorSpec = value[7..];
        }
        else if (value.StartsWith("path:", StringComparison.OrdinalIgnoreCase))
        {
            gradientType = "path";
            colorSpec = value[5..];
        }

        var parts = colorSpec.Split('-');
        if (parts.Length < 2)
            throw new ArgumentException(
                "Gradient requires at least 2 colors separated by '-', e.g. FF0000-0000FF");

        var colorParts = parts.ToList();
        string? focusPoint = null;
        int angle = 5400000; // default 90° = top→bottom

        if (gradientType != null)
        {
            // For radial/path: last segment may be a focus keyword (tl/tr/bl/br/center)
            var last = colorParts.Last().ToLowerInvariant();
            if (last is "tl" or "tr" or "bl" or "br" or "center" or "c")
            {
                focusPoint = last;
                colorParts.RemoveAt(colorParts.Count - 1);
            }
        }
        else
        {
            // For linear: last segment is angle if it's a short integer
            if (colorParts.Count >= 2 &&
                int.TryParse(colorParts.Last(), out var angleDeg) &&
                colorParts.Last().Length <= 3)
            {
                angle = angleDeg * 60000;
                colorParts.RemoveAt(colorParts.Count - 1);
            }
        }

        // If only one color remains after removing angle/focus, duplicate it
        if (colorParts.Count == 1)
            colorParts.Add(colorParts[0]);

        var gradFill = new Drawing.GradientFill();
        var gsLst = new Drawing.GradientStopList();

        for (int i = 0; i < colorParts.Count; i++)
        {
            var cp = colorParts[i];
            int pos;
            var atIdx = cp.IndexOf('@');
            if (atIdx >= 0 && int.TryParse(cp[(atIdx + 1)..], out var pct))
            {
                pos = Math.Clamp(pct, 0, 100) * 1000;
                cp = cp[..atIdx];
            }
            else
            {
                pos = colorParts.Count == 1
                    ? 0
                    : (int)((long)i * 100000 / (colorParts.Count - 1));
            }
            var gs = new Drawing.GradientStop { Position = pos };
            gs.AppendChild(BuildColorElement(cp));
            gsLst.AppendChild(gs);
        }

        gradFill.AppendChild(gsLst);

        if (gradientType is "radial" or "path")
        {
            // Build path gradient fill with fillToRect controlling the focal point
            var (l, t, r, b) = (focusPoint ?? "center") switch
            {
                "tl" => (0, 0, 100000, 100000),       // top-left focal point
                "tr" => (100000, 0, 0, 100000),        // top-right
                "bl" => (0, 100000, 100000, 0),        // bottom-left
                "br" => (100000, 100000, 0, 0),        // bottom-right
                _ => (50000, 50000, 50000, 50000)       // center
            };

            var pathFill = new Drawing.PathGradientFill { Path = Drawing.PathShadeValues.Circle };
            pathFill.AppendChild(new Drawing.FillToRectangle
            {
                Left = l, Top = t, Right = r, Bottom = b
            });
            gradFill.AppendChild(pathFill);
        }
        else
        {
            gradFill.AppendChild(new Drawing.LinearGradientFill { Angle = angle, Scaled = true });
        }

        return gradFill;
    }
}
