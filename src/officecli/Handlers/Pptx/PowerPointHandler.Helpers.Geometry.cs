// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{

    private static long ParseEmu(string value) => Core.EmuConverter.ParseEmu(value);

    private static string FormatEmu(long emu) => Core.EmuConverter.FormatEmu(emu);

    private static string FormatLineWidth(long emu) => Core.EmuConverter.FormatLineWidth(emu);

    /// <summary>
    /// Parse SVG-like path syntax into a Drawing.CustomGeometry element.
    /// Format: "M x,y L x,y C x1,y1 x2,y2 x,y Q x1,y1 x,y Z"
    ///   M = moveTo, L = lineTo, C = cubicBezTo, Q = quadBezTo, A = arcTo, Z = close
    /// Coordinates use 0-100 relative space, internally scaled ×1000 to OOXML standard 0-100000.
    /// Example: "M 0,0 L 100,0 L 100,100 L 0,100 Z" (rectangle in 0-100 space)
    /// </summary>
    private static Drawing.CustomGeometry ParseCustomGeometry(string value)
    {
        var path = new Drawing.Path();

        // Parse SVG-like commands
        var tokens = value.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        long maxX = 0, maxY = 0;
        int i = 0;

        while (i < tokens.Length)
        {
            var cmd = tokens[i].ToUpperInvariant();
            i++;

            switch (cmd)
            {
                case "M":
                {
                    var (x, y) = ParsePointToken(tokens[i++]);
                    path.AppendChild(new Drawing.MoveTo(new Drawing.Point { X = x.ToString(), Y = y.ToString() }));
                    TrackMax(ref maxX, ref maxY, x, y);
                    break;
                }
                case "L":
                {
                    var (x, y) = ParsePointToken(tokens[i++]);
                    path.AppendChild(new Drawing.LineTo(new Drawing.Point { X = x.ToString(), Y = y.ToString() }));
                    TrackMax(ref maxX, ref maxY, x, y);
                    break;
                }
                case "C":
                {
                    // Cubic bezier: 3 points (control1, control2, end)
                    var (x1, y1) = ParsePointToken(tokens[i++]);
                    var (x2, y2) = ParsePointToken(tokens[i++]);
                    var (x3, y3) = ParsePointToken(tokens[i++]);
                    path.AppendChild(new Drawing.CubicBezierCurveTo(
                        new Drawing.Point { X = x1.ToString(), Y = y1.ToString() },
                        new Drawing.Point { X = x2.ToString(), Y = y2.ToString() },
                        new Drawing.Point { X = x3.ToString(), Y = y3.ToString() }
                    ));
                    TrackMax(ref maxX, ref maxY, x3, y3);
                    break;
                }
                case "Q":
                {
                    // Quadratic bezier: 2 points (control, end)
                    var (x1, y1) = ParsePointToken(tokens[i++]);
                    var (x2, y2) = ParsePointToken(tokens[i++]);
                    path.AppendChild(new Drawing.QuadraticBezierCurveTo(
                        new Drawing.Point { X = x1.ToString(), Y = y1.ToString() },
                        new Drawing.Point { X = x2.ToString(), Y = y2.ToString() }
                    ));
                    TrackMax(ref maxX, ref maxY, x2, y2);
                    break;
                }
                case "A":
                {
                    // arcTo. NodeBuilder.ReconstructCustomGeometryPath emits only
                    // "A{wR},{hR}" (the width/height radii), so the token carries a
                    // single "wR,hR" coordinate pair. stAng/swAng are not round-tripped
                    // by the emitter, so default them to 0 (a degenerate-but-valid arc
                    // that preserves the ArcTo element across Get→Add). Radii are
                    // scaled ×1000 like every other coordinate token.
                    var (wr, hr) = ParsePointToken(tokens[i++]);
                    path.AppendChild(new Drawing.ArcTo
                    {
                        WidthRadius = wr.ToString(),
                        HeightRadius = hr.ToString(),
                        StartAngle = "0",
                        SwingAngle = "0",
                    });
                    break;
                }
                case "Z":
                    path.AppendChild(new Drawing.CloseShapePath());
                    break;
                default:
                    // Skip unknown tokens
                    break;
            }
        }

        // Set path dimensions to bounding box
        if (maxX > 0) path.Width = maxX;
        if (maxY > 0) path.Height = maxY;

        return new Drawing.CustomGeometry(
            new Drawing.AdjustValueList(),
            new Drawing.ShapeGuideList(),
            new Drawing.AdjustHandleList(),
            new Drawing.ConnectionSiteList(),
            new Drawing.Rectangle { Left = "0", Top = "0", Right = "r", Bottom = "b" },
            new Drawing.PathList(path)
        );
    }

    /// <summary>
    /// Parse "x,y" coordinate token and scale ×1000 to OOXML standard 0-100000 range.
    /// Input coordinates are 0-100 relative space.
    /// </summary>
    private static (long x, long y) ParsePointToken(string token)
    {
        var parts = token.Split(',');
        if (parts.Length < 2)
            throw new ArgumentException($"Invalid coordinate '{token}'. Expected 'x,y' format (e.g. '100,200').");
        if (!long.TryParse(parts[0].Trim(), out var x))
            throw new ArgumentException($"Invalid x coordinate '{parts[0].Trim()}' in '{token}'. Expected a number.");
        if (!long.TryParse(parts[1].Trim(), out var y))
            throw new ArgumentException($"Invalid y coordinate '{parts[1].Trim()}' in '{token}'. Expected a number.");
        // Scale from user space (0-100) to OOXML standard (0-100000)
        return (x * 1000, y * 1000);
    }

    private static void TrackMax(ref long maxX, ref long maxY, long x, long y)
    {
        if (x > maxX) maxX = x;
        if (y > maxY) maxY = y;
    }

    /// <summary>
    /// Change the z-order of a shape within the ShapeTree.
    /// Values: "front" (topmost), "back" (bottommost), "forward" (+1), "backward" (-1),
    ///         or an integer for absolute position (1-based, 1 = back, N = front).
    /// </summary>
    private static void ApplyZOrder(DocumentFormat.OpenXml.Packaging.SlidePart slidePart, Shape shape, string value)
        => ApplyZOrder(slidePart, (OpenXmlElement)shape, value);

    // Generalized overload — picture/chart/table/group/connector all participate
    // in the slide shape-tree z-order. AddShape/AddPicture/AddChart/AddTable/
    // AddGroup/AddConnector all reach this so dump-emit `zorder=N` round-trips
    // for every content element type, not just typed Shape.
    private static void ApplyZOrder(DocumentFormat.OpenXml.Packaging.SlidePart slidePart, OpenXmlElement shape, string value)
    {
        // CONSISTENCY(nested-group): a shape nested inside a GroupShape has the
        // group as its DOM parent. ZOrder still applies within that local sibling
        // scope — accept ShapeTree or any GroupShape container.
        var container = shape.Parent as OpenXmlCompositeElement;
        if (container is not ShapeTree && container is not GroupShape)
            throw new InvalidOperationException("Shape is not in a ShapeTree or GroupShape");

        // Get all content elements (Shape, Picture, GraphicFrame, GroupShape, ConnectionShape)
        // that participate in z-order (skip structural elements like nvGrpSpPr, grpSpPr)
        var contentElements = container.ChildElements
            .Where(e => e is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape)
            .ToList();
        var currentIndex = contentElements.IndexOf(shape);
        if (currentIndex < 0) return;

        int targetIndex;
        switch (value.ToLowerInvariant())
        {
            case "front" or "top" or "bringtofront":
                targetIndex = contentElements.Count - 1;
                break;
            case "back" or "bottom" or "sendtoback":
                targetIndex = 0;
                break;
            case "forward" or "bringforward" or "+1":
                targetIndex = Math.Min(currentIndex + 1, contentElements.Count - 1);
                break;
            case "backward" or "sendbackward" or "-1":
                targetIndex = Math.Max(currentIndex - 1, 0);
                break;
            default:
                // Absolute position (1-based: 1 = back, N = front)
                if (int.TryParse(value, out var pos))
                    targetIndex = Math.Clamp(pos - 1, 0, contentElements.Count - 1);
                else
                    throw new ArgumentException($"Invalid z-order value: {value}. Use front/back/forward/backward or a number.");
                break;
        }

        if (targetIndex == currentIndex) return;

        // Remove shape from its current position
        shape.Remove();

        // Insert at new position
        if (targetIndex >= contentElements.Count - 1)
        {
            // Front: append after last content element (or at end of tree)
            container.AppendChild(shape);
        }
        else if (targetIndex <= 0)
        {
            // Back: insert before the first content element
            var firstContent = container.ChildElements
                .FirstOrDefault(e => e is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape);
            if (firstContent != null)
                firstContent.InsertBeforeSelf(shape);
            else
                container.AppendChild(shape);
        }
        else
        {
            // Refresh content list after removal
            var updatedContent = container.ChildElements
                .Where(e => e is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape)
                .ToList();
            if (targetIndex < updatedContent.Count)
                updatedContent[targetIndex].InsertBeforeSelf(shape);
            else
                container.AppendChild(shape);
        }
    }

    /// <summary>
    /// Apply a position/size property (x, y, width, height) to offset and extents.
    /// Returns true if the key was handled.
    /// </summary>
    private static bool TryApplyPositionSize(string key, string value, Drawing.Offset offset, Drawing.Extents extents)
    {
        // CONSISTENCY(geometry-aliases): left/top mirror x/y, matching Add
        // (Add.Shape.cs accepts left→x, top→y). The 10 PPTX Set geometry
        // switches gate on these aliases too; this canonicalizes the sink.
        key = key switch { "left" => "x", "top" => "y", _ => key };
        var emu = ParseEmu(value);
        // Unified bounds check for every EMU-valued geometry field.
        // ECMA-376 a:off uses ST_Coordinate (signed long) and a:ext uses
        // ST_PositiveCoordinate, but PowerPoint's drawing pipeline truncates
        // everything past INT32_MAX EMU (~5688 km worth of slide) — a larger
        // value silently corrupts the layout instead of round-tripping. Error
        // messages start with "Invalid" so OutputFormatter routes the
        // ArgumentException to invalid_value, not internal_error.
        if (emu > int.MaxValue)
            throw new ArgumentException($"Invalid {key} '{value}': exceeds the maximum supported shape coordinate (INT32_MAX EMU).");
        switch (key)
        {
            case "x":
                if (emu < int.MinValue)
                    throw new ArgumentException($"Invalid x '{value}': below the minimum supported shape coordinate (INT32_MIN EMU).");
                offset.X = emu; return true;
            case "y":
                if (emu < int.MinValue)
                    throw new ArgumentException($"Invalid y '{value}': below the minimum supported shape coordinate (INT32_MIN EMU).");
                offset.Y = emu; return true;
            case "width":
                if (emu < 0) throw new ArgumentException($"Invalid width '{value}': negative values are not allowed.");
                extents.Cx = emu; return true;
            case "height":
                if (emu < 0) throw new ArgumentException($"Invalid height '{value}': negative values are not allowed.");
                extents.Cy = emu; return true;
            default: return false;
        }
    }

    /// <summary>
    /// Populate an &lt;a:avLst&gt; with &lt;a:gd&gt; adjust handles from a
    /// canonical <c>adj=name:fmla,name:fmla</c> spec. Pre-existing children
    /// on the avLst are cleared first so a re-apply replaces rather than
    /// appends. Both name and fmla are pass-through strings — the OOXML
    /// schema accepts any non-empty token for @name (preset-defined,
    /// usually adj / adj1 / adj2 / …) and any well-formed formula
    /// expression for @fmla ("val N", "*/ adj1 width …", named references
    /// resolved by the preset's own definition).
    /// </summary>
    internal static void ApplyAdjustHandles(Drawing.AdjustValueList avLst, string spec,
        Drawing.ShapeTypeValues? preset = null)
    {
        avLst.RemoveAllChildren<Drawing.ShapeGuide>();
        if (string.IsNullOrWhiteSpace(spec)) return;
        // R19b BUG: a preset whose definition declares MORE THAN ONE adjust
        // guide (e.g. hexagon = adj + vf, star6 = adj + hf) corrupts the file
        // in real PowerPoint (0x80070570) if the authored avLst contains only
        // a subset of those guides — even though the OpenXML SDK validates it.
        // PowerPoint requires either an empty avLst (uses built-in defaults) or
        // the COMPLETE declared guide set. So we collect the user-supplied
        // guides keyed by canonical name, then emit the preset's full guide set
        // in declaration order, using the user formula where given and the
        // preset's default formula for the rest.
        (string Name, string Fmla)[]? declared = null;
        if (preset != null && PresetAdjustGuides.TryGetValue(PresetToken(preset.Value), out var declaredGuides))
            declared = declaredGuides;

        // R18 BUG B: a preset that declares NO adjust guides (rect, ellipse,
        // line, …) corrupts the file in real PowerPoint (0x80070570) if the
        // authored avLst contains any <a:gd> — even though the OpenXML SDK
        // validates it. Reject up front instead of writing an unopenable file.
        if (declared is { Length: 0 })
            throw new ArgumentException(
                $"Preset geometry '{PresetToken(preset!.Value)}' has no adjust handles; " +
                "remove the adj= property. PowerPoint reports files that carry adjust " +
                "values on this preset as corrupt.");

        var supplied = new Dictionary<string, string>(StringComparer.Ordinal);
        var orderSupplied = new List<string>();
        int idx = 0;
        foreach (var raw in spec.Split(',', StringSplitOptions.RemoveEmptyEntries))
        {
            var entry = raw.Trim();
            if (entry.Length == 0) continue;
            var colonIdx = entry.IndexOf(':');
            if (colonIdx <= 0 || colonIdx == entry.Length - 1)
                throw new ArgumentException(
                    $"Invalid adj spec '{entry}'. Expected 'name:formula' tokens (e.g. 'adj1:val 6000').");
            var name = entry[..colonIdx].Trim();
            var fmla = entry[(colonIdx + 1)..].Trim();
            if (name.Length == 0 || fmla.Length == 0)
                throw new ArgumentException(
                    $"Invalid adj spec '{entry}'. Both name and formula must be non-empty.");
            // Bare numeric convenience: "adj:25000" means "adj:val 25000".
            if (long.TryParse(fmla, out _)) fmla = $"val {fmla}";
            // Validate the guide-formula grammar. fmla is xs:string in the
            // schema, so our validator stays green on garbage — but real
            // PowerPoint refuses the file (0x80070570) when fmla isn't a
            // well-formed ECMA-376 guide formula.
            ValidateGuideFormula(entry, fmla);
            // R18 BUG A (generalized): PowerPoint validates each <a:gd name="…">
            // against the names the preset's own definition declares; an unknown
            // name (e.g. "adj1" on a single-handle preset whose guide is literally
            // "adj") makes real PowerPoint refuse the file (0x80070570) even
            // though the OpenXML SDK considers it schema-valid. Single-guide
            // presets: remap position 0 to the declared name (previously
            // special-cased for donut / noSmoking only). Multi-guide presets:
            // reject undeclared names outright.
            if (declared is { Length: 1 })
            {
                if (idx > 0)
                    throw new ArgumentException(
                        $"Preset geometry '{PresetToken(preset!.Value)}' declares a single adjust " +
                        $"handle ('{declared[0].Name}'); got extra handle '{name}'.");
                name = declared[0].Name;
            }
            else if (declared is { Length: > 1 } && !declared.Any(g => g.Name == name))
            {
                throw new ArgumentException(
                    $"Adjust handle '{name}' is not declared by preset geometry " +
                    $"'{PresetToken(preset!.Value)}'. Valid handles: " +
                    string.Join(", ", declared.Select(g => g.Name)) + ".");
            }
            if (supplied.TryAdd(name, fmla)) orderSupplied.Add(name);
            else supplied[name] = fmla;
            idx++;
        }

        // If this preset declares more than one guide, always emit the full
        // guide set so PowerPoint accepts the authored avLst (R19b: a subset
        // of a multi-guide preset's guides is rejected with 0x80070570).
        if (declared is { Length: > 1 })
        {
            foreach (var (name, defFmla) in declared)
            {
                var fmla = supplied.TryGetValue(name, out var userFmla) ? userFmla : defFmla;
                avLst.AppendChild(new Drawing.ShapeGuide { Name = name, Formula = fmla });
            }
            return;
        }

        // Unknown / single-guide preset: keep prior behavior — emit exactly what
        // the user supplied, in their order.
        foreach (var name in orderSupplied)
            avLst.AppendChild(new Drawing.ShapeGuide { Name = name, Formula = supplied[name] });
    }

    // ECMA-376 §20.1.9.11 guide formula: an operator followed by numeric or
    // guide-name arguments (e.g. "val 25000", "*/ w 1 2", "+- adj 0 100000").
    private static readonly HashSet<string> GuideFormulaOps = new(StringComparer.Ordinal)
    {
        "val", "*/", "+-", "+/", "?:", "abs", "at2", "cat2", "cos",
        "max", "min", "mod", "pin", "sat2", "sin", "sqrt", "tan",
    };

    private static void ValidateGuideFormula(string entry, string fmla)
    {
        var tokens = fmla.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var ok = tokens.Length >= 2 && GuideFormulaOps.Contains(tokens[0])
            && tokens.Skip(1).All(t =>
                long.TryParse(t, out _)
                || System.Text.RegularExpressions.Regex.IsMatch(t, "^[A-Za-z_][A-Za-z0-9_]*$"));
        if (!ok)
            throw new ArgumentException(
                $"Invalid adj formula in '{entry}'. Expected an ECMA-376 guide formula " +
                "such as 'val 25000' (or a bare number), got '" + fmla + "'.");
    }

    /// <summary>
    /// OOXML token for a preset geometry (e.g. <c>roundRect</c>), used as the
    /// key into <see cref="PresetAdjustGuides"/>.
    /// </summary>
    private static string PresetToken(Drawing.ShapeTypeValues preset)
        => ((IEnumValue)preset).Value;
}
