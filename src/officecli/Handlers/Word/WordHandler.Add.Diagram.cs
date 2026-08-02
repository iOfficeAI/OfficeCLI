// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core.Diagram;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // A 'diagram' is an ADD-only synthesizer (like 'equation'): it parses
    // mermaid text, lays out a graph via the shared format-agnostic engine
    // (Core/Diagram), and expands into native, editable drawing shapes +
    // connectors in the body. It is deliberately NOT a persistent element —
    // after Add it is a set of ordinary <w:drawing> shapes, so it has no
    // matching Set/Get/Query on a "diagram" node (documented exception to the
    // Add-and-Set feature checklist). The parse + layout are shared with the
    // pptx emitter; only this mapping onto docx DrawingML differs. The one
    // format-specific concern vs pptx: docx has no slide to resize, so the
    // diagram is scaled to fit the section's text-area width (never enlarged),
    // and all shapes are floating anchors positioned relative to the margin.
    private const double DiagramCmToEmu = 360000.0;

    private string AddDiagram(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        // Input mirrors `equation` / the pptx diagram: canonical `mermaid`
        // (+ aliases text/dsl) inline, or `src`/`path` to a .mmd file.
        var mermaidText = properties.GetValueOrDefault("mermaid")
                          ?? properties.GetValueOrDefault("text")
                          ?? properties.GetValueOrDefault("dsl");
        if (string.IsNullOrWhiteSpace(mermaidText)
            && (properties.TryGetValue("src", out var srcFile) || properties.TryGetValue("path", out srcFile))
            && !string.IsNullOrWhiteSpace(srcFile))
        {
            if (!System.IO.File.Exists(srcFile))
                throw new ArgumentException($"diagram source file not found: '{srcFile}'.");
            mermaidText = System.IO.File.ReadAllText(srcFile);
        }
        if (string.IsNullOrWhiteSpace(mermaidText))
            throw new ArgumentException("diagram requires inline 'mermaid' text (aliases: text, dsl) or a 'src' .mmd file path.");

        // render mode: native (built-in editable shapes) | image (real mermaid.js in
        // a headless browser → embedded PNG, covers EVERY mermaid type at full
        // fidelity) | auto (default: image when a browser is available, else native).
        var renderMode = (properties.GetValueOrDefault("render") ?? "auto").Trim().ToLowerInvariant();
        bool forceImage = renderMode is "image" or "svg" or "browser";
        if (forceImage && !MermaidImageRenderer.IsAvailable())
            throw new ArgumentException(
                "render=image needs mermaid-cli (mmdc) or a headless browser (Chrome/Chromium/Edge). "
                + "Install one, or use render=native for the built-in synthesizer.");
        bool wantImage = forceImage
            || (renderMode is not ("native" or "shapes") && MermaidImageRenderer.IsAvailable());
        if (wantImage)
            return AddDiagramAsImage(parent, parentPath, index, properties, mermaidText, allowNativeFallback: !forceImage);

        return AddDiagramNative(parent, parentPath, index, properties, mermaidText);
    }

    // Built-in synthesizer: mermaid → laid-out graph → native <w:drawing> shapes.
    private string AddDiagramNative(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties, string mermaidText)
    {
        var lo = DiagramCompiler.Compile(mermaidText);
        if (lo.Nodes.Count == 0)
            throw new ArgumentException("diagram parsed to zero nodes — check the mermaid syntax.");

        var (host, hostRoot) = ResolveDrawingHost(parent, parentPath);

        // Fit-to-box: docx has no slide to resize (no pptx-style poster), so the
        // diagram always scales into the available space. `width`/`height` give an
        // explicit box (may enlarge, mirroring picture/chart); with neither, fit the
        // section's text-area width and never enlarge (keeps small graphs readable).
        double natW = lo.SlideWidthCm, natH = lo.SlideHeightCm;
        double contentCm = SectionContentWidthCm();
        bool hasW = properties.TryGetValue("width", out var wStr);
        bool hasH = properties.TryGetValue("height", out var hStr);
        double scale;
        if (hasW || hasH)
        {
            double boxW = hasW ? ParseEmu(wStr!) / DiagramCmToEmu : contentCm;
            double boxH = hasH ? ParseEmu(hStr!) / DiagramCmToEmu : double.PositiveInfinity;
            scale = natW > 0.01 ? Math.Min(boxW / natW, boxH / natH) : 1.0;
        }
        else
        {
            scale = natW > 0.01 ? Math.Min(1.0, contentCm / natW) : 1.0;
        }
        long Emu(double cm) => (long)Math.Round(cm * scale * DiagramCmToEmu);
        // Font scales WITH the box (the layout sized every box to hold its text at
        // the base point size, so any uniform scale keeps text fitting). Floor at 1
        // only to avoid a 0pt run — a fixed higher floor (e.g. 6) forces the font
        // LARGER than the shrunken box on a heavily fit-scaled wide diagram →
        // overflow/mid-word wrap (the "text too big for the box" symptom). The node
        // bodyPr's normAutofit shrinks further if a rounding edge still overflows.
        int fontPt = Math.Max(1, (int)Math.Round(18 * lo.FontScale * scale));
        int labelPt = Math.Max(1, (int)Math.Round(10 * scale));

        // Drawings aren't in the document yet, so NextDocPropId() would return the
        // same value for each; allocate one base and increment locally so every id
        // (the group's wp:docPr + each child's wps:cNvPr) is unique.
        uint nextId = NextDocPropId();

        // Compute every child's absolute EMU box FIRST so the group bounding box —
        // and thus each child's group-relative offset — can be derived. z-order:
        // nodes behind, edges above, labels on top (children render in document
        // order inside the group).
        var nodeBoxes = lo.Nodes
            .Select(n => (n, x: Emu(n.X), y: Emu(n.Y), cx: Emu(n.W), cy: Emu(n.H)))
            .ToList();
        var edgeBoxes = new List<(RoutedEdge e, long minX, long minY, long w, long h)>();
        foreach (var e in lo.Edges)
        {
            if (e.Points.Count < 2) continue;
            long[] px = e.Points.Select(p => Emu(p.X)).ToArray();
            long[] py = e.Points.Select(p => Emu(p.Y)).ToArray();
            long mnX = px.Min(), mnY = py.Min(), w = px.Max() - mnX, h = py.Max() - mnY;
            const long pad = 12700; // 1pt — keep an axis-aligned segment non-degenerate
            if (w < pad) { mnX -= (pad - w) / 2; w = pad; }
            if (h < pad) { mnY -= (pad - h) / 2; h = pad; }
            edgeBoxes.Add((e, mnX, mnY, w, h));
        }
        var labelBoxes = lo.Labels
            .Select(lbl =>
            {
                double lw = Math.Max(1.0, DiagramLabelWidthCm(lbl.Text));
                return (lbl, x: Emu(lbl.Cx - lw / 2), y: Emu(lbl.Cy - 0.26), cx: Emu(lw), cy: Emu(0.52));
            })
            .ToList();

        // Group bounding box across every child.
        var lefts = new List<long>(); var tops = new List<long>();
        var rights = new List<long>(); var bots = new List<long>();
        foreach (var b in nodeBoxes) { lefts.Add(b.x); tops.Add(b.y); rights.Add(b.x + b.cx); bots.Add(b.y + b.cy); }
        foreach (var b in edgeBoxes) { lefts.Add(b.minX); tops.Add(b.minY); rights.Add(b.minX + b.w); bots.Add(b.minY + b.h); }
        foreach (var b in labelBoxes) { lefts.Add(b.x); tops.Add(b.y); rights.Add(b.x + b.cx); bots.Add(b.y + b.cy); }
        long gMinX = lefts.Min(), gMinY = tops.Min();
        long gW = rights.Max() - gMinX, gH = bots.Max() - gMinY;

        // Build the <wps:wsp> children in group-relative coordinates (off = abs − gMin).
        var kids = new StringBuilder();
        foreach (var b in nodeBoxes)
        {
            var (geom, fill, line) = DiagramStyles.ByShape[b.n.Shape];
            kids.Append(BuildDiagramNodeWsp(nextId++, geom, fill, line, b.n.Label, fontPt,
                b.x - gMinX, b.y - gMinY, b.cx, b.cy));
        }
        foreach (var b in edgeBoxes)
            kids.Append(BuildDiagramEdgeWsp(nextId++, b.e.Points, b.e.ArrowAtEnd, b.e.Dashed, Emu,
                b.minX, b.minY, b.w, b.h, gMinX, gMinY));
        foreach (var b in labelBoxes)
            // Opaque (flowchart) labels mask the edge line; sequence labels sit in
            // empty space → no fill, so they don't break the lifeline they cross.
            kids.Append(BuildDiagramNodeWsp(nextId++, "rect", b.lbl.Opaque ? "FFFFFF" : null, null, b.lbl.Text, labelPt,
                b.x - gMinX, b.y - gMinY, b.cx, b.cy, label: true));

        // ONE drawing: a wpg group wrapping every child, anchored at the group's
        // top-left (margin-relative). chOff/chExt == off/ext so children keep the
        // coordinates computed above; a later `set width/height` on the group
        // scales them via that baseline (mirrors the pptx group wrapper), so the
        // whole diagram stays adjustable as a unit after Add.
        string groupXml = BuildDiagramGroupDrawing(nextId++, gMinX, gMinY, gW, gH, kids.ToString());
        var para = new Paragraph();
        para.AppendChild(new Run(ParseDrawingFromXml(groupXml)));
        AssignParaId(para);
        InsertAtIndexOrAppend(host, para, index);

        // The whole diagram is a single grouped drawing → one group anchor
        // (/body/group[N]); `set width/height` on it resizes the diagram as a unit.
        int groupIdx = CountGroupsInHost(host, para);
        return $"{hostRoot}/group[{groupIdx}]";
    }

    // High-fidelity path: render with the real mermaid.js (headless browser) to PNG
    // and embed it as a picture, stamping the source into alt-text so the diagram
    // travels in the file and is regenerable. In auto mode any render failure falls
    // back to the native synthesizer.
    private string AddDiagramAsImage(OpenXmlElement parent, string parentPath, int? index,
                                     Dictionary<string, string> properties, string mermaidText, bool allowNativeFallback)
    {
        // Bake theme/layout/look into the source as frontmatter so they render and
        // round-trip via alt-text; native fallback keeps the ORIGINAL source (its
        // parser has no frontmatter/elk support). Mirrors the pptx image path.
        var composedText = MermaidImageRenderer.ComposeSource(mermaidText,
            properties.GetValueOrDefault("theme"),
            properties.GetValueOrDefault("layout"),
            properties.GetValueOrDefault("look"));
        var background = properties.GetValueOrDefault("background");

        string imgPath;
        try { imgPath = MermaidImageRenderer.RenderToPngFile(composedText, background); }
        // A syntax error is bad input — surface it (with mermaid's line-numbered
        // message) so the caller can fix the source. Never fall back to native: the
        // synthesizer would reject the same broken text or, worse, draw garbage.
        catch (MermaidSyntaxException) { throw; }
        catch when (allowNativeFallback) { return AddDiagramNative(parent, parentPath, index, properties, mermaidText); }
        try
        {
            var pic = new Dictionary<string, string>(properties);
            foreach (var k in new[] { "mermaid", "text", "dsl", "src", "path", "render", "poster",
                                      "theme", "layout", "look", "background" })
                pic.Remove(k);
            pic["src"] = imgPath;
            if (!(pic.TryGetValue("alt", out var a) && !string.IsNullOrEmpty(a)))
                pic["alt"] = MermaidImageRenderer.SourceTag + composedText;

            // poster: grow the PAGE to the whole diagram instead of shrinking a long
            // flowchart into an unreadable sliver. The pptx path grows the slide; the
            // Word analogue grows the section's page size. Word caps a page edge at
            // 22in (55.88cm) — far below pptx's 142cm — so a very long chart is still
            // limited (splitting it stays the better answer past that), but a
            // moderately long one becomes readable on one tall page. The page size is
            // per SECTION: this sets the diagram's section, so on an export-a-diagram
            // document (the diagram is the content) the whole page becomes the poster;
            // mixed content in the same section shares the grown page (the explicit
            // opt-in's tradeoff, mirroring pptx growing the one slide).
            // poster resolution mirrors the pptx path: explicit poster=true always
            // grows the page; poster=false always fits the page; UNSET is the
            // ADAPTIVE DEFAULT — grow the page only when fitting the diagram to the
            // page would shrink it below the readability floor. Auto-poster stands
            // down when the caller pinned an explicit width/height.
            {
                bool posterSet = properties.ContainsKey("poster");
                bool posterOn = OfficeCli.Core.ParseHelpers.IsTruthy(properties.GetValueOrDefault("poster"));
                bool hasExplicitBox = pic.ContainsKey("width") || pic.ContainsKey("height");
                using (var s = System.IO.File.OpenRead(imgPath))
                {
                    var dims = OfficeCli.Core.ImageSource.TryGetDimensions(s);
                    if (dims is { Width: > 0, Height: > 0 } d)
                    {
                        bool grow = posterOn
                            || (!posterSet && !hasExplicitBox
                                && MermaidImageRenderer.ExceedsOnePageReadably(
                                    d.Width, d.Height, SectionContentWidthCm(), SectionContentHeightCm()));
                        if (grow)
                        {
                            const double maxPageEdgeCm = 55.88; // Word's 22in page limit
                            const double marginCm = 0.5;
                            double wCm = d.Width / 96.0 * 2.54, hCm = d.Height / 96.0 * 2.54;
                            double clamp = Math.Min(1.0, (maxPageEdgeCm - 2 * marginCm) / Math.Max(wCm, hCm));
                            wCm *= clamp; hCm *= clamp;
                            SetLastSectionPageCm(wCm + 2 * marginCm, hCm + 2 * marginCm, marginCm);
                            pic["width"] = wCm.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + "cm";
                            pic.Remove("height"); // aspect preserved
                            return AddPicture(parent, parentPath, index, pic);
                        }
                    }
                }
            }
            // Sizing parity with the native path AND the pptx image path: the diagram
            // is ALWAYS scaled to FIT its box with aspect preserved (never stretched).
            // The box is the caller's width/height when given, else the section's
            // content BOX (text-area width AND available page height, so a tall
            // flowchart stays on one page). We emit WIDTH only and let AddPicture
            // derive the height from the aspect ratio — passing both width and height
            // straight through would squash a portrait diagram into a wide box.
            {
                double boxWCm = pic.TryGetValue("width", out var wOverride)
                    ? ParseEmu(wOverride) / DiagramCmToEmu : SectionContentWidthCm();
                double boxHCm = pic.TryGetValue("height", out var hOverride)
                    ? ParseEmu(hOverride) / DiagramCmToEmu : SectionContentHeightCm();
                double outWCm = boxWCm; // fallback: dims unreadable → plain width-fit
                using (var s = System.IO.File.OpenRead(imgPath))
                {
                    var dims = OfficeCli.Core.ImageSource.TryGetDimensions(s);
                    if (dims is { Width: > 0, Height: > 0 } d)
                    {
                        double fit = Math.Min(boxWCm / d.Width, boxHCm / d.Height);
                        outWCm = d.Width * fit;
                    }
                }
                pic["width"] = outWCm.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + "cm";
                pic.Remove("height"); // aspect preserved: AddPicture derives height from width
            }
            return AddPicture(parent, parentPath, index, pic);
        }
        finally { try { System.IO.File.Delete(imgPath); } catch { /* best effort */ } }
    }

    private const string DiagramNs =
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
        "xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" " +
        "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"";

    // Build the <wpg:wgp> group drawing that wraps every child. chOff/chExt ==
    // off/ext → children keep the absolute-within-group coordinates they were
    // built with; a later `set width/height` shrinks ext while chExt stays the
    // baseline, so Word scales the children (same model as the pptx group).
    private static string BuildDiagramGroupDrawing(uint groupId, long posX, long posY, long cx, long cy, string childrenXml)
    {
        return
            $"<w:drawing {DiagramNs}><wp:anchor distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" simplePos=\"0\" " +
            "relativeHeight=\"2510000\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">" +
            "<wp:simplePos x=\"0\" y=\"0\"/>" +
            $"<wp:positionH relativeFrom=\"margin\"><wp:posOffset>{posX}</wp:posOffset></wp:positionH>" +
            $"<wp:positionV relativeFrom=\"margin\"><wp:posOffset>{posY}</wp:posOffset></wp:positionV>" +
            $"<wp:extent cx=\"{cx}\" cy=\"{cy}\"/><wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/><wp:wrapNone/>" +
            $"<wp:docPr id=\"{groupId}\" name=\"Diagram {groupId}\"/>" +
            "<wp:cNvGraphicFramePr/>" +
            "<a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\">" +
            "<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr>" +
            $"<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>" +
            "</wpg:grpSpPr>" +
            childrenXml +
            "</wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing>";
    }

    // A node/label as a <wps:wsp> child of the group. off/ext are in the group's
    // child coordinate space (== absolute-within-group EMU here). No wp:anchor —
    // the enclosing group owns placement.
    private static string BuildDiagramNodeWsp(uint id, string preset, string? fill, string? line,
                                              string text, int fontPt, long x, long y, long cx, long cy,
                                              bool label = false)
    {
        string fillXml = string.IsNullOrEmpty(fill)
            ? "<a:noFill/>"
            : $"<a:solidFill><a:srgbClr val=\"{fill}\"/></a:solidFill>";
        string lnXml = string.IsNullOrEmpty(line)
            ? "<a:ln><a:noFill/></a:ln>"
            : $"<a:ln w=\"9525\"><a:solidFill><a:srgbClr val=\"{line}\"/></a:solidFill></a:ln>";
        int szHalfPt = fontPt * 2;
        // rFonts with an eastAsia slot so Word resolves CJK glyphs. Without it a
        // textbox run inherits the Latin default (Calibri) and East-Asian text can
        // render blank; PowerPoint auto-applies the theme's CJK font, Word doesn't.
        const string rFonts = "<w:rFonts w:eastAsia=\"SimSun\" w:hint=\"eastAsia\"/>";
        string txbx =
            "<wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr>" +
            $"<w:r><w:rPr>{rFonts}<w:sz w:val=\"{szHalfPt}\"/><w:szCs w:val=\"{szHalfPt}\"/></w:rPr>" +
            $"<w:t xml:space=\"preserve\">{SecurityElement.Escape(text)}</w:t></w:r></w:p></w:txbxContent></wps:txbx>";
        // Zero the text insets. Word's default insets (~0.25cm L/R, ~0.13cm T/B)
        // are fixed EMU, not scaled — on a fit-shrunk node box (~1cm wide) they
        // eat over half the width, forcing the text to wrap and clip (looks like
        // "font too big for the box"). The layout already bakes visual padding
        // into the box size. Labels: single-line no-wrap; nodes: wrap + normAutofit.
        string bodyPr = label
            ? "<wps:bodyPr rot=\"0\" wrap=\"none\" lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\" anchor=\"ctr\" anchorCtr=\"1\"><a:noAutofit/></wps:bodyPr>"
            : "<wps:bodyPr rot=\"0\" lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\" anchor=\"ctr\" anchorCtr=\"0\"><a:normAutofit/></wps:bodyPr>";
        string nm = label ? "DiagramLabel" : "DiagramShape";
        return
            $"<wps:wsp><wps:cNvPr id=\"{id}\" name=\"{nm} {id}\"/><wps:cNvSpPr/><wps:spPr>" +
            $"<a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>" +
            $"<a:prstGeom prst=\"{preset}\"><a:avLst/></a:prstGeom>{fillXml}{lnXml}</wps:spPr>" +
            $"{txbx}{bodyPr}</wps:wsp>";
    }

    // An edge as a <wps:wsp> child: a custGeom polyline whose path is relative to
    // its own box, positioned at (absMinX−groupMinX, absMinY−groupMinY) in the
    // group's child coordinate space.
    private static string BuildDiagramEdgeWsp(uint id, IReadOnlyList<Pt> points,
                                              bool arrowAtEnd, bool dashed, Func<double, long> emu,
                                              long absMinX, long absMinY, long w, long h,
                                              long groupMinX, long groupMinY)
    {
        var path = new StringBuilder();
        for (int i = 0; i < points.Count; i++)
        {
            long x = emu(points[i].X) - absMinX, y = emu(points[i].Y) - absMinY;
            path.Append(i == 0
                ? $"<a:moveTo><a:pt x=\"{x}\" y=\"{y}\"/></a:moveTo>"
                : $"<a:lnTo><a:pt x=\"{x}\" y=\"{y}\"/></a:lnTo>");
        }
        string dash = dashed ? "<a:prstDash val=\"dash\"/>" : "";
        string arrow = arrowAtEnd ? "<a:tailEnd type=\"triangle\"/>" : "";
        string ln = $"<a:ln w=\"12700\" cap=\"flat\"><a:solidFill><a:srgbClr val=\"{DiagramStyles.EdgeColor}\"/></a:solidFill>{dash}<a:round/>{arrow}</a:ln>";
        string custGeom =
            $"<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l=\"0\" t=\"0\" r=\"{w}\" b=\"{h}\"/>" +
            $"<a:pathLst><a:path w=\"{w}\" h=\"{h}\">{path}</a:path></a:pathLst></a:custGeom>";
        return
            $"<wps:wsp><wps:cNvPr id=\"{id}\" name=\"DiagramEdge {id}\"/><wps:cNvSpPr/><wps:spPr>" +
            $"<a:xfrm><a:off x=\"{absMinX - groupMinX}\" y=\"{absMinY - groupMinY}\"/><a:ext cx=\"{w}\" cy=\"{h}\"/></a:xfrm>" +
            $"{custGeom}<a:noFill/>{ln}</wps:spPr><wps:bodyPr/></wps:wsp>";
    }

    private static double DiagramLabelWidthCm(string text)
    {
        double w = 0;
        foreach (var c in text) w += c > 0x2E80 ? 0.58 : 0.30;
        return Math.Min(w, 5.0) + 0.4;
    }

    // Text-area width (page width − left/right margins) of the last section, in
    // cm. Falls back to US-Letter with 1in margins (~16.51cm) when unset.
    private double SectionContentWidthCm()
    {
        try
        {
            var body = _doc?.MainDocumentPart?.Document?.Body;
            var sect = body?.Elements<SectionProperties>().LastOrDefault()
                       ?? body?.Descendants<SectionProperties>().LastOrDefault();
            var pgSz = sect?.Elements<PageSize>().FirstOrDefault();
            var pgMar = sect?.Elements<PageMargin>().FirstOrDefault();
            long wTw = pgSz?.Width?.Value ?? 12240u;
            long lTw = pgMar?.Left?.Value ?? 1440u;
            long rTw = pgMar?.Right?.Value ?? 1440u;
            double contentCm = (wTw - lTw - rTw) / 1440.0 * 2.54;
            return contentCm > 1.0 ? contentCm : 16.51;
        }
        catch { return 16.51; }
    }

    // Text-area height (page height − top/bottom margins) of the last section, in
    // cm. Falls back to US-Letter with 1in margins (~24.13cm) when unset. Used to
    // cap a fit-to-page diagram image so a tall graph stays on one page.
    // Set the last section's page size and uniform margins (poster). Creates the
    // sectPr / pgSz / pgMar if absent. Twips = cm * 1440 / 2.54.
    private void SetLastSectionPageCm(double pageWCm, double pageHCm, double marginCm)
    {
        var body = _doc?.MainDocumentPart?.Document?.Body;
        if (body == null) return;
        var sect = body.Elements<SectionProperties>().LastOrDefault();
        if (sect == null) { sect = new SectionProperties(); body.Append(sect); }
        uint W(double cm) => (uint)Math.Round(cm * 1440.0 / 2.54);

        var pgSz = sect.Elements<PageSize>().FirstOrDefault();
        if (pgSz == null) { pgSz = new PageSize(); sect.InsertAt(pgSz, 0); }
        pgSz.Width = W(pageWCm);
        pgSz.Height = W(pageHCm);

        var pgMar = sect.Elements<PageMargin>().FirstOrDefault();
        if (pgMar == null) { pgMar = new PageMargin(); sect.InsertAfter(pgMar, pgSz); }
        pgMar.Top = (int)W(marginCm);
        pgMar.Bottom = (int)W(marginCm);
        pgMar.Left = (uint)W(marginCm);
        pgMar.Right = (uint)W(marginCm);
    }

    private double SectionContentHeightCm()
    {
        try
        {
            var body = _doc?.MainDocumentPart?.Document?.Body;
            var sect = body?.Elements<SectionProperties>().LastOrDefault()
                       ?? body?.Descendants<SectionProperties>().LastOrDefault();
            var pgSz = sect?.Elements<PageSize>().FirstOrDefault();
            var pgMar = sect?.Elements<PageMargin>().FirstOrDefault();
            long hTw = pgSz?.Height?.Value ?? 15840u;
            long tTw = (long)(pgMar?.Top?.Value ?? 1440);
            long bTw = (long)(pgMar?.Bottom?.Value ?? 1440);
            double contentCm = (hTw - tTw - bTw) / 1440.0 * 2.54;
            return contentCm > 1.0 ? contentCm : 24.13;
        }
        catch { return 24.13; }
    }
}
