// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Speaker Notes helpers ====================

    private static string GetNotesText(NotesSlidePart notesPart)
    {
        var spTree = notesPart.NotesSlide?.CommonSlideData?.ShapeTree;
        if (spTree == null) return "";

        foreach (var shape in spTree.Elements<Shape>())
        {
            var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>();
            if (ph?.Index?.Value == 1) // body/notes placeholder
            {
                return string.Join("\n", shape.TextBody?.Elements<Drawing.Paragraph>()
                    .Select(p => string.Concat(p.Elements<Drawing.Run>().Select(r => r.Text?.Text ?? "")))
                    ?? Enumerable.Empty<string>());
            }
        }
        return "";
    }

    /// <summary>
    /// Walk the notes body placeholder's first run and mirror its run-level
    /// formatting onto the notes DocumentNode.Format dictionary. Required so
    /// that `Set /slide[N]/notes --prop bold=true color=#FF0000` is observable
    /// via `Get /slide[N]/notes` (round-trip parity with shape Get).
    /// </summary>
    private static void PopulateNotesFormat(NotesSlidePart notesPart, DocumentNode node)
    {
        var spTree = notesPart.NotesSlide?.CommonSlideData?.ShapeTree;
        if (spTree == null) return;
        Shape? notesShape = null;
        foreach (var shape in spTree.Elements<Shape>())
        {
            var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>();
            if (ph?.Index?.Value == 1) { notesShape = shape; break; }
        }
        if (notesShape == null) return;
        var firstRun = notesShape.TextBody?
            .Elements<Drawing.Paragraph>()
            .SelectMany(p => p.Elements<Drawing.Run>())
            .FirstOrDefault();
        if (firstRun == null) return;
        // Reuse RunToNode (Format-builder for runs) and merge its Format keys
        // into the notes node — skip `text` so we don't clobber the notes-level
        // text (concatenation of all runs across all paragraphs).
        var runNode = RunToNode(firstRun, node.Path + "/run[1]", notesPart);
        foreach (var kv in runNode.Format)
        {
            if (kv.Key == "text") continue;
            node.Format[kv.Key] = kv.Value;
        }
        // Reading direction (rtl) lives on the paragraph, not the run.
        var firstPara = notesShape.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
        if (firstPara?.ParagraphProperties?.RightToLeft?.Value == true)
            node.Format["direction"] = "rtl";
    }

    private static void SetNotesText(NotesSlidePart notesPart, string text)
    {
        var spTree = notesPart.NotesSlide?.CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Notes slide has no shape tree");

        // Find body placeholder (idx=1)
        Shape? notesShape = null;
        foreach (var shape in spTree.Elements<Shape>())
        {
            var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>();
            if (ph?.Index?.Value == 1)
            {
                notesShape = shape;
                break;
            }
        }

        if (notesShape == null)
        {
            notesShape = new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = 3, Name = "Notes Placeholder 2" },
                    new NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties(
                        new PlaceholderShape { Type = PlaceholderValues.Body, Index = 1 }
                    )
                ),
                new ShapeProperties(),
                new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle(), new Drawing.Paragraph())
            );
            spTree.AppendChild(notesShape);
        }

        var textBody = notesShape.TextBody
            ?? (notesShape.TextBody = new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle()));

        textBody.RemoveAllChildren<Drawing.Paragraph>();
        foreach (var line in text.Split('\n'))
        {
            textBody.AppendChild(new Drawing.Paragraph(
                new Drawing.Run(
                    new Drawing.RunProperties { Language = "en-US" },
                    new Drawing.Text { Text = line }
                )
            ));
        }

        notesPart.NotesSlide!.Save();
    }

    /// <summary>
    /// Apply reading direction (rtl/ltr) to the notes body shape on a notes
    /// slide. Mirrors the shape direction fix in PowerPointHandler.Add.Shape.cs:
    /// sets &lt;a:pPr rtl="1"/&gt; on every paragraph and rtlCol="1" on the
    /// shape's bodyPr. RTL notes are required for Arabic / Hebrew authors
    /// reviewing speaker notes.
    /// </summary>
    private static void ApplyNotesDirection(NotesSlidePart notesPart, string value)
    {
        var spTree = notesPart.NotesSlide?.CommonSlideData?.ShapeTree;
        if (spTree == null) return;
        Shape? notesShape = null;
        foreach (var shape in spTree.Elements<Shape>())
        {
            var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>();
            if (ph?.Index?.Value == 1)
            {
                notesShape = shape;
                break;
            }
        }
        if (notesShape == null) return;
        bool rtl = ParsePptDirectionRtl(value);
        foreach (var para in notesShape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
        {
            var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
            // Clear semantics: direction=ltr strips the rtl attribute.
            if (rtl) pProps.RightToLeft = true;
            else pProps.RightToLeft = null;
        }
        var bodyPr = notesShape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
        if (bodyPr != null)
        {
            if (rtl)
                bodyPr.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("", "rtlCol", "", "1"));
            else
                bodyPr.RemoveAttribute("rtlCol", "");
        }
    }

    // ==================== Notes master ====================

    // The default 7.5" x 10" notes page. The placeholder rectangles in
    // BuildNotesMasterXml are PowerPoint's own numbers for this page size;
    // NotesMasterRect scales them for decks that override <p:notesSz>.
    private const long StdNotesCx = 6858000;
    private const long StdNotesCy = 9144000;

    /// <summary>
    /// Scale one of PowerPoint's stock notes-master rectangles to the deck's
    /// actual &lt;p:notesSz&gt;, passing the stock numbers through unchanged on
    /// the default notes page.
    /// </summary>
    private static (long X, long Y, long Cx, long Cy) NotesMasterRect(
        long pageCx, long pageCy, long x, long y, long cx, long cy)
    {
        long sx(long v) => pageCx == StdNotesCx ? v : (long)Math.Round(v * (double)pageCx / StdNotesCx);
        long sy(long v) => pageCy == StdNotesCy ? v : (long)Math.Round(v * (double)pageCy / StdNotesCy);
        return (sx(x), sy(y), sx(cx), sy(cy));
    }

    /// <summary>
    /// Create the presentation's notes master if it has none, wire its theme
    /// relationship, and register it in &lt;p:notesMasterIdLst&gt;.
    ///
    /// BUG(notes-orphan): notes slides emit empty &lt;p:spPr/&gt; on their
    /// placeholders and inherit position, size and text style from the notes
    /// master. Before this, `add /slide[N] --type notes` wrote notesSlide1.xml
    /// with no master to inherit from and no notesSlide rels part at all, so the
    /// notes text had no geometry anywhere in the package. Text extraction still
    /// worked, which is why it went unnoticed; conformant readers (python-pptx)
    /// throw on the missing notesMaster relationship. The equivalent creation
    /// code already existed, but only on the raw-set /notesMaster dump-replay
    /// path in PowerPointHandler.cs — which the add/set notes verbs never reach.
    /// </summary>
    private static NotesMasterPart EnsureNotesMasterPart(PresentationPart presentationPart)
    {
        if (presentationPart.NotesMasterPart != null) return presentationPart.NotesMasterPart;

        var notesSz = presentationPart.Presentation?.NotesSize;
        long w = notesSz?.Cx?.Value ?? SlideSizeDefaults.NotesPortraitCx;
        long h = notesSz?.Cy?.Value ?? SlideSizeDefaults.NotesPortraitCy;
        var sldSz = presentationPart.Presentation?.SlideSize;
        long slideCx = sldSz?.Cx?.Value ?? SlideSizeDefaults.Widescreen16x9Cx;
        long slideCy = sldSz?.Cy?.Value ?? SlideSizeDefaults.Widescreen16x9Cy;

        var nmPart = presentationPart.AddNewPart<NotesMasterPart>();
        nmPart.NotesMaster = new NotesMaster(BuildNotesMasterXml(w, h, slideCx, slideCy));

        // A notes master must carry a theme relationship; the +mn-lt / tx1
        // references in its notesStyle resolve through it. Give it its OWN
        // theme part rather than sharing the presentation's — GrowSlideMasterParts
        // clones for grown slide masters for the same reason, PptxBatchEmitter
        // emits `add-part theme /notesMaster` whenever the notes master has a
        // theme at all, and PowerPoint itself writes a separate theme2.xml.
        // Sharing made every dump->replay round-trip grow a duplicate theme.
        // Fall back to a slide master's theme for decks where theme1.xml hangs
        // off the master rather than the presentation.
        var themeSource = presentationPart.ThemePart?.Theme
            ?? presentationPart.SlideMasterParts.FirstOrDefault()?.ThemePart?.Theme;
        var nmTheme = nmPart.AddNewPart<ThemePart>();
        nmTheme.Theme = themeSource != null
            ? (Drawing.Theme)themeSource.CloneNode(true)
            : new Drawing.Theme();
        nmTheme.Theme.Save();

        // AddNewPart only wires the relationship. Without a <p:notesMasterId>
        // entry the part is an orphan reference and PowerPoint refuses the
        // file. An existing but empty or stale <p:notesMasterIdLst> leaves it
        // just as orphaned as a missing one, so test the entries, not the list.
        // Schema order (CT_Presentation): sldMasterIdLst -> notesMasterIdLst
        // -> handoutMasterIdLst -> sldIdLst.
        var pres = presentationPart.Presentation;
        if (pres != null)
        {
            var nmIdLst = pres.NotesMasterIdList;
            if (nmIdLst == null)
            {
                nmIdLst = new NotesMasterIdList();
                if (pres.SlideMasterIdList != null)
                    pres.InsertAfter(nmIdLst, pres.SlideMasterIdList);
                else
                    pres.PrependChild(nmIdLst);
            }
            bool declared = nmIdLst.Elements<NotesMasterId>().Any(e =>
                !string.IsNullOrEmpty(e.Id?.Value)
                && presentationPart.Parts.Any(p => p.RelationshipId == e.Id!.Value));
            if (!declared)
            {
                nmIdLst.RemoveAllChildren<NotesMasterId>();
                nmIdLst.AppendChild(new NotesMasterId { Id = presentationPart.GetIdOfPart(nmPart) });
            }
            pres.Save();
        }

        nmPart.NotesMaster.Save();
        return nmPart;
    }

    /// <summary>
    /// Build notesMaster1.xml. Written as XML rather than an SDK object graph
    /// because it is six near-identical placeholders plus nine near-identical
    /// notesStyle levels; the object-graph form is ~200 lines of constructor
    /// soup with the geometry buried in it.
    /// </summary>
    private static string BuildNotesMasterXml(long pageCx, long pageCy, long slideCx, long slideCy)
    {
        var hdr = NotesMasterRect(pageCx, pageCy, 0, 0, 2971800, 458788);
        var dt = NotesMasterRect(pageCx, pageCy, 3884613, 0, 2971800, 458788);
        var ftr = NotesMasterRect(pageCx, pageCy, 0, 8685213, 2971800, 458787);
        var num = NotesMasterRect(pageCx, pageCy, 3884613, 8685213, 2971800, 458787);

        // The slide-image placeholder is a thumbnail of the slide, so its height
        // follows the deck's aspect ratio rather than a fixed number — PowerPoint
        // recomputes it the same way (4572000-wide box: 2571750 high on 16:9,
        // 3429000 on 4:3). Getting this wrong letterboxes every notes page.
        var imgBox = NotesMasterRect(pageCx, pageCy, 1143000, 685800, 4572000, 0);
        long imgCy = slideCx > 0 ? (long)Math.Round(imgBox.Cx * (double)slideCy / slideCx) : imgBox.Cy;
        var img = (imgBox.X, imgBox.Y, imgBox.Cx, Cy: imgCy);

        // Body sits below the thumbnail, snapped to the next quarter-inch grid
        // line strictly below it. Reproduces PowerPoint's 3429000 on 16:9 and
        // 4343400 on 4:3 decks.
        const long QuarterInch = 228600;
        var bodyBox = NotesMasterRect(pageCx, pageCy, 685800, 0, 5486400, 4114800);
        // Never let a tall thumbnail push the notes box into the footer band.
        // Clamping the height alone is not enough: on a wide/short notes page
        // the snapped origin can already sit below the footer, which would make
        // the height term negative.
        long bodyY = Math.Min(
            (imgBox.Y + imgCy) / QuarterInch * QuarterInch + QuarterInch,
            Math.Max(imgBox.Y, ftr.Y - QuarterInch));
        long bodyCy = Math.Max(QuarterInch, Math.Min(bodyBox.Cy, ftr.Y - bodyY));
        var body = (bodyBox.X, Y: bodyY, bodyBox.Cx, Cy: bodyCy);

        static string Xfrm((long X, long Y, long Cx, long Cy) r) =>
            $"<a:xfrm><a:off x=\"{r.X}\" y=\"{r.Y}\"/><a:ext cx=\"{r.Cx}\" cy=\"{r.Cy}\"/></a:xfrm>"
            + "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>";

        // Header/footer/date/slide-number text bodies: 12pt, anchored to the
        // outer edge of the notes page like PowerPoint's.
        static string Banner(string algn, string? field) =>
            "<p:txBody><a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\"/>"
            + $"<a:lstStyle><a:lvl1pPr algn=\"{algn}\"><a:defRPr sz=\"1200\"/></a:lvl1pPr></a:lstStyle>"
            + "<a:p>"
            + (field ?? "")
            + "<a:endParaRPr lang=\"en-US\"/></a:p></p:txBody>";

        // No cached <a:t>: the reader computes the value, so we never bake a
        // stale date or page number into the file.
        const string DateField =
            "<a:fld id=\"{B1E4B0A2-5C8E-4E5F-9C1D-2F3A4B5C6D71}\" type=\"datetimeFigureOut\">"
            + "<a:rPr lang=\"en-US\"/></a:fld>";
        const string SlideNumField =
            "<a:fld id=\"{B1E4B0A2-5C8E-4E5F-9C1D-2F3A4B5C6D72}\" type=\"slidenum\">"
            + "<a:rPr lang=\"en-US\"/></a:fld>";

        var notesStyle = new System.Text.StringBuilder("<p:notesStyle>");
        for (int lvl = 1; lvl <= 9; lvl++)
        {
            notesStyle.Append(
                $"<a:lvl{lvl}pPr marL=\"{(lvl - 1) * 457200}\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" "
                + "eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\">"
                + "<a:defRPr sz=\"1200\" kern=\"1200\">"
                + "<a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill>"
                + "<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>"
                + $"</a:defRPr></a:lvl{lvl}pPr>");
        }
        notesStyle.Append("</p:notesStyle>");

        return "<p:notesMaster xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
            + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
            + "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
            + "<p:cSld><p:spTree>"
            + "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
            + "<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>"
            + "<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>"

            + "<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Header Placeholder 1\"/>"
            + "<p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>"
            + "<p:nvPr><p:ph type=\"hdr\" sz=\"quarter\"/></p:nvPr></p:nvSpPr>"
            + $"<p:spPr>{Xfrm(hdr)}</p:spPr>{Banner("l", null)}</p:sp>"

            + "<p:sp><p:nvSpPr><p:cNvPr id=\"3\" name=\"Date Placeholder 2\"/>"
            + "<p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>"
            + "<p:nvPr><p:ph type=\"dt\" idx=\"1\"/></p:nvPr></p:nvSpPr>"
            + $"<p:spPr>{Xfrm(dt)}</p:spPr>{Banner("r", DateField)}</p:sp>"

            + "<p:sp><p:nvSpPr><p:cNvPr id=\"4\" name=\"Slide Image Placeholder 3\"/>"
            + "<p:cNvSpPr><a:spLocks noGrp=\"1\" noRot=\"1\" noChangeAspect=\"1\"/></p:cNvSpPr>"
            + "<p:nvPr><p:ph type=\"sldImg\" idx=\"2\"/></p:nvPr></p:nvSpPr>"
            + $"<p:spPr>{Xfrm(img)}<a:noFill/>"
            + "<a:ln w=\"12700\"><a:solidFill><a:prstClr val=\"black\"/></a:solidFill></a:ln></p:spPr>"
            + "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp>"

            + "<p:sp><p:nvSpPr><p:cNvPr id=\"5\" name=\"Notes Placeholder 4\"/>"
            + "<p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>"
            + "<p:nvPr><p:ph type=\"body\" sz=\"quarter\" idx=\"3\"/></p:nvPr></p:nvSpPr>"
            + $"<p:spPr>{Xfrm(body)}</p:spPr>"
            + "<p:txBody><a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\"/>"
            + "<a:lstStyle/><a:p><a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp>"

            + "<p:sp><p:nvSpPr><p:cNvPr id=\"6\" name=\"Footer Placeholder 5\"/>"
            + "<p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>"
            + "<p:nvPr><p:ph type=\"ftr\" sz=\"quarter\" idx=\"4\"/></p:nvPr></p:nvSpPr>"
            + $"<p:spPr>{Xfrm(ftr)}</p:spPr>{Banner("l", null)}</p:sp>"

            + "<p:sp><p:nvSpPr><p:cNvPr id=\"7\" name=\"Slide Number Placeholder 6\"/>"
            + "<p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>"
            + "<p:nvPr><p:ph type=\"sldNum\" sz=\"quarter\" idx=\"5\"/></p:nvPr></p:nvSpPr>"
            + $"<p:spPr>{Xfrm(num)}</p:spPr>{Banner("r", SlideNumField)}</p:sp>"

            + "</p:spTree></p:cSld>"
            + "<p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" "
            + "accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" "
            + "accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>"
            + notesStyle
            + "</p:notesMaster>";
    }

    private static NotesSlidePart EnsureNotesSlidePart(SlidePart slidePart)
    {
        if (slidePart.NotesSlidePart != null) return slidePart.NotesSlidePart;

        var notesPart = slidePart.AddNewPart<NotesSlidePart>();
        notesPart.NotesSlide = new NotesSlide(
            new CommonSlideData(
                new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties { Id = 1, Name = "" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()
                    ),
                    new GroupShapeProperties(new Drawing.TransformGroup()),
                    // Slide image placeholder (idx=0)
                    new Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties { Id = 2, Name = "Slide Image Placeholder 1" },
                            new NonVisualShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties(
                                new PlaceholderShape { Type = PlaceholderValues.SlideImage, Index = 0 }
                            )
                        ),
                        new ShapeProperties(),
                        new TextBody(new Drawing.BodyProperties(), new Drawing.ListStyle(), new Drawing.Paragraph())
                    ),
                    // Notes body placeholder (idx=1)
                    new Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties { Id = 3, Name = "Notes Placeholder 2" },
                            new NonVisualShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties(
                                new PlaceholderShape { Type = PlaceholderValues.Body, Index = 1 }
                            )
                        ),
                        new ShapeProperties(),
                        new TextBody(
                            new Drawing.BodyProperties(),
                            new Drawing.ListStyle(),
                            new Drawing.Paragraph()
                        )
                    )
                )
            ),
            // Notes slides inherit the notes master's color map verbatim.
            new ColorMapOverride(new Drawing.MasterColorMapping())
        );

        // The notes slide's placeholders carry an empty <p:spPr/> and inherit
        // geometry and text style from the notes master, so the master must
        // exist and be related from this part. Without it the notes text has no
        // position or size anywhere in the package. PowerPoint also links the
        // notes slide back to its slide; mirror that.
        var presentationPart = (slidePart.OpenXmlPackage as PresentationDocument)?.PresentationPart;
        if (presentationPart != null)
            notesPart.AddPart(EnsureNotesMasterPart(presentationPart));
        notesPart.AddPart(slidePart);

        notesPart.NotesSlide.Save();
        return notesPart;
    }
}
