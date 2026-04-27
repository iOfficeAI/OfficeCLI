// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;
using SpreadsheetDrawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeCli.Handlers;

// Per-element-type Add helpers for drawing/anchor paths (ole, picture, shape, slicer, sparkline). Mechanically extracted from the Add() god-method.
public partial class ExcelHandler
{
    private string AddOle(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        // ---- Excel OLE insertion (modern form, Office 2010+) ----
        //
        // Structure produced:
        //   Worksheet > oleObjects > oleObject(progId, shapeId, r:id=embedRel)
        //     > objectPr(defaultSize=0, r:id=iconRel)
        //       > anchor(moveWithCells=1)
        //         > from(col, colOff, row, rowOff)
        //         > to  (col, colOff, row, rowOff)
        //
        // We skip the legacy VML shape that Excel historically
        // generates as a fallback — when the modern objectPr/anchor
        // is present, Office 2010+ renders from it directly. The
        // constraint-required shapeId still needs a value, so we
        // allocate one in the legal range (1-67098623) unique per
        // worksheet. For round-trip fidelity, we also create an
        // empty legacy VmlDrawingPart and register the shapeId
        // there so the relationship target exists.
        var oleSheetSegs = parentPath.TrimStart('/').Split('/', 2);
        var oleSheetName = oleSheetSegs[0];
        var oleWorksheet = FindWorksheet(oleSheetName)
            ?? throw new ArgumentException($"Sheet not found: {oleSheetName}");

        var oleSrc = OfficeCli.Core.OleHelper.RequireSource(properties);
        OfficeCli.Core.OleHelper.WarnOnUnknownOleProps(properties);

        // CONSISTENCY(excel-ole-display): Excel OLE does not have a
        // DrawAspect concept — worksheet objects are always shown as
        // icons via objectPr/anchor, so 'display' would be a no-op.
        // Set already rejects it; Add must too, for symmetry.
        if (properties.ContainsKey("display"))
            throw new ArgumentException(
                "'display' property is not supported for Excel OLE "
                + "(Excel always shows objects as icon). Remove --prop display.");

        // CONSISTENCY(ole-name): Word/PPT OLE accept --prop name=... and
        // round-trip it via Get. SpreadsheetML x:oleObject has no Name
        // attribute in the schema, so there is nowhere to persist it.
        // Throw explicitly rather than silently dropping the value —
        // keep 'name' in KnownOleProps so Word/PPT still accept it.
        if (properties.ContainsKey("name"))
            throw new ArgumentException(
                "'name' property is not supported for Excel OLE "
                + "(Spreadsheet OleObject schema has no Name attribute). Remove --prop name.");

        // 1. Embedded payload.
        var (oleEmbedRelId, _) = OfficeCli.Core.OleHelper.AddEmbeddedPart(oleWorksheet, oleSrc, _filePath);

        // 2. Icon preview image part.
        var (_, oleIconRelId) = OfficeCli.Core.OleHelper.CreateIconPart(oleWorksheet, properties);

        // 3. Resolve ProgID.
        var oleProgId = OfficeCli.Core.OleHelper.ResolveProgId(properties, oleSrc);

        // 4. Anchor: accept either cell range "B2:E6" or x/y/width/height (column units).
        // CONSISTENCY(ole-width-units): sub-cell precision is carried in
        // ColumnOffset/RowOffset (EMU) so unit-qualified widths like
        // "6cm" survive a round-trip. When the user passes a cell range
        // or a bare integer cell count, the remainder offsets are 0 and
        // behavior matches the legacy whole-cell path.
        int oleFromCol, oleFromRow, oleToCol, oleToRow;
        // FromMarker offsets are always zero (anchor starts at cell boundary);
        // ToMarker offsets carry the sub-cell EMU remainder for unit-qualified
        // width/height inputs, preserving round-trip precision.
        const long oleFromColOff = 0, oleFromRowOff = 0;
        long oleToColOff = 0, oleToRowOff = 0;
        if (properties.TryGetValue("anchor", out var oleAnchorStr) && !string.IsNullOrWhiteSpace(oleAnchorStr))
        {
            // CONSISTENCY(ole-width-units): anchor= defines the full
            // rectangle (start+end cells), so width/height on the same
            // Add call would be ambiguous and are silently dropped.
            // Warn loudly rather than fail, so existing scripts keep
            // working but users notice the dropped value.
            if (properties.ContainsKey("width") || properties.ContainsKey("height"))
                Console.Error.WriteLine(
                    "Warning: 'width'/'height' are ignored when 'anchor' is provided (anchor defines the full rectangle).");
            var m = Regex.Match(oleAnchorStr, @"^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$", RegexOptions.IgnoreCase);
            if (!m.Success)
                throw new ArgumentException($"Invalid anchor: '{oleAnchorStr}'. Expected e.g. 'B2' or 'B2:E6'.");
            // CONSISTENCY(xdr-coords): XDR ColumnId/RowId are 0-based;
            // ColumnNameToIndex returns 1-based, so subtract 1 here.
            oleFromCol = ColumnNameToIndex(m.Groups[1].Value) - 1;
            oleFromRow = int.Parse(m.Groups[2].Value) - 1;
            if (m.Groups[3].Success)
            {
                oleToCol = ColumnNameToIndex(m.Groups[3].Value) - 1;
                oleToRow = int.Parse(m.Groups[4].Value) - 1;
            }
            else
            {
                oleToCol = oleFromCol + 2;
                oleToRow = oleFromRow + 3;
            }
        }
        else
        {
            var (ax, ay, awEmu, ahEmu) = ParseAnchorBoundsEmu(properties, "1", "1", "3", "4");
            oleFromCol = ax;
            oleFromRow = ay;
            // Split the EMU extent into (whole cells, sub-cell offset).
            // EmuPerCol/Row constants live in ExcelHandler.Helpers.cs.
            long wholeCols = awEmu / EmuPerColApprox;
            long remCols = awEmu % EmuPerColApprox;
            long wholeRows = ahEmu / EmuPerRowApprox;
            long remRows = ahEmu % EmuPerRowApprox;
            oleToCol = ax + (int)wholeCols;
            oleToRow = ay + (int)wholeRows;
            oleToColOff = remCols;
            oleToRowOff = remRows;
        }

        // 5. Ensure the legacy VmlDrawingPart exists and carry an
        //    empty shape placeholder referencing our shapeId. This
        //    keeps the schema happy without writing VML rendering
        //    logic — Excel 2010+ renders from objectPr/anchor anyway.
        var oleVmlPart = oleWorksheet.VmlDrawingParts.FirstOrDefault()
            ?? oleWorksheet.AddNewPart<VmlDrawingPart>();
        // Allocate a unique shapeId per worksheet (1025+N is the
        // conventional Excel starting point for legacy VML shapes).
        var existingOleCount = GetSheet(oleWorksheet).Descendants<OleObject>().Count();
        uint oleShapeId = (uint)(1025 + existingOleCount);
        EnsureExcelVmlShapeForOle(oleVmlPart, oleShapeId, oleFromCol, oleFromRow, oleToCol, oleToRow);

        // Ensure worksheet references the VML drawing part.
        var oleWsElement = GetSheet(oleWorksheet);
        if (oleWsElement.GetFirstChild<LegacyDrawing>() == null)
        {
            var vmlRelId = oleWorksheet.GetIdOfPart(oleVmlPart);
            // LegacyDrawing must sit after the AutoFilter/Phonetic
            // region per schema order — safe to insert before the
            // last known printing-related elements. Use InsertAfter
            // relative to AutoFilter when present, else append.
            var lgd = new LegacyDrawing { Id = vmlRelId };
            var pageSetup = oleWsElement.GetFirstChild<PageSetup>();
            if (pageSetup != null)
                oleWsElement.InsertAfter(lgd, pageSetup);
            else
                oleWsElement.AppendChild(lgd);
        }

        // 6. Build the oleObject element + objectPr/anchor.
        var oleObj = new OleObject
        {
            ProgId = oleProgId,
            ShapeId = oleShapeId,
            Id = oleEmbedRelId,
        };
        var objectPr = new EmbeddedObjectProperties
        {
            DefaultSize = false,
            Id = oleIconRelId,
        };
        var anchor = new ObjectAnchor { MoveWithCells = true };
        anchor.AppendChild(new FromMarker(
            new XDR.ColumnId(oleFromCol.ToString()),
            new XDR.ColumnOffset(oleFromColOff.ToString()),
            new XDR.RowId(oleFromRow.ToString()),
            new XDR.RowOffset(oleFromRowOff.ToString())));
        anchor.AppendChild(new ToMarker(
            new XDR.ColumnId(oleToCol.ToString()),
            new XDR.ColumnOffset(oleToColOff.ToString()),
            new XDR.RowId(oleToRow.ToString()),
            new XDR.RowOffset(oleToRowOff.ToString())));
        objectPr.AppendChild(anchor);
        oleObj.AppendChild(objectPr);

        // 7. Find/create oleObjects collection and append.
        var oleObjects = oleWsElement.GetFirstChild<OleObjects>();
        if (oleObjects == null)
        {
            oleObjects = new OleObjects();
            // Schema: oleObjects sits between picture and controls;
            // safest is after tableParts if present, else before
            // pageSetup, else append.
            var insertBefore = oleWsElement.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.ExtensionList>()
                ?? (OpenXmlElement?)null;
            if (insertBefore != null)
                oleWsElement.InsertBefore(oleObjects, insertBefore);
            else
                oleWsElement.AppendChild(oleObjects);
        }
        oleObjects.AppendChild(oleObj);

        SaveWorksheet(oleWorksheet);

        var oleCount = oleWsElement.Descendants<OleObject>().Count();
        return $"/{oleSheetName}/ole[{oleCount}]";
    }

    private string AddPicture(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        var picSegments = parentPath.TrimStart('/').Split('/', 2);
        var picSheetName = picSegments[0];
        var picWorksheet = FindWorksheet(picSheetName)
            ?? throw new ArgumentException($"Sheet not found: {picSheetName}");

        if (!properties.TryGetValue("path", out var imgPath)
            && !properties.TryGetValue("src", out imgPath))
            throw new ArgumentException("'src' property is required for picture type");

        // CONSISTENCY(picture-emu): use ParseAnchorBoundsEmu like OLE,
        // so width/height accept unit-qualified strings ("6cm", "2in")
        // in addition to bare integer cell counts.
        var (px, py, pwEmu, phEmu) = ParseAnchorBoundsEmu(properties, "0", "0", "5", "5");
        // P9: accept `altText=` as alias for `alt=`.
        var alt = properties.GetValueOrDefault("alt")
            ?? properties.GetValueOrDefault("altText")
            ?? properties.GetValueOrDefault("alttext", "");

        var picDrawingsPart = picWorksheet.DrawingsPart
            ?? picWorksheet.AddNewPart<DrawingsPart>();

        if (picDrawingsPart.WorksheetDrawing == null)
        {
            picDrawingsPart.WorksheetDrawing = new XDR.WorksheetDrawing();
            picDrawingsPart.WorksheetDrawing.Save();

            if (GetSheet(picWorksheet).GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>() == null)
            {
                var drawingRelId = picWorksheet.GetIdOfPart(picDrawingsPart);
                GetSheet(picWorksheet).Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = drawingRelId });
                SaveWorksheet(picWorksheet);
            }
        }

        var (xlImgStream, imgPartType) = OfficeCli.Core.ImageSource.Resolve(imgPath);
        using var xlImgDispose = xlImgStream;

        // CONSISTENCY(svg-dual-rep): same dual-representation as Word
        // and PPT — main r:embed points to a PNG fallback, SVG is
        // referenced via a:blip/a:extLst asvg:svgBlip.
        string imgRelId;
        string? xlSvgRelId = null;
        if (imgPartType == ImagePartType.Svg)
        {
            var svgPart = picDrawingsPart.AddImagePart(ImagePartType.Svg);
            svgPart.FeedData(xlImgStream);
            xlSvgRelId = picDrawingsPart.GetIdOfPart(svgPart);

            if (properties.TryGetValue("fallback", out var xlFallback) && !string.IsNullOrWhiteSpace(xlFallback))
            {
                var (fbRaw, fbType) = OfficeCli.Core.ImageSource.Resolve(xlFallback);
                using var fbDispose = fbRaw;
                var fbPart = picDrawingsPart.AddImagePart(fbType);
                fbPart.FeedData(fbRaw);
                imgRelId = picDrawingsPart.GetIdOfPart(fbPart);
            }
            else
            {
                var pngPart = picDrawingsPart.AddImagePart(ImagePartType.Png);
                pngPart.FeedData(new MemoryStream(
                    OfficeCli.Core.SvgImageHelper.TransparentPng1x1, writable: false));
                imgRelId = picDrawingsPart.GetIdOfPart(pngPart);
            }
        }
        else
        {
            var imgPart = picDrawingsPart.AddImagePart(imgPartType);
            imgPart.FeedData(xlImgStream);
            imgRelId = picDrawingsPart.GetIdOfPart(imgPart);
        }

        var picId = picDrawingsPart.WorksheetDrawing.Descendants<XDR.NonVisualDrawingProperties>()
            .Select(p => (uint?)p.Id?.Value ?? 0u).DefaultIfEmpty(0u).Max() + 1;
        // CONSISTENCY(picture-emu): split EMU extent into whole-cell
        // count + sub-cell offset, matching the OLE anchor path.
        long picWholeCols = pwEmu / EmuPerColApprox;
        long picRemCols = pwEmu % EmuPerColApprox;
        long picWholeRows = phEmu / EmuPerRowApprox;
        long picRemRows = phEmu % EmuPerRowApprox;

        // DEFERRED(xlsx/picture-anchor-mode) P12: honor `anchorMode=`
        // oneCell|absolute|twoCell. Default remains twoCell for back-compat.
        // oneCell → <xdr:oneCellAnchor> with from + ext; picture auto-scales
        //           if the column/row containing "from" is resized.
        // absolute → <xdr:absoluteAnchor> with pos (x/y EMU) + ext; picture
        //            does not move or resize with cells.
        // twoCell  → <xdr:twoCellAnchor> with from + to markers (default).
        //
        // CONSISTENCY(ole-width-units): `anchor=B2:E6` (cell-range) is
        // parsed here the same way as the OLE and shape branches; it
        // implies anchorMode=twoCell. `anchor=oneCell|twoCell|absolute`
        // is still honored as the mode for back-compat. Explicit
        // `anchorMode=` always wins. When both `anchor=<range>` and
        // `x/y/width/height` are supplied, anchor wins with a warning
        // (same convention as the shape/OLE branches).
        var picAnchorRaw = properties.GetValueOrDefault("anchor");
        var picAnchorModeExplicit = properties.GetValueOrDefault("anchorMode");
        bool picHasRange = false;
        int picRangeFromCol = 0, picRangeFromRow = 0, picRangeToCol = -1, picRangeToRow = -1;
        // `anchor=` is either a cell-range ("B2" / "B2:E6") or an
        // anchorMode token ("oneCell"/"twoCell"/"absolute"). Prefer the
        // cell-range interpretation; fall back to mode-token only when
        // the value is a recognized token. Explicit `anchorMode=` wins
        // the mode selection regardless.
        if (!string.IsNullOrWhiteSpace(picAnchorRaw) && !IsAnchorModeToken(picAnchorRaw))
        {
            if (!TryParseCellRangeAnchor(picAnchorRaw, out picRangeFromCol, out picRangeFromRow, out picRangeToCol, out picRangeToRow))
                throw new ArgumentException($"Invalid anchor: '{picAnchorRaw}'. Expected e.g. 'B2', 'B2:E6', or one of 'oneCell'/'twoCell'/'absolute'.");
            picHasRange = true;
            if (properties.ContainsKey("width") || properties.ContainsKey("height")
                || properties.ContainsKey("x") || properties.ContainsKey("y"))
                Console.Error.WriteLine(
                    "Warning: 'x'/'y'/'width'/'height' are ignored when 'anchor' is a cell range (anchor defines the full rectangle).");
        }
        var picAnchorMode = (picAnchorModeExplicit
            ?? (picHasRange ? "twoCell" : picAnchorRaw)
            ?? "twoCell").Trim().ToLowerInvariant();

        var picShape = BuildPictureElementWithTransform(picId, alt ?? "", imgRelId, xlSvgRelId, properties);

        // For oneCell / absolute anchors the size is carried by an <xdr:ext>
        // element instead of a To marker, so we must also stamp the extent
        // onto the picture's Transform2D so rotation / flip metadata plus
        // the rendered size stay in sync.
        if (picAnchorMode is "onecell" or "absolute")
        {
            var picXfrm = picShape.Descendants<Drawing.Transform2D>().FirstOrDefault();
            if (picXfrm != null)
            {
                var ext2d = picXfrm.Extents ?? new Drawing.Extents();
                ext2d.Cx = pwEmu;
                ext2d.Cy = phEmu;
                picXfrm.Extents = ext2d;
            }
        }

        OpenXmlElement anchor;
        switch (picAnchorMode)
        {
            case "onecell":
            {
                int oneFromCol = picHasRange ? picRangeFromCol : px;
                int oneFromRow = picHasRange ? picRangeFromRow : py;
                var oneAnchor = new XDR.OneCellAnchor(
                    new XDR.FromMarker(
                        new XDR.ColumnId(oneFromCol.ToString()),
                        new XDR.ColumnOffset("0"),
                        new XDR.RowId(oneFromRow.ToString()),
                        new XDR.RowOffset("0")
                    ),
                    new XDR.Extent { Cx = pwEmu, Cy = phEmu },
                    picShape,
                    new XDR.ClientData()
                );
                anchor = oneAnchor;
                break;
            }
            case "absolute":
            {
                // Absolute anchor pos: accept `x=`/`y=` in the same unit
                // syntax as width/height (bare EMU, or "1in", "2cm").
                long absX = 0, absY = 0;
                if (properties.TryGetValue("x", out var absXs))
                    absX = OfficeCli.Core.EmuConverter.ParseEmu(absXs);
                if (properties.TryGetValue("y", out var absYs))
                    absY = OfficeCli.Core.EmuConverter.ParseEmu(absYs);
                var absAnchor = new XDR.AbsoluteAnchor(
                    new XDR.Position { X = absX, Y = absY },
                    new XDR.Extent { Cx = pwEmu, Cy = phEmu },
                    picShape,
                    new XDR.ClientData()
                );
                anchor = absAnchor;
                break;
            }
            default:
            {
                int twoFromCol, twoFromRow, twoToCol, twoToRow;
                long twoToColOff, twoToRowOff;
                if (picHasRange)
                {
                    twoFromCol = picRangeFromCol;
                    twoFromRow = picRangeFromRow;
                    if (picRangeToCol >= 0)
                    {
                        twoToCol = picRangeToCol;
                        twoToRow = picRangeToRow;
                        twoToColOff = 0;
                        twoToRowOff = 0;
                    }
                    else
                    {
                        // Single-cell range in twoCell mode: fall back to width/height extent.
                        twoToCol = twoFromCol + (int)picWholeCols;
                        twoToRow = twoFromRow + (int)picWholeRows;
                        twoToColOff = picRemCols;
                        twoToRowOff = picRemRows;
                    }
                }
                else
                {
                    twoFromCol = px;
                    twoFromRow = py;
                    twoToCol = px + (int)picWholeCols;
                    twoToRow = py + (int)picWholeRows;
                    twoToColOff = picRemCols;
                    twoToRowOff = picRemRows;
                }
                anchor = new XDR.TwoCellAnchor(
                    new XDR.FromMarker(
                        new XDR.ColumnId(twoFromCol.ToString()),
                        new XDR.ColumnOffset("0"),
                        new XDR.RowId(twoFromRow.ToString()),
                        new XDR.RowOffset("0")
                    ),
                    new XDR.ToMarker(
                        new XDR.ColumnId(twoToCol.ToString()),
                        new XDR.ColumnOffset(twoToColOff.ToString()),
                        new XDR.RowId(twoToRow.ToString()),
                        new XDR.RowOffset(twoToRowOff.ToString())
                    ),
                    picShape,
                    new XDR.ClientData()
                );
                break;
            }
        }

        picDrawingsPart.WorksheetDrawing.AppendChild(anchor);

        // P10: picture decorative=true — emit <a:extLst><a:ext uri="...">
        // <a16:decorative val="1"/></a:ext></a:extLst> under <xdr:cNvPr>.
        // Requires declaring xmlns:a16 on the drawing root; mirrors the
        // sparkline pattern of adding namespaces idempotently.
        if (properties.TryGetValue("decorative", out var picDec) && IsTruthy(picDec))
        {
            var picCNvPrDec = anchor.Descendants<XDR.NonVisualDrawingProperties>().FirstOrDefault();
            if (picCNvPrDec != null)
            {
                const string a16Ns = "http://schemas.microsoft.com/office/drawing/2014/main";
                var wsDrawingRoot = picDrawingsPart.WorksheetDrawing;
                if (wsDrawingRoot.LookupNamespace("a16") == null)
                    wsDrawingRoot.AddNamespaceDeclaration("a16", a16Ns);
                var decInner = new OpenXmlUnknownElement("a16", "decorative", a16Ns);
                decInner.SetAttribute(new OpenXmlAttribute("", "val", "", "1"));
                var ext = new Drawing.Extension { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };
                ext.Append(decInner);
                var extLst = picCNvPrDec.GetFirstChild<Drawing.ExtensionList>()
                    ?? picCNvPrDec.AppendChild(new Drawing.ExtensionList());
                extLst.Append(ext);
            }
        }

        // P8: picture-level hyperlink — <a:hlinkClick> under <xdr:cNvPr>.
        // External URL → add rel on DrawingsPart, reference its rId.
        // Internal (starts with '#') → no rel, use Location attribute.
        // CONSISTENCY(xlsx-hyperlink): mirrors cell link handling in
        // commit 60e1455.
        var picHlink = properties.GetValueOrDefault("hyperlink")
            ?? properties.GetValueOrDefault("link");
        if (!string.IsNullOrWhiteSpace(picHlink))
        {
            var picCNvPr = anchor.Descendants<XDR.NonVisualDrawingProperties>().FirstOrDefault();
            if (picCNvPr != null)
            {
                Drawing.HyperlinkOnClick hlClick;
                if (picHlink.StartsWith("#"))
                {
                    // No rel, no @r:id — pure in-document jump via @location.
                    hlClick = new Drawing.HyperlinkOnClick { Id = "" };
                    hlClick.SetAttribute(new OpenXmlAttribute(
                        "", "location", "", picHlink.Substring(1)));
                }
                else
                {
                    var hlUri = new Uri(picHlink, UriKind.RelativeOrAbsolute);
                    var hlRel = picDrawingsPart.AddHyperlinkRelationship(hlUri, isExternal: true);
                    hlClick = new Drawing.HyperlinkOnClick { Id = hlRel.Id };
                }
                picCNvPr.AppendChild(hlClick);
            }
        }

        picDrawingsPart.WorksheetDrawing.Save();

        // DEFERRED(xlsx/picture-anchor-mode) P12: enumerate all anchor
        // kinds (twoCell / oneCell / absolute) when counting picture slots.
        var picAnchors = picDrawingsPart.WorksheetDrawing
            .Elements<OpenXmlElement>()
            .Where(a => (a is XDR.TwoCellAnchor || a is XDR.OneCellAnchor || a is XDR.AbsoluteAnchor)
                && a.Descendants<XDR.Picture>().Any())
            .ToList();
        var picIdx = picAnchors.IndexOf(anchor) + 1;

        return $"/{picSheetName}/picture[{picIdx}]";
    }

    private string AddShape(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        var shpSegments = parentPath.TrimStart('/').Split('/', 2);
        var shpSheetName = shpSegments[0];
        var shpWorksheet = FindWorksheet(shpSheetName)
            ?? throw new ArgumentException($"Sheet not found: {shpSheetName}");

        // CONSISTENCY(ole-width-units): accept `anchor=B2:F7` as a cell
        // range (same grammar as OLE's anchor=), alongside the legacy
        // x/y/width/height (column/row units) form. When both are
        // supplied, warn and let anchor= win — it defines the full
        // rectangle, so width/height are ambiguous.
        // CONSISTENCY(ref-alias): `ref=<cell>` maps to single-cell
        // anchor `<cell>:<cell>`, matching cell/comment/table which
        // accept `ref=` as the placement address. Explicit `anchor=`
        // wins if both are given.
        if (!properties.ContainsKey("anchor")
            && properties.TryGetValue("ref", out var shpRefProp)
            && !string.IsNullOrWhiteSpace(shpRefProp))
        {
            var refTrim = shpRefProp.Trim();
            if (!refTrim.Contains(':'))
            {
                // Single-cell ref (e.g. "B2"): expand to a 1x1 cell
                // rectangle (B2:C3) so the shape has a visible extent.
                // Using identical from/to markers produces a
                // zero-width/height invisible shape in Excel.
                if (TryParseCellRangeAnchor(refTrim, out var rc, out var rr, out _, out _))
                    refTrim = $"{refTrim}:{IndexToColumnName(rc + 2)}{rr + 2}";
                else
                    refTrim = $"{refTrim}:{refTrim}";
            }
            properties["anchor"] = refTrim;
        }
        int sx, sy, sw, sh;
        if (properties.TryGetValue("anchor", out var shpAnchorStr) && !string.IsNullOrWhiteSpace(shpAnchorStr))
        {
            if (properties.ContainsKey("width") || properties.ContainsKey("height")
                || properties.ContainsKey("x") || properties.ContainsKey("y"))
                Console.Error.WriteLine(
                    "Warning: 'x'/'y'/'width'/'height' are ignored when 'anchor' is provided (anchor defines the full rectangle).");
            if (!TryParseCellRangeAnchor(shpAnchorStr, out var sxFrom, out var syFrom, out var sxTo, out var syTo))
                throw new ArgumentException($"Invalid anchor: '{shpAnchorStr}'. Expected e.g. 'B2' or 'B2:F7'.");
            sx = sxFrom;
            sy = syFrom;
            if (sxTo < 0) { sxTo = sx + 4; syTo = sy + 2; }
            sw = sxTo - sx;
            sh = syTo - sy;
        }
        else
        {
            (sx, sy, sw, sh) = ParseAnchorBounds(properties, "1", "1", "5", "3");
        }
        var shpText = properties.GetValueOrDefault("text", "") ?? "";
        var shpName = properties.GetValueOrDefault("name", "");

        var shpDrawingsPart = shpWorksheet.DrawingsPart
            ?? shpWorksheet.AddNewPart<DrawingsPart>();

        if (shpDrawingsPart.WorksheetDrawing == null)
        {
            shpDrawingsPart.WorksheetDrawing = new XDR.WorksheetDrawing();
            shpDrawingsPart.WorksheetDrawing.Save();

            if (GetSheet(shpWorksheet).GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>() == null)
            {
                var drawingRelId = shpWorksheet.GetIdOfPart(shpDrawingsPart);
                GetSheet(shpWorksheet).Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = drawingRelId });
                SaveWorksheet(shpWorksheet);
            }
        }

        var shpId = shpDrawingsPart.WorksheetDrawing.Descendants<XDR.NonVisualDrawingProperties>()
            .Select(p => (uint?)p.Id?.Value ?? 0u).DefaultIfEmpty(0u).Max() + 1;
        if (string.IsNullOrEmpty(shpName)) shpName = $"Shape {shpId}";

        // CONSISTENCY(shape-preset): map `preset=` to a:prstGeom prst value
        // using the same token set PowerPointHandler.ParsePresetShape accepts.
        // textbox ignores preset (always "rect"). Default for shape: "rect".
        var shpPreset = Drawing.ShapeTypeValues.Rectangle;
        if (string.Equals(type, "shape", StringComparison.OrdinalIgnoreCase))
        {
            // CONSISTENCY(shape-preset-aliases): preset is canonical;
            // accept geometry/shape as aliases (mirrors Set).
            var rawPreset = properties.GetValueOrDefault("preset")
                ?? properties.GetValueOrDefault("geometry")
                ?? properties.GetValueOrDefault("shape");
            if (!string.IsNullOrWhiteSpace(rawPreset))
                shpPreset = ParseExcelShapePreset(rawPreset);
        }

        // Build ShapeProperties
        var shpXfrm = new Drawing.Transform2D(
            new Drawing.Offset { X = 0, Y = 0 },
            new Drawing.Extents { Cx = 0, Cy = 0 }
        );
        ApplyTransform2DRotationFlip(shpXfrm, properties);
        var spPr = new XDR.ShapeProperties(
            shpXfrm,
            new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = shpPreset }
        );

        // Fill — single-color `fill=` OR gradient `gradientFill=C1-C2[-C3][:angle]`.
        // SH6/shape-gradient-fill: keep `fill=` strictly single-color; gradient has its own prop
        // to avoid ambiguity (FF0000-0000FF could otherwise collide with single ARGB literals).
        if (properties.TryGetValue("gradientFill", out var shpGradFill)
            && !string.IsNullOrWhiteSpace(shpGradFill))
        {
            spPr.AppendChild(BuildShapeGradientFill(shpGradFill));
        }
        else if (properties.TryGetValue("fill", out var shpFill))
        {
            if (shpFill.Equals("none", StringComparison.OrdinalIgnoreCase))
                spPr.AppendChild(new Drawing.NoFill());
            else
            {
                var (rgb, alpha) = ParseHelpers.SanitizeColorForOoxml(shpFill);
                var solidFill = new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = rgb });
                spPr.AppendChild(solidFill);
            }
        }

        // Line/border
        if (properties.TryGetValue("line", out var shpLine))
        {
            if (shpLine.Equals("none", StringComparison.OrdinalIgnoreCase))
                spPr.AppendChild(new Drawing.Outline(new Drawing.NoFill()));
            else
            {
                var (lRgb, _) = ParseHelpers.SanitizeColorForOoxml(shpLine);
                spPr.AppendChild(new Drawing.Outline(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = lRgb })));
            }
        }

        // Effects (shadow, glow, reflection, softEdge) — shape-level only for shapes with fill
        // For fill=none shapes, shadow/glow go to text-level (rPr) below.
        // CT_EffectList schema order: blur → fillOverlay → glow → innerShdw → outerShdw → prstShdw → reflection → softEdge
        // Build each effect into a typed slot, then AppendChild in schema order below.
        var isNoFillShape = properties.TryGetValue("fill", out var fillCheck) && fillCheck.Equals("none", StringComparison.OrdinalIgnoreCase);
        Drawing.Glow? shpGlowEl = null;
        Drawing.OuterShadow? shpShadowEl = null;
        Drawing.Reflection? shpReflEl = null;
        Drawing.SoftEdge? shpSoftEl = null;
        if (!isNoFillShape)
        {
            if (properties.TryGetValue("shadow", out var shpShadow) && !shpShadow.Equals("none", StringComparison.OrdinalIgnoreCase))
            {
                var normalizedShadow = shpShadow.Replace(':', '-');
                if (IsValidBooleanString(normalizedShadow) && IsTruthy(normalizedShadow)) normalizedShadow = "000000";
                shpShadowEl = OfficeCli.Core.DrawingEffectsHelper.BuildOuterShadow(normalizedShadow, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor);
            }
            if (properties.TryGetValue("glow", out var shpGlow) && !shpGlow.Equals("none", StringComparison.OrdinalIgnoreCase))
            {
                var normalizedGlow = shpGlow.Replace(':', '-');
                if (IsValidBooleanString(normalizedGlow) && IsTruthy(normalizedGlow)) normalizedGlow = "4472C4";
                shpGlowEl = OfficeCli.Core.DrawingEffectsHelper.BuildGlow(normalizedGlow, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor);
            }
        }
        if (properties.TryGetValue("reflection", out var shpRefl) && !shpRefl.Equals("none", StringComparison.OrdinalIgnoreCase))
            shpReflEl = OfficeCli.Core.DrawingEffectsHelper.BuildReflection(shpRefl);
        if (properties.TryGetValue("softedge", out var shpSoft) && !shpSoft.Equals("none", StringComparison.OrdinalIgnoreCase))
            shpSoftEl = OfficeCli.Core.DrawingEffectsHelper.BuildSoftEdge(shpSoft);
        if (shpGlowEl != null || shpShadowEl != null || shpReflEl != null || shpSoftEl != null)
        {
            // CONSISTENCY(effect-list-schema-order): glow → outerShdw → reflection → softEdge
            var shpEffectList = new Drawing.EffectList();
            if (shpGlowEl != null) shpEffectList.AppendChild(shpGlowEl);
            if (shpShadowEl != null) shpEffectList.AppendChild(shpShadowEl);
            if (shpReflEl != null) shpEffectList.AppendChild(shpReflEl);
            if (shpSoftEl != null) shpEffectList.AppendChild(shpSoftEl);
            spPr.AppendChild(shpEffectList);
        }

        // Build TextBody with runs
        var bodyPr = new Drawing.BodyProperties { Anchor = Drawing.TextAnchoringTypeValues.Center };
        if (properties.TryGetValue("valign", out var shpValign))
        {
            // CONSISTENCY(shape-valign): mirror Set vocabulary so Add path
            // doesn't drop a known prop that round-trips through Get.
            bodyPr.Anchor = shpValign.ToLowerInvariant() switch
            {
                "top" or "t" => Drawing.TextAnchoringTypeValues.Top,
                "center" or "ctr" or "middle" or "m" or "c" => Drawing.TextAnchoringTypeValues.Center,
                "bottom" or "b" => Drawing.TextAnchoringTypeValues.Bottom,
                _ => throw new ArgumentException($"Invalid valign value: '{shpValign}'. Valid: top, center, bottom.")
            };
        }
        if (properties.TryGetValue("margin", out var shpMargin))
        {
            // CONSISTENCY(spacing-units): mirror Set — accept unit-qualified
            // input and 4-CSV round-trip from Get.
            var (lE, tE, rE, bE) = ParseShapeMarginToEmu(shpMargin);
            bodyPr.LeftInset = lE;
            bodyPr.TopInset = tE;
            bodyPr.RightInset = rE;
            bodyPr.BottomInset = bE;
        }
        var txBody = new XDR.TextBody(bodyPr, new Drawing.ListStyle());

        var lines = shpText.Replace("\\n", "\n").Split('\n');
        foreach (var line in lines)
        {
            var rPr = new Drawing.RunProperties { Language = "en-US" };

            // R2-3: accept both bare (`size`, `bold`, `color`, `font`) and `font.*`
            // sub-prop forms (`font.size`, `font.bold`, `font.color`, `font.name`,
            // `font.italic`, `font.underline`) for consistency with cell/comment.
            // Schema order: attributes → solidFill → effectLst → latin/ea
            string? rawSize = properties.GetValueOrDefault("size")
                ?? properties.GetValueOrDefault("font.size");
            if (rawSize != null)
                rPr.FontSize = (int)Math.Round(ParseHelpers.ParseFontSize(rawSize) * 100);

            string? rawBold = properties.GetValueOrDefault("bold")
                ?? properties.GetValueOrDefault("font.bold");
            if (rawBold != null && IsTruthy(rawBold))
                rPr.Bold = true;

            string? rawItalic = properties.GetValueOrDefault("italic")
                ?? properties.GetValueOrDefault("font.italic");
            if (rawItalic != null && IsTruthy(rawItalic))
                rPr.Italic = true;

            if (properties.TryGetValue("font.underline", out var shpUnder)
                || properties.TryGetValue("underline", out shpUnder))
            {
                var uv = shpUnder.ToLowerInvariant();
                rPr.Underline = uv switch
                {
                    "true" or "single" or "sng" => Drawing.TextUnderlineValues.Single,
                    "double" or "dbl" => Drawing.TextUnderlineValues.Double,
                    "none" or "false" => Drawing.TextUnderlineValues.None,
                    _ => Drawing.TextUnderlineValues.Single
                };
            }

            // Fill (color) before fonts
            string? rawColor = properties.GetValueOrDefault("color")
                ?? properties.GetValueOrDefault("font.color");
            if (rawColor != null)
            {
                var (cRgb, _) = ParseHelpers.SanitizeColorForOoxml(rawColor);
                rPr.AppendChild(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = cRgb }));
            }

            // Text-level effects for fill=none shapes
            var isNoFill = properties.TryGetValue("fill", out var f) && f.Equals("none", StringComparison.OrdinalIgnoreCase);
            if (isNoFill)
            {
                // CONSISTENCY(effect-list-schema-order): glow → outerShdw per CT_EffectList
                Drawing.Glow? txtGlowEl = null;
                Drawing.OuterShadow? txtShadowEl = null;
                if (properties.TryGetValue("shadow", out var ts) && !ts.Equals("none", StringComparison.OrdinalIgnoreCase))
                {
                    var normalizedTs = ts.Replace(':', '-');
                    if (IsValidBooleanString(normalizedTs) && IsTruthy(normalizedTs)) normalizedTs = "000000";
                    txtShadowEl = OfficeCli.Core.DrawingEffectsHelper.BuildOuterShadow(normalizedTs, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor);
                }
                if (properties.TryGetValue("glow", out var tg) && !tg.Equals("none", StringComparison.OrdinalIgnoreCase))
                {
                    var normalizedTg = tg.Replace(':', '-');
                    if (IsValidBooleanString(normalizedTg) && IsTruthy(normalizedTg)) normalizedTg = "4472C4";
                    txtGlowEl = OfficeCli.Core.DrawingEffectsHelper.BuildGlow(normalizedTg, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor);
                }
                if (txtGlowEl != null || txtShadowEl != null)
                {
                    var txtEffects = new Drawing.EffectList();
                    if (txtGlowEl != null) txtEffects.AppendChild(txtGlowEl);
                    if (txtShadowEl != null) txtEffects.AppendChild(txtShadowEl);
                    rPr.AppendChild(txtEffects);
                }
            }

            // Fonts last (schema order). Accept `font=Arial` or `font.name=Arial`.
            string? rawFontName = properties.GetValueOrDefault("font.name")
                ?? properties.GetValueOrDefault("font");
            if (rawFontName != null)
            {
                rPr.AppendChild(new Drawing.LatinFont { Typeface = rawFontName });
                rPr.AppendChild(new Drawing.EastAsianFont { Typeface = rawFontName });
            }

            var pPr = new Drawing.ParagraphProperties();
            if (properties.TryGetValue("align", out var shpAlign))
            {
                pPr.Alignment = shpAlign.ToLowerInvariant() switch
                {
                    "center" or "c" or "ctr" => Drawing.TextAlignmentTypeValues.Center,
                    "right" or "r" => Drawing.TextAlignmentTypeValues.Right,
                    _ => Drawing.TextAlignmentTypeValues.Left
                };
            }

            txBody.AppendChild(new Drawing.Paragraph(
                pPr,
                new Drawing.Run(rPr, new Drawing.Text(line))
            ));
        }

        var shape = new XDR.Shape(
            new XDR.NonVisualShapeProperties(
                new XDR.NonVisualDrawingProperties { Id = shpId, Name = shpName },
                new XDR.NonVisualShapeDrawingProperties()
            ),
            spPr,
            txBody
        );

        var shpAnchor = new XDR.TwoCellAnchor(
            new XDR.FromMarker(
                new XDR.ColumnId(sx.ToString()),
                new XDR.ColumnOffset("0"),
                new XDR.RowId(sy.ToString()),
                new XDR.RowOffset("0")
            ),
            new XDR.ToMarker(
                new XDR.ColumnId((sx + sw).ToString()),
                new XDR.ColumnOffset("0"),
                new XDR.RowId((sy + sh).ToString()),
                new XDR.RowOffset("0")
            ),
            shape,
            new XDR.ClientData()
        );

        shpDrawingsPart.WorksheetDrawing.AppendChild(shpAnchor);
        shpDrawingsPart.WorksheetDrawing.Save();

        var shpAnchors = shpDrawingsPart.WorksheetDrawing.Elements<XDR.TwoCellAnchor>()
            .Where(a => a.Descendants<XDR.Shape>().Any()).ToList();
        var shpIdx = shpAnchors.IndexOf(shpAnchor) + 1;

        return $"/{shpSheetName}/shape[{shpIdx}]";
    }

    private string AddSlicer(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        return AddSlicer(parentPath, properties);
    }

    private string AddSparkline(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        var spkSegments = parentPath.TrimStart('/').Split('/', 2);
        var spkSheetName = spkSegments[0];
        var spkWorksheet = FindWorksheet(spkSheetName)
            ?? throw new ArgumentException($"Sheet not found: {spkSheetName}");

        // CONSISTENCY(canonical-key): 'location'/'dataRange' are canonical;
        // 'cell'/'range'/'data' retained as legacy aliases.
        var spkCell = properties.GetValueOrDefault("location")
            ?? properties.GetValueOrDefault("cell")
            ?? throw new ArgumentException("Sparkline requires 'location' (or 'cell') property (e.g. F1)");
        var spkRange = properties.GetValueOrDefault("dataRange")
            ?? properties.GetValueOrDefault("datarange")
            ?? properties.GetValueOrDefault("range")
            ?? properties.GetValueOrDefault("data")
            ?? throw new ArgumentException("Sparkline requires 'dataRange' (or 'range'/'data') property (e.g. A1:E1)");

        // Determine sparkline type
        var spkTypeStr = properties.GetValueOrDefault("type", "line").ToLowerInvariant();
        var spkType = spkTypeStr switch
        {
            "column" => X14.SparklineTypeValues.Column,
            "stacked" or "winloss" or "win-loss" => X14.SparklineTypeValues.Stacked,
            _ => X14.SparklineTypeValues.Line
        };

        // Build the SparklineGroup
        var spkGroup = new X14.SparklineGroup();
        // Only set Type attribute for non-line (line is default in OOXML)
        if (spkType != X14.SparklineTypeValues.Line)
            spkGroup.Type = spkType;

        // Series color
        var spkColor = properties.GetValueOrDefault("color", "4472C4");
        spkGroup.SeriesColor = new X14.SeriesColor { Rgb = ParseHelpers.NormalizeArgbColor(spkColor) };

        // Negative color
        if (properties.TryGetValue("negativecolor", out var negColor))
            spkGroup.NegativeColor = new X14.NegativeColor { Rgb = ParseHelpers.NormalizeArgbColor(negColor) };

        // Boolean flags
        if (properties.TryGetValue("markers", out var markersVal) && ParseHelpers.IsTruthy(markersVal))
            spkGroup.Markers = true;
        if (properties.TryGetValue("highpoint", out var highVal) && ParseHelpers.IsTruthy(highVal))
            spkGroup.High = true;
        if (properties.TryGetValue("lowpoint", out var lowVal) && ParseHelpers.IsTruthy(lowVal))
            spkGroup.Low = true;
        if (properties.TryGetValue("firstpoint", out var firstVal) && ParseHelpers.IsTruthy(firstVal))
            spkGroup.First = true;
        if (properties.TryGetValue("lastpoint", out var lastVal) && ParseHelpers.IsTruthy(lastVal))
            spkGroup.Last = true;
        if (properties.TryGetValue("negative", out var negVal) && ParseHelpers.IsTruthy(negVal))
            spkGroup.Negative = true;

        // Marker colors
        if (properties.TryGetValue("highmarkercolor", out var highMC))
            spkGroup.HighMarkerColor = new X14.HighMarkerColor { Rgb = ParseHelpers.NormalizeArgbColor(highMC) };
        if (properties.TryGetValue("lowmarkercolor", out var lowMC))
            spkGroup.LowMarkerColor = new X14.LowMarkerColor { Rgb = ParseHelpers.NormalizeArgbColor(lowMC) };
        if (properties.TryGetValue("firstmarkercolor", out var firstMC))
            spkGroup.FirstMarkerColor = new X14.FirstMarkerColor { Rgb = ParseHelpers.NormalizeArgbColor(firstMC) };
        if (properties.TryGetValue("lastmarkercolor", out var lastMC))
            spkGroup.LastMarkerColor = new X14.LastMarkerColor { Rgb = ParseHelpers.NormalizeArgbColor(lastMC) };
        if (properties.TryGetValue("markerscolor", out var markersMC))
            spkGroup.MarkersColor = new X14.MarkersColor { Rgb = ParseHelpers.NormalizeArgbColor(markersMC) };

        // Line weight
        if (properties.TryGetValue("lineweight", out var lwVal) && double.TryParse(lwVal, out var lw))
            spkGroup.LineWeight = lw;

        // Build the Sparkline element
        // Ensure range includes sheet reference
        var spkFormulaRef = spkRange.Contains('!') ? spkRange : $"{spkSheetName}!{spkRange}";
        var sparkline = new X14.Sparkline
        {
            Formula = new DocumentFormat.OpenXml.Office.Excel.Formula(spkFormulaRef),
            ReferenceSequence = new DocumentFormat.OpenXml.Office.Excel.ReferenceSequence(spkCell)
        };
        var sparklines = new X14.Sparklines();
        sparklines.Append(sparkline);
        spkGroup.Append(sparklines);

        // Add to worksheet extension list
        var spkWs = GetSheet(spkWorksheet);
        var spkExtList = spkWs.GetFirstChild<WorksheetExtensionList>()
            ?? spkWs.AppendChild(new WorksheetExtensionList());

        // Find existing sparkline extension or create new one
        var spkExt = spkExtList.Elements<WorksheetExtension>()
            .FirstOrDefault(e => e.Uri == "{05C60535-1F16-4fd2-B633-E4A46CF9E463}");
        X14.SparklineGroups spkGroups;
        if (spkExt != null)
        {
            spkGroups = spkExt.GetFirstChild<X14.SparklineGroups>()
                ?? spkExt.AppendChild(new X14.SparklineGroups());
        }
        else
        {
            spkExt = new WorksheetExtension { Uri = "{05C60535-1F16-4fd2-B633-E4A46CF9E463}" };
            spkExt.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            spkGroups = new X14.SparklineGroups();
            spkExt.Append(spkGroups);
            spkExtList.Append(spkExt);
        }

        spkGroups.Append(spkGroup);

        // Ensure worksheet root declares mc:Ignorable="x14" so Excel opts-in
        // to the x14 extension namespace where sparklines live. Without this,
        // Excel silently drops the entire extLst block and no sparklines render.
        var spkWsRoot = spkWs;
        const string spkMcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        const string spkX14Ns = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
        if (spkWsRoot.LookupNamespace("mc") == null)
            spkWsRoot.AddNamespaceDeclaration("mc", spkMcNs);
        if (spkWsRoot.LookupNamespace("x14") == null)
            spkWsRoot.AddNamespaceDeclaration("x14", spkX14Ns);
        var spkIgnorable = spkWsRoot.MCAttributes?.Ignorable?.Value ?? "";
        if (!spkIgnorable.Split(' ').Contains("x14"))
        {
            spkWsRoot.MCAttributes ??= new MarkupCompatibilityAttributes();
            spkWsRoot.MCAttributes.Ignorable = string.IsNullOrEmpty(spkIgnorable) ? "x14" : $"{spkIgnorable} x14";
        }

        SaveWorksheet(spkWorksheet);

        // Count all sparkline groups to determine index
        var allSpkGroups = spkGroups.Elements<X14.SparklineGroup>().ToList();
        var spkIdx = allSpkGroups.IndexOf(spkGroup) + 1;
        return $"/{spkSheetName}/sparkline[{spkIdx}]";
    }

}
