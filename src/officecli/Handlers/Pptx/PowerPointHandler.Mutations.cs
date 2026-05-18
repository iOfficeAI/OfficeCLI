// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    public string? Remove(string path)
    {
        // CONSISTENCY(null-path-guard): callers that pass null get an
        // ArgumentNullException instead of a confusing downstream NRE.
        // Mirrors the Word/Excel guards on the same surface.
        ArgumentNullException.ThrowIfNull(path);

        // CONSISTENCY(container-remove-guard): reject removal of required
        // structural container paths. Matches the Word/Excel guards.
        if (IsProtectedPptxContainerPath(path))
            throw new ArgumentException(
                $"Cannot remove container element '{path}': it is a required structural element of the document.");

        path = NormalizePptxPathSegmentCasing(path);
        path = NormalizeCellPath(path);
        path = ResolveIdPath(path);
        path = ResolveLastPredicates(path);

        // /slide[*] — remove every slide. Used by `dump` to clear a non-empty
        // target before replaying slide content, so round-trip replay onto an
        // already-populated deck does not double the slide count. No-op when
        // the deck is empty (clean replay case).
        if (path == "/slide[*]")
        {
            var presentationPart0 = _doc.PresentationPart
                ?? throw new InvalidOperationException("Presentation not found");
            var presentation0 = presentationPart0.Presentation
                ?? throw new InvalidOperationException("No presentation");
            var slideIdList0 = presentation0.GetFirstChild<SlideIdList>();
            if (slideIdList0 != null)
            {
                foreach (var sid in slideIdList0.Elements<SlideId>().ToList())
                {
                    var rid = sid.RelationshipId?.Value;
                    sid.Remove();
                    if (rid != null)
                    {
                        try { presentationPart0.DeletePart(presentationPart0.GetPartById(rid)); }
                        catch { /* part already gone */ }
                    }
                }
                presentation0.Save();
            }
            return null;
        }

        // BUG-R36-B11: /slide[N]/comment[M] removal.
        var cmtRemoveMatch = Regex.Match(path, @"^/slide\[(\d+)\]/comment\[(\d+)\]$");
        if (cmtRemoveMatch.Success)
        {
            if (!RemoveSlideComment(path))
                throw new ArgumentException($"Comment not found: {path}");
            return null;
        }

        // Modern p188 threaded comment removal — top-level path removes the
        // whole thread (mirror PowerPoint UI); reply path removes just one
        // reply.
        var mcRemoveMatch = Regex.Match(path,
            @"^/slide\[(\d+)\]/moderncomment\[(\d+)\](?:/reply\[(\d+)\])?$");
        if (mcRemoveMatch.Success)
        {
            if (!RemoveModernComment(path))
                throw new ArgumentException($"Modern comment not found: {path}");
            return null;
        }

        // Handle /slide[N]/notes path (no index bracket)
        var notesMatch = Regex.Match(path, @"^/slide\[(\d+)\]/notes$");
        if (notesMatch.Success)
        {
            var notesSlideIdx = int.Parse(notesMatch.Groups[1].Value);
            var notesSlideParts = GetSlideParts().ToList();
            if (notesSlideIdx < 1 || notesSlideIdx > notesSlideParts.Count)
                throw new ArgumentException($"Slide {notesSlideIdx} not found (total: {notesSlideParts.Count})");
            var notesSlidePart = notesSlideParts[notesSlideIdx - 1];
            if (notesSlidePart.NotesSlidePart != null)
            {
                notesSlidePart.DeletePart(notesSlidePart.NotesSlidePart);
            }
            return null;
        }

        // Handle /slide[N]/table[M]/tr[R] — remove a table row
        var tableRowMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]$");
        if (tableRowMatch.Success)
        {
            var trSlideIdx = int.Parse(tableRowMatch.Groups[1].Value);
            var tableIdx = int.Parse(tableRowMatch.Groups[2].Value);
            var rowIdx = int.Parse(tableRowMatch.Groups[3].Value);

            var trSlideParts = GetSlideParts().ToList();
            if (trSlideIdx < 1 || trSlideIdx > trSlideParts.Count)
                throw new ArgumentException($"Slide {trSlideIdx} not found (total: {trSlideParts.Count})");

            var trSlidePart = trSlideParts[trSlideIdx - 1];
            var trShapeTree = GetSlide(trSlidePart).CommonSlideData?.ShapeTree
                ?? throw new InvalidOperationException("Slide has no shapes");

            var tables = trShapeTree.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<Drawing.Table>().Any()).ToList();
            if (tableIdx < 1 || tableIdx > tables.Count)
                throw new ArgumentException($"Table {tableIdx} not found (total: {tables.Count})");

            var table = tables[tableIdx - 1].Descendants<Drawing.Table>().First();
            var rows = table.Elements<Drawing.TableRow>().ToList();
            if (rowIdx < 1 || rowIdx > rows.Count)
                throw new ArgumentException($"Row {rowIdx} not found (total: {rows.Count})");

            // BUG-R2-table-merge BUG-6b: a table with 0 rows is invalid OOXML —
            // PowerPoint errors on open. Reject removing the only remaining
            // row; users must remove the table itself.
            if (rows.Count == 1)
                throw new ArgumentException(
                    "Cannot remove the last row of a table. Remove the table itself instead (a table with 0 rows is invalid OOXML).");

            // BUG-R2-table-merge BUG-4b: snapshot orphan-vMerge fixups before
            // removal. Any cell in the doomed row with rowSpan>1 anchors a
            // vertical merge whose continuation cells (vMerge=true) below
            // become invisible if not promoted. Record the column slot and
            // remaining-rows budget so the post-Remove pass can clear them.
            var rowSpanFixups = new List<(int colIdx, int budget)>();
            var anchorRow = rows[rowIdx - 1];
            int slotAcc = 0;
            foreach (var anchorCell in anchorRow.Elements<Drawing.TableCell>())
            {
                int gSpan = anchorCell.GridSpan?.Value ?? 1;
                int rSpan = anchorCell.RowSpan?.Value ?? 1;
                if (rSpan > 1)
                    rowSpanFixups.Add((slotAcc, rSpan - 1));
                slotAcc += gSpan;
            }

            anchorRow.Remove();

            if (rowSpanFixups.Count > 0)
            {
                var rowsAfter = table.Elements<Drawing.TableRow>().ToList();
                foreach (var (colIdx, budget) in rowSpanFixups)
                {
                    int cleared = 0;
                    for (int ri = rowIdx - 1; ri < rowsAfter.Count && cleared < budget; ri++)
                    {
                        var fixCell = ResolvePptxCellAtSlot(rowsAfter[ri], colIdx);
                        if (fixCell == null) break;
                        bool isContinuation = fixCell.VerticalMerge?.Value == true;
                        if (!isContinuation) break;
                        fixCell.VerticalMerge = null;
                        cleared++;
                    }
                }
            }

            return null;
        }

        // Handle /slide[N]/table[M]/col[C] — remove a table column
        var tableColMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]/col\[(\d+)\]$");
        if (tableColMatch.Success)
        {
            var colSlideIdx = int.Parse(tableColMatch.Groups[1].Value);
            var colTableIdx = int.Parse(tableColMatch.Groups[2].Value);
            var colIdx = int.Parse(tableColMatch.Groups[3].Value);

            var (colSlidePart, colTable) = ResolveTable(colSlideIdx, colTableIdx);
            var tableGrid = colTable.GetFirstChild<Drawing.TableGrid>()
                ?? throw new InvalidOperationException("Table has no grid");
            var gridCols = tableGrid.Elements<Drawing.GridColumn>().ToList();
            if (colIdx < 1 || colIdx > gridCols.Count)
                throw new ArgumentException($"Column {colIdx} not found (total: {gridCols.Count})");

            // Remove the grid column
            gridCols[colIdx - 1].Remove();

            // Remove the corresponding cell from each row
            foreach (var row in colTable.Elements<Drawing.TableRow>())
            {
                var cells = row.Elements<Drawing.TableCell>().ToList();
                if (colIdx <= cells.Count)
                    cells[colIdx - 1].Remove();
            }

            // Update GraphicFrame container width
            var graphicFrame = colTable.Ancestors<GraphicFrame>().FirstOrDefault();
            if (graphicFrame?.Transform?.Extents != null)
            {
                long totalColWidth = tableGrid.Elements<Drawing.GridColumn>()
                    .Sum(gc => gc.Width?.Value ?? 914400);
                graphicFrame.Transform.Extents.Cx = totalColWidth;
            }

            GetSlide(colSlidePart).Save();
            return null;
        }

        // BUG C-P-4: /slide[N]/shape[M]/animation[K] removal. Mirrors the
        // enumeration model used by AddAnimation/Get/Set (EnumerateShape-
        // AnimationCTns) so Add/Get/Set/Remove all share the same indexing.
        var animRemoveMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/animation\[(\d+)\]$");
        if (animRemoveMatch.Success)
        {
            var animSlideIdx = int.Parse(animRemoveMatch.Groups[1].Value);
            var animShapeIdx = int.Parse(animRemoveMatch.Groups[2].Value);
            var animKIdx = int.Parse(animRemoveMatch.Groups[3].Value);
            var (animSlidePart, animShape) = ResolveShape(animSlideIdx, animShapeIdx);
            RemoveSingleShapeAnimation(animSlidePart, animShape, animKIdx);
            GetSlide(animSlidePart).Save();
            return null;
        }

        // CONSISTENCY(master-layout-shape-edit): typed Remove on master/layout
        // shape paths. Mirrors the Add/Set branches added in 237b7fb4; the
        // parent path (everything before /shape[K]) is resolved via the shared
        // TryResolveMasterOrLayoutShapeParent helper so all three forms work:
        //   /slidemaster[N]/shape[K]
        //   /slidelayout[N]/shape[K]
        //   /slidemaster[N]/slidelayout[L]/shape[K]
        // No referential cleanup needed — slides reference layouts, not the
        // shapes inside them, so dropping a shape from a master/layout shape
        // tree is a pure tree-edit. The container-remove guard above already
        // rejects removing the master/layout part itself.
        var masterLayoutShapeMatch = Regex.Match(path,
            @"^(/slidemaster\[\d+\](?:/slidelayout\[\d+\])?|/slidelayout\[\d+\])/shape\[(\d+)\]$",
            RegexOptions.IgnoreCase);
        if (masterLayoutShapeMatch.Success)
        {
            var parentPath = masterLayoutShapeMatch.Groups[1].Value;
            var shapeIdx1 = int.Parse(masterLayoutShapeMatch.Groups[2].Value);
            var resolved = TryResolveMasterOrLayoutShapeParent(parentPath)
                ?? throw new ArgumentException($"Invalid master/layout parent path: {parentPath}");
            var (mlTree, _, mlRoot, _) = resolved;
            var mlShapes = mlTree.Elements<Shape>().ToList();
            if (shapeIdx1 < 1 || shapeIdx1 > mlShapes.Count)
                throw new ArgumentException($"Shape {shapeIdx1} not found (total: {mlShapes.Count})");
            mlShapes[shapeIdx1 - 1].Remove();
            mlRoot.Save();
            return null;
        }

        // CONSISTENCY(pptx-group-flatten): optional /group[K] ancestors between
        // /slide[N] and the leaf element type, so Remove works on paths Query
        // emits (e.g. /slide[1]/group[2]/shape[3]) without requiring callers
        // to strip the group prefix. The ancestor segment is tied to the leaf
        // segment so /slide[1]/group[1] still parses as "remove the group at
        // root" (ancestor empty, leaf = group[1]) rather than "remove slide
        // with group ancestor 1".
        var slideMatch = Regex.Match(path, @"^/slide\[(\d+)\](?:((?:/group\[\d+\])*)/(\w+)\[(\w+)\])?$");
        if (!slideMatch.Success)
            throw new ArgumentException($"Invalid path: {path}. Expected format: /slide[N] or /slide[N]/element[M] (e.g. /slide[1], /slide[1]/shape[2])");

        var slideIdx = int.Parse(slideMatch.Groups[1].Value);
        var groupAncestorChain = slideMatch.Groups[2].Value;

        if (!slideMatch.Groups[3].Success)
        {
            // Remove entire slide
            var presentationPart = _doc.PresentationPart
                ?? throw new InvalidOperationException("Presentation not found");
            var presentation = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
            var slideIdList = presentation.GetFirstChild<SlideIdList>()
                ?? throw new InvalidOperationException("No slides");

            var slideIds = slideIdList.Elements<SlideId>().ToList();
            if (slideIdx < 1 || slideIdx > slideIds.Count)
                throw new ArgumentException($"Slide {slideIdx} not found (total: {slideIds.Count})");

            var slideId = slideIds[slideIdx - 1];
            var relId = slideId.RelationshipId?.Value;
            slideId.Remove();
            if (relId != null)
                presentationPart.DeletePart(presentationPart.GetPartById(relId));
            presentation.Save();
            return null;
        }

        // Remove element from slide
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shapes");

        // Walk down group ancestors to scope element lookup to the correct
        // container. Path /slide[1]/group[2]/shape[3] resolves "shape[3]"
        // inside the second group of slide 1, not at the slide root.
        // `container` is what the element lookup runs on; `shapeTree` is
        // kept pointing at the slide root for helpers that need slide-wide
        // context (picture cleanup, zoom resolution, etc).
        OpenXmlCompositeElement container = shapeTree;
        if (!string.IsNullOrEmpty(groupAncestorChain))
        {
            foreach (Match gm in Regex.Matches(groupAncestorChain, @"/group\[(\d+)\]"))
            {
                var gIdx = int.Parse(gm.Groups[1].Value);
                var groupsAtScope = container.Elements<GroupShape>().ToList();
                if (gIdx < 1 || gIdx > groupsAtScope.Count)
                    throw new ArgumentException($"Group {gIdx} not found in scope (have {groupsAtScope.Count})");
                container = groupsAtScope[gIdx - 1];
            }
        }

        var elementType = slideMatch.Groups[3].Value;
        var elementIdxToken = slideMatch.Groups[4].Value;

        // Placeholder remove: accept both numeric index (/slide[N]/placeholder[K])
        // and type-name selectors (/slide[N]/placeholder[title]) that Query emits.
        // ResolvePlaceholderShape materializes layout-inherited placeholders onto
        // the slide; removing that materialized Shape is the canonical delete.
        if (elementType == "placeholder")
        {
            if (!string.IsNullOrEmpty(groupAncestorChain))
                throw new ArgumentException("placeholder remove does not support /group[K] ancestors");
            var phShape = ResolvePlaceholderShape(slidePart, elementIdxToken);
            var phShapeId = phShape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0;
            if (phShapeId != 0)
                RemoveShapeAnimations(GetSlide(slidePart), (uint)phShapeId);
            phShape.Remove();
            GetSlide(slidePart).Save();
            return null;
        }

        if (!int.TryParse(elementIdxToken, out var elementIdx))
            throw new ArgumentException($"Invalid index '{elementIdxToken}' for element type '{elementType}'. Expected a positive integer.");

        if (elementType == "shape")
        {
            var shapes = container.Elements<Shape>().ToList();
            if (elementIdx < 1 || elementIdx > shapes.Count)
                throw new ArgumentException($"Shape {elementIdx} not found");
            var shapeToRemove = shapes[elementIdx - 1];
            var shapeId = shapeToRemove.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0;
            if (shapeId > 0)
                RemoveShapeAnimations(GetSlide(slidePart), (uint)shapeId);
            shapeToRemove.Remove();
        }
        else if (elementType is "picture" or "pic" or "video" or "audio")
        {
            List<Picture> pics;
            if (elementType is "video")
                pics = container.Elements<Picture>()
                    .Where(p => p.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<Drawing.VideoFromFile>() != null).ToList();
            else if (elementType is "audio")
                pics = container.Elements<Picture>()
                    .Where(p => p.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<Drawing.AudioFromFile>() != null).ToList();
            else
                pics = container.Elements<Picture>().ToList();

            if (elementIdx < 1 || elementIdx > pics.Count)
                throw new ArgumentException($"{elementType} {elementIdx} not found (total: {pics.Count})");

            var pic = pics[elementIdx - 1];
            RemovePictureWithCleanup(slidePart, shapeTree, pic);
        }
        else if (elementType == "table")
        {
            var tables = container.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<Drawing.Table>().Any()).ToList();
            if (elementIdx < 1 || elementIdx > tables.Count)
                throw new ArgumentException($"Table {elementIdx} not found");
            tables[elementIdx - 1].Remove();
        }
        else if (elementType == "chart")
        {
            var charts = container.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<C.ChartReference>().Any()).ToList();
            if (elementIdx < 1 || elementIdx > charts.Count)
                throw new ArgumentException($"Chart {elementIdx} not found");
            var chartGf = charts[elementIdx - 1];
            // Clean up ChartPart
            var chartRef = chartGf.Descendants<C.ChartReference>().FirstOrDefault();
            if (chartRef?.Id?.Value != null)
            {
                try { slidePart.DeletePart(chartRef.Id.Value); } catch { }
            }
            chartGf.Remove();
        }
        else if (elementType is "connector" or "connection")
        {
            var connectors = container.Elements<ConnectionShape>().ToList();
            if (elementIdx < 1 || elementIdx > connectors.Count)
                throw new ArgumentException($"Connector {elementIdx} not found");
            connectors[elementIdx - 1].Remove();
        }
        else if (elementType == "group")
        {
            // Ungroup: move children back to parent container (slide root or outer group),
            // then remove group. `container` is the parent of the group being removed,
            // so children naturally land at the right level — root-level ungroup keeps
            // children at slide root; nested-group ungroup moves them up one level only.
            var groups = container.Elements<GroupShape>().ToList();
            if (elementIdx < 1 || elementIdx > groups.Count)
                throw new ArgumentException($"Group {elementIdx} not found");
            var group = groups[elementIdx - 1];
            var children = group.ChildElements
                .Where(e => e is Shape or Picture or ConnectionShape or GraphicFrame or GroupShape)
                .ToList();
            foreach (var child in children)
            {
                child.Remove();
                container.AppendChild(child);
            }
            group.Remove();
        }
        else if (elementType is "3dmodel" or "model3d")
        {
            var model3dElements = GetModel3DElements(shapeTree);
            if (elementIdx < 1 || elementIdx > model3dElements.Count)
                throw new ArgumentException($"3D model {elementIdx} not found (total: {model3dElements.Count})");
            var m3dAc = model3dElements[elementIdx - 1];
            // Clean up model part and image parts
            var m3dRNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            foreach (var el in m3dAc.Descendants().Where(d => d.LocalName == "blip" || d.LocalName == "model3d"))
            {
                var embedAttr = el.GetAttribute("embed", m3dRNs);
                if (!string.IsNullOrEmpty(embedAttr.Value))
                {
                    try { slidePart.DeletePart(embedAttr.Value); } catch { }
                }
            }
            m3dAc.Remove();
        }
        else if (elementType is "zoom" or "slidezoom")
        {
            var zoomElements = GetZoomElements(shapeTree);
            if (elementIdx < 1 || elementIdx > zoomElements.Count)
                throw new ArgumentException($"Zoom {elementIdx} not found (total: {zoomElements.Count})");
            var zmAc = zoomElements[elementIdx - 1];
            // Clean up image relationship if not referenced by other elements
            var zmBlip = zmAc.Descendants().FirstOrDefault(d => d.LocalName == "blip");
            if (zmBlip != null)
            {
                var rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                var embedAttr = zmBlip.GetAttribute("embed", rNs);
                if (!string.IsNullOrEmpty(embedAttr.Value))
                {
                    var relId = embedAttr.Value;
                    // Check if any other element references this image
                    zmAc.Remove();
                    var slideXml = GetSlide(slidePart).OuterXml;
                    if (!slideXml.Contains(relId))
                    {
                        try { slidePart.DeletePart(relId); } catch { }
                    }
                    GetSlide(slidePart).Save();
                    return null;
                }
            }
            zmAc.Remove();
        }
        else if (elementType is "ole" or "object" or "embed")
        {
            // Remove the GraphicFrame wrapper whose graphicData hosts a
            // strong-typed p:oleObj. Index is 1-based among OLE frames on
            // this slide. Also deletes the backing embedded part and the
            // icon image part so the package doesn't bloat with orphaned
            // binaries — same rationale as the picture-replacement quirk
            // noted in CLAUDE.md.
            var oleFrames = shapeTree.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<DocumentFormat.OpenXml.Presentation.OleObject>().Any())
                .ToList();
            if (elementIdx < 1 || elementIdx > oleFrames.Count)
                throw new ArgumentException($"OLE object {elementIdx} not found (total: {oleFrames.Count})");
            var oleFrame = oleFrames[elementIdx - 1];
            var oleObjEl = oleFrame.Descendants<DocumentFormat.OpenXml.Presentation.OleObject>().First();
            // 1. Delete the embedded payload part by rel id.
            if (oleObjEl.Id?.Value is string embedRel && !string.IsNullOrEmpty(embedRel))
            {
                try { slidePart.DeletePart(embedRel); } catch { }
            }
            // 2. Delete the inner icon image part (Blip inside p:pic).
            var iconBlip = oleObjEl.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
            if (iconBlip?.Embed?.Value is string iconRel && !string.IsNullOrEmpty(iconRel))
            {
                try { slidePart.DeletePart(iconRel); } catch { }
            }
            oleFrame.Remove();
        }
        else
        {
            throw new ArgumentException($"Unknown element type: {elementType}. Supported: shape, picture, video, audio, table, chart, connector/connection, group, zoom, 3dmodel, ole");
        }

        GetSlide(slidePart).Save();
        return null;
    }

    public string Move(string sourcePath, string? targetParentPath, InsertPosition? position)
    {
        var index = position?.Index;
        sourcePath = ResolveIdPath(sourcePath);
        sourcePath = ResolveLastPredicates(sourcePath);
        if (targetParentPath != null)
        {
            targetParentPath = ResolveIdPath(targetParentPath);
            targetParentPath = ResolveLastPredicates(targetParentPath);
        }

        // Infer --to from --after/--before full path if not specified
        var anchorFullPath = position?.After ?? position?.Before;
        if (string.IsNullOrEmpty(targetParentPath) && anchorFullPath != null && anchorFullPath.StartsWith("/"))
        {
            var resolvedAnchor = ResolveIdPath(anchorFullPath);
            var lastSlash = resolvedAnchor.LastIndexOf('/');
            if (lastSlash > 0)
                targetParentPath = resolvedAnchor[..lastSlash];
        }

        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation not found");
        var slideParts = GetSlideParts().ToList();

        // Case 0: Move table row within the same table.
        // Path: /slide[N]/table[K]/tr[R]. Cross-table row moves are out of
        // scope (column counts may differ; user can copy + remove instead).
        var trMoveMatch = Regex.Match(sourcePath, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]$");
        if (trMoveMatch.Success)
        {
            return MoveTableRow(trMoveMatch, position, targetParentPath);
        }

        // Case 0b: Move table column within the same table.
        // Path: /slide[N]/table[K]/col[C]. Same-table only — column has no
        // standalone OOXML element (it's gridCol + per-row tc), and merging
        // grids across tables of different widths is ambiguous.
        var colMoveMatch = Regex.Match(sourcePath, @"^/slide\[(\d+)\]/table\[(\d+)\]/col\[(\d+)\]$");
        if (colMoveMatch.Success)
        {
            return MoveTableColumn(colMoveMatch, position, targetParentPath);
        }

        // Case 1: Move entire slide (reorder)
        var slideOnlyMatch = Regex.Match(sourcePath, @"^/slide\[(\d+)\]$");
        if (slideOnlyMatch.Success)
        {
            var slideIdx = int.Parse(slideOnlyMatch.Groups[1].Value);
            var movePresentation = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
            var slideIdList = movePresentation.GetFirstChild<SlideIdList>()
                ?? throw new InvalidOperationException("No slides");
            var slideIds = slideIdList.Elements<SlideId>().ToList();
            if (slideIdx < 1 || slideIdx > slideIds.Count)
                throw new ArgumentException($"Slide {slideIdx} not found (total: {slideIds.Count})");

            var slideId = slideIds[slideIdx - 1];

            // Resolve after/before anchor BEFORE removing
            SlideId? afterAnchor = null, beforeAnchor = null;
            if (position?.After != null)
            {
                var afterMatch = Regex.Match(position.After.StartsWith("/") ? position.After : "/" + position.After, @"/slide\[(\d+)\]");
                if (afterMatch.Success)
                {
                    var ai = int.Parse(afterMatch.Groups[1].Value);
                    if (ai >= 1 && ai <= slideIds.Count) afterAnchor = slideIds[ai - 1];
                }
                if (afterAnchor == null) throw new ArgumentException($"After anchor not found: {position.After}");
            }
            else if (position?.Before != null)
            {
                var beforeMatch = Regex.Match(position.Before.StartsWith("/") ? position.Before : "/" + position.Before, @"/slide\[(\d+)\]");
                if (beforeMatch.Success)
                {
                    var bi = int.Parse(beforeMatch.Groups[1].Value);
                    if (bi >= 1 && bi <= slideIds.Count) beforeAnchor = slideIds[bi - 1];
                }
                if (beforeAnchor == null) throw new ArgumentException($"Before anchor not found: {position.Before}");
            }

            // Self-move guard: if the anchor is the slide being moved, the anchor's
            // parent will be null after Remove() and InsertAfterSelf/InsertBeforeSelf
            // will throw InvalidOperationException. Detect and no-op the move.
            // CONSISTENCY(slide-move): same guard for both After and Before anchors.
            if (ReferenceEquals(afterAnchor, slideId) || ReferenceEquals(beforeAnchor, slideId))
            {
                // Moving a slide after/before itself is a no-op.
                var sameNewSlideIds = slideIdList.Elements<SlideId>().ToList();
                var sameIdx = sameNewSlideIds.IndexOf(slideId) + 1;
                return $"/slide[{sameIdx}]";
            }

            slideId.Remove();

            if (afterAnchor != null)
                afterAnchor.InsertAfterSelf(slideId);
            else if (beforeAnchor != null)
                beforeAnchor.InsertBeforeSelf(slideId);
            else if (index.HasValue)
            {
                var remaining = slideIdList.Elements<SlideId>().ToList();
                if (index.Value >= 0 && index.Value < remaining.Count)
                    remaining[index.Value].InsertBeforeSelf(slideId);
                else
                    slideIdList.AppendChild(slideId);
            }
            else
            {
                slideIdList.AppendChild(slideId);
            }

            movePresentation.Save();
            var newSlideIds = slideIdList.Elements<SlideId>().ToList();
            var newIdx = newSlideIds.IndexOf(slideId) + 1;
            return $"/slide[{newIdx}]";
        }

        // Case 2: Move element within/across slides
        var (srcSlidePart, srcElement) = ResolveSlideElement(sourcePath, slideParts);

        // Determine target
        string effectiveParentPath;
        SlidePart tgtSlidePart;
        ShapeTree tgtShapeTree;

        if (string.IsNullOrEmpty(targetParentPath))
        {
            // Reorder within same parent
            tgtSlidePart = srcSlidePart;
            tgtShapeTree = GetSlide(srcSlidePart).CommonSlideData?.ShapeTree
                ?? throw new InvalidOperationException("Slide has no shape tree");
            var srcSlideIdx = slideParts.IndexOf(srcSlidePart) + 1;
            effectiveParentPath = $"/slide[{srcSlideIdx}]";
        }
        else
        {
            effectiveParentPath = targetParentPath;
            var tgtSlideMatch = Regex.Match(targetParentPath, @"^/slide\[(\d+)\]$");
            if (!tgtSlideMatch.Success)
                throw new ArgumentException($"Target must be a slide: /slide[N]");
            var tgtSlideIdx = int.Parse(tgtSlideMatch.Groups[1].Value);
            if (tgtSlideIdx < 1 || tgtSlideIdx > slideParts.Count)
                throw new ArgumentException($"Slide {tgtSlideIdx} not found (total: {slideParts.Count})");
            tgtSlidePart = slideParts[tgtSlideIdx - 1];
            tgtShapeTree = GetSlide(tgtSlidePart).CommonSlideData?.ShapeTree
                ?? throw new InvalidOperationException("Slide has no shape tree");
        }

        // Reject cross-slide move of placeholder shapes (would cause duplicate IDs)
        if (srcSlidePart != tgtSlidePart)
        {
            var nvSpPr = srcElement.Descendants<DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties>().FirstOrDefault();
            if (nvSpPr?.ApplicationNonVisualDrawingProperties?.PlaceholderShape != null)
                throw new ArgumentException("Cannot move placeholder shapes across slides");
        }

        // Copy relationships BEFORE removing from source (so rel IDs are still accessible).
        // For cross-slide moves, also capture the original rel ids so we can
        // delete now-orphaned parts from the source slide after the move
        // (e.g. OLE embedded payload + icon blip). Without this, Query("ole")
        // on the source still surfaces the stray EmbeddedPackagePart as an
        // "orphan" OLE node — see Ppt_MoveOleBetweenSlides_SucceedsOrErrorsClearly.
        var oldSourceRelIds = new List<string>();
        if (srcSlidePart != tgtSlidePart)
        {
            var rNsUri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            foreach (var el in srcElement.Descendants().Prepend(srcElement))
            {
                foreach (var attr in el.GetAttributes())
                {
                    if (attr.NamespaceUri == rNsUri && !string.IsNullOrEmpty(attr.Value))
                        oldSourceRelIds.Add(attr.Value);
                }
            }
            CopyRelationships(srcElement, srcSlidePart, tgtSlidePart);
        }

        // Resolve after/before anchor for shape-level move
        OpenXmlElement? shapeAfterAnchor = null, shapeBeforeAnchor = null;
        if (position?.After != null)
        {
            var anchorPath = ResolveIdPath(position.After);
            var (_, anchor) = ResolveSlideElement(anchorPath, slideParts);
            shapeAfterAnchor = anchor;
        }
        else if (position?.Before != null)
        {
            var anchorPath = ResolveIdPath(position.Before);
            var (_, anchor) = ResolveSlideElement(anchorPath, slideParts);
            shapeBeforeAnchor = anchor;
        }

        srcElement.Remove();

        if (shapeAfterAnchor != null)
            shapeAfterAnchor.InsertAfterSelf(srcElement);
        else if (shapeBeforeAnchor != null)
            shapeBeforeAnchor.InsertBeforeSelf(srcElement);
        else
            InsertAtPosition(tgtShapeTree, srcElement, index);

        GetSlide(srcSlidePart).Save();
        if (srcSlidePart != tgtSlidePart)
            GetSlide(tgtSlidePart).Save();

        // Post-move cleanup: delete any source-slide rels the moved element
        // used exclusively, otherwise they linger as "orphan" parts detected
        // by Query("ole") and other listers.
        if (srcSlidePart != tgtSlidePart && oldSourceRelIds.Count > 0)
        {
            var srcSlideXml = GetSlide(srcSlidePart).OuterXml;
            foreach (var oldRelId in oldSourceRelIds.Distinct())
            {
                // Keep rels still referenced anywhere else in the source slide XML.
                if (srcSlideXml.Contains($"\"{oldRelId}\"")) continue;
                try { srcSlidePart.DeletePart(oldRelId); } catch { }
            }
        }

        return ComputeElementPath(effectiveParentPath, srcElement, tgtShapeTree);
    }

    public (string NewPath1, string NewPath2) Swap(string path1, string path2)
    {
        path1 = ResolveIdPath(path1);
        path2 = ResolveIdPath(path2);
        path1 = ResolveLastPredicates(path1);
        path2 = ResolveLastPredicates(path2);
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation not found");
        var slideParts = GetSlideParts().ToList();

        // Case 1: Swap two slides
        var slide1Match = Regex.Match(path1, @"^/slide\[(\d+)\]$");
        var slide2Match = Regex.Match(path2, @"^/slide\[(\d+)\]$");
        if (slide1Match.Success && slide2Match.Success)
        {
            var presentation = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
            var slideIdList = presentation.GetFirstChild<SlideIdList>()
                ?? throw new InvalidOperationException("No slides");
            var slideIds = slideIdList.Elements<SlideId>().ToList();
            var idx1 = int.Parse(slide1Match.Groups[1].Value);
            var idx2 = int.Parse(slide2Match.Groups[1].Value);
            if (idx1 < 1 || idx1 > slideIds.Count) throw new ArgumentException($"Slide {idx1} not found (total: {slideIds.Count})");
            if (idx2 < 1 || idx2 > slideIds.Count) throw new ArgumentException($"Slide {idx2} not found (total: {slideIds.Count})");
            if (idx1 == idx2) return (path1, path2);

            SwapXmlElements(slideIds[idx1 - 1], slideIds[idx2 - 1]);
            presentation.Save();
            return ($"/slide[{idx2}]", $"/slide[{idx1}]");
        }

        // CONSISTENCY(table-sub-paths): same lockstep fix as Move (commit
        // 6ba5bb67) — Swap also needs explicit tr / col branches before
        // falling through to ResolveSlideElement, which only accepts the
        // two-segment /slide[N]/elem[M] form.

        // Case 2a: Swap two table rows (same table only).
        var tr1Match = Regex.Match(path1, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]$");
        var tr2Match = Regex.Match(path2, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]$");
        if (tr1Match.Success && tr2Match.Success)
        {
            var sIdx = int.Parse(tr1Match.Groups[1].Value);
            var tIdx = int.Parse(tr1Match.Groups[2].Value);
            if (int.Parse(tr2Match.Groups[1].Value) != sIdx ||
                int.Parse(tr2Match.Groups[2].Value) != tIdx)
                throw new ArgumentException(
                    $"Cross-table row swap is not supported. Both rows must share /slide[{sIdx}]/table[{tIdx}].");
            var r1 = int.Parse(tr1Match.Groups[3].Value);
            var r2 = int.Parse(tr2Match.Groups[3].Value);

            var (trSlidePart, trTable) = ResolveTable(sIdx, tIdx);
            var rows = trTable.Elements<Drawing.TableRow>().ToList();
            if (r1 < 1 || r1 > rows.Count) throw new ArgumentException($"Row {r1} not found (total: {rows.Count})");
            if (r2 < 1 || r2 > rows.Count) throw new ArgumentException($"Row {r2} not found (total: {rows.Count})");
            if (r1 == r2)
                return ($"/slide[{sIdx}]/table[{tIdx}]/tr[{r1}]", $"/slide[{sIdx}]/table[{tIdx}]/tr[{r2}]");

            SwapXmlElements(rows[r1 - 1], rows[r2 - 1]);
            GetSlide(trSlidePart).Save();
            return ($"/slide[{sIdx}]/table[{tIdx}]/tr[{r2}]", $"/slide[{sIdx}]/table[{tIdx}]/tr[{r1}]");
        }
        if (tr1Match.Success != tr2Match.Success)
            throw new ArgumentException(
                "Both swap paths must be table rows in the same table; mixed types are not supported.");

        // Case 2b: Swap two table columns (same table only). Columns are
        // virtual (gridCol + per-row tc): swap the GridColumn entries in
        // <a:tblGrid>, then swap each row's tc at the matching slot. Each
        // pair shares a parent (same row / same grid), so SwapXmlElements
        // applies; the function does not support cross-parent swaps.
        var col1Match = Regex.Match(path1, @"^/slide\[(\d+)\]/table\[(\d+)\]/col\[(\d+)\]$");
        var col2Match = Regex.Match(path2, @"^/slide\[(\d+)\]/table\[(\d+)\]/col\[(\d+)\]$");
        if (col1Match.Success && col2Match.Success)
        {
            var sIdx = int.Parse(col1Match.Groups[1].Value);
            var tIdx = int.Parse(col1Match.Groups[2].Value);
            if (int.Parse(col2Match.Groups[1].Value) != sIdx ||
                int.Parse(col2Match.Groups[2].Value) != tIdx)
                throw new ArgumentException(
                    $"Cross-table column swap is not supported. Both columns must share /slide[{sIdx}]/table[{tIdx}].");
            var c1 = int.Parse(col1Match.Groups[3].Value);
            var c2 = int.Parse(col2Match.Groups[3].Value);

            var (colSlidePart, colTable) = ResolveTable(sIdx, tIdx);
            var grid = colTable.TableGrid
                ?? throw new InvalidOperationException("Table has no <a:tblGrid>");
            var gridCols = grid.Elements<Drawing.GridColumn>().ToList();
            if (c1 < 1 || c1 > gridCols.Count) throw new ArgumentException($"Column {c1} not found (total: {gridCols.Count})");
            if (c2 < 1 || c2 > gridCols.Count) throw new ArgumentException($"Column {c2} not found (total: {gridCols.Count})");
            if (c1 == c2)
                return ($"/slide[{sIdx}]/table[{tIdx}]/col[{c1}]", $"/slide[{sIdx}]/table[{tIdx}]/col[{c2}]");

            // Reject merges crossing either column slot — same guard the
            // column move/copy use, since a swap that splits a merge
            // produces silently broken cells.
            foreach (var row in colTable.Elements<Drawing.TableRow>())
            {
                var rowCells = row.Elements<Drawing.TableCell>().ToList();
                foreach (var cIdx in new[] { c1, c2 })
                {
                    if (cIdx - 1 >= rowCells.Count) continue;
                    var tc = rowCells[cIdx - 1];
                    var span = tc.GridSpan?.Value ?? 1;
                    var hMerge = tc.HorizontalMerge?.Value ?? false;
                    var vMerge = tc.VerticalMerge?.Value ?? false;
                    if (span > 1 || hMerge || vMerge)
                        throw new ArgumentException(
                            $"Cannot swap column {cIdx}: a row contains a merged cell (gridSpan/hMerge/vMerge) " +
                            "spanning that column. Unmerge before performing column-level operations.");
                }
            }

            SwapXmlElements(gridCols[c1 - 1], gridCols[c2 - 1]);
            foreach (var row in colTable.Elements<Drawing.TableRow>())
            {
                var rowCells = row.Elements<Drawing.TableCell>().ToList();
                if (c1 - 1 < rowCells.Count && c2 - 1 < rowCells.Count)
                    SwapXmlElements(rowCells[c1 - 1], rowCells[c2 - 1]);
            }
            GetSlide(colSlidePart).Save();
            return ($"/slide[{sIdx}]/table[{tIdx}]/col[{c2}]", $"/slide[{sIdx}]/table[{tIdx}]/col[{c1}]");
        }
        if (col1Match.Success != col2Match.Success)
            throw new ArgumentException(
                "Both swap paths must be table columns in the same table; mixed types are not supported.");

        // Case 3: Swap two elements within the same slide
        var (slide1Part, elem1) = ResolveSlideElement(path1, slideParts);
        var (slide2Part, elem2) = ResolveSlideElement(path2, slideParts);
        if (slide1Part != slide2Part)
            throw new ArgumentException("Cannot swap elements on different slides");

        SwapXmlElements(elem1, elem2);
        GetSlide(slide1Part).Save();

        var slideIdx = slideParts.IndexOf(slide1Part) + 1;
        var parentPath = $"/slide[{slideIdx}]";
        var shapeTree = GetSlide(slide1Part).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");
        var newPath1 = ComputeElementPath(parentPath, elem1, shapeTree);
        var newPath2 = ComputeElementPath(parentPath, elem2, shapeTree);
        return (newPath1, newPath2);
    }

    // Resolve the Drawing.TableCell occupying a specific gridCol slot in a
    // pptx row, accounting for gridSpan-merged cells. Returns null if the
    // row's total span is shorter than slot+1.
    private static Drawing.TableCell? ResolvePptxCellAtSlot(Drawing.TableRow trow, int slot)
    {
        int acc = 0;
        foreach (var c in trow.Elements<Drawing.TableCell>())
        {
            int span = c.GridSpan?.Value ?? 1;
            if (slot >= acc && slot < acc + span) return c;
            acc += span;
        }
        return null;
    }

    internal static void SwapXmlElements(OpenXmlElement a, OpenXmlElement b)
    {
        if (a == b || a.Parent == null || b.Parent == null) return;
        var parent = a.Parent;
        var aNext = a.NextSibling();
        var bNext = b.NextSibling();

        a.Remove();
        b.Remove();

        if (aNext == b)
        {
            // A was directly before B: [... A B ...] → [... B A ...]
            if (bNext != null)
                bNext.InsertBeforeSelf(b);
            else
                parent.AppendChild(b);
            b.InsertAfterSelf(a);
        }
        else if (bNext == a)
        {
            // B was directly before A: [... B A ...] → [... A B ...]
            if (aNext != null)
                aNext.InsertBeforeSelf(a);
            else
                parent.AppendChild(a);
            a.InsertBeforeSelf(b);
        }
        else
        {
            // Non-adjacent: insert each where the other was
            if (aNext != null)
                aNext.InsertBeforeSelf(b);
            else
                parent.AppendChild(b);
            if (bNext != null)
                bNext.InsertBeforeSelf(a);
            else
                parent.AppendChild(a);
        }
    }

    public string CopyFrom(string sourcePath, string targetParentPath, InsertPosition? position)
    {
        var index = position?.Index;
        sourcePath = ResolveIdPath(sourcePath);
        targetParentPath = ResolveIdPath(targetParentPath);
        sourcePath = ResolveLastPredicates(sourcePath);
        targetParentPath = ResolveLastPredicates(targetParentPath);
        var slideParts = GetSlideParts().ToList();

        // Table row clone: --from /slide[N]/table[K]/tr[R] [target /slide[N]/table[K]].
        // Same-table only (cross-table row copy is out of scope; column counts
        // may differ silently). If targetParentPath is null/empty, defaults to
        // source table — i.e. "duplicate row in place".
        var trCloneMatch = Regex.Match(sourcePath, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]$");
        if (trCloneMatch.Success)
        {
            return CopyTableRow(trCloneMatch, position, targetParentPath);
        }

        // Table column clone: --from /slide[N]/table[K]/col[C]. Same-table
        // only. Clones the gridCol entry plus the per-row tc cells in lockstep.
        var colCloneMatch = Regex.Match(sourcePath, @"^/slide\[(\d+)\]/table\[(\d+)\]/col\[(\d+)\]$");
        if (colCloneMatch.Success)
        {
            return CopyTableColumn(colCloneMatch, position, targetParentPath);
        }

        // Table cell clone: --from /slide[N]/table[K]/tr[R]/tc[C]. Same-row
        // only — cross-row tc copy is ambiguous (column slot shifts) and
        // cross-table is rejected for the same reason as row/col copies.
        // Without this branch the path falls through to ResolveSlideElement,
        // which only accepts /slide[N]/element[M] and throws "Invalid element
        // path".
        var tcCloneMatch = Regex.Match(sourcePath, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]/tc\[(\d+)\]$");
        if (tcCloneMatch.Success)
        {
            return CopyTableCell(tcCloneMatch, position, targetParentPath);
        }

        // Whole-slide clone: --from /slide[N] to / (or null == "duplicate in
        // place" at presentation root, i.e. append the clone after the source
        // slide).
        var slideCloneMatch = Regex.Match(sourcePath, @"^/slide\[(\d+)\]$");
        if (slideCloneMatch.Success && (targetParentPath is null or "/" or "" or "/presentation"))
        {
            return CloneSlide(slideCloneMatch, slideParts, index);
        }

        var (srcSlidePart, srcElement) = ResolveSlideElement(sourcePath, slideParts);
        var clone = srcElement.CloneNode(true);

        // Assign new unique cNvPr.Id to the clone to avoid duplicate IDs on the target slide
        var cloneNvPr = clone.Descendants<NonVisualDrawingProperties>().FirstOrDefault();
        if (cloneNvPr != null)
        {
            var tgtSlideMatchPre = Regex.Match(targetParentPath, @"^/slide\[(\d+)\]$");
            if (tgtSlideMatchPre.Success)
            {
                var tgtIdx = int.Parse(tgtSlideMatchPre.Groups[1].Value);
                if (tgtIdx >= 1 && tgtIdx <= slideParts.Count)
                {
                    var tgtTree = GetSlide(slideParts[tgtIdx - 1]).CommonSlideData?.ShapeTree;
                    if (tgtTree != null)
                        cloneNvPr.Id = GenerateUniqueShapeId(tgtTree);
                }
            }
        }

        var tgtSlideMatch = Regex.Match(targetParentPath, @"^/slide\[(\d+)\]$");
        if (!tgtSlideMatch.Success)
            throw new ArgumentException($"Target must be a slide: /slide[N]");
        var tgtSlideIdx = int.Parse(tgtSlideMatch.Groups[1].Value);
        if (tgtSlideIdx < 1 || tgtSlideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {tgtSlideIdx} not found (total: {slideParts.Count})");

        var tgtSlidePart = slideParts[tgtSlideIdx - 1];
        var tgtShapeTree = GetSlide(tgtSlidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");

        // Copy relationships if across slides
        if (srcSlidePart != tgtSlidePart)
            CopyRelationships(clone, srcSlidePart, tgtSlidePart);

        InsertAtPosition(tgtShapeTree, clone, index);
        GetSlide(tgtSlidePart).Save();

        return ComputeElementPath(targetParentPath, clone, tgtShapeTree);
    }

    /// <summary>
    /// Move a table row within its table by --before/--after/--index. Cross-table
    /// moves are intentionally rejected: column counts may differ silently and
    /// "move row across tables" has no precedent in the Office UI.
    /// </summary>
    private string MoveTableRow(Match trMatch, InsertPosition? position, string? targetParentPath)
    {
        var slideIdx = int.Parse(trMatch.Groups[1].Value);
        var tableIdx = int.Parse(trMatch.Groups[2].Value);
        var rowIdx = int.Parse(trMatch.Groups[3].Value);

        // If targetParentPath is supplied it must point at the same table.
        if (!string.IsNullOrEmpty(targetParentPath))
        {
            var expected = $"/slide[{slideIdx}]/table[{tableIdx}]";
            if (!string.Equals(targetParentPath, expected, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    $"Cross-table row move is not supported. Source row's table is {expected}; target was {targetParentPath}. " +
                    "Use `add --from <row>` followed by `remove <row>` to copy a row to a different table.");
        }

        var (slidePart, table) = ResolveTable(slideIdx, tableIdx);
        var rows = table.Elements<Drawing.TableRow>().ToList();
        if (rowIdx < 1 || rowIdx > rows.Count)
            throw new ArgumentException($"Row {rowIdx} not found (total: {rows.Count})");
        var row = rows[rowIdx - 1];

        // Resolve --before/--after anchor relative to sibling rows (1-based)
        // before mutating, then convert to a 0-based target position.
        int? targetIdx = null;
        if (position?.After != null || position?.Before != null)
        {
            var anchorPath = ResolveIdPath(position.After ?? position.Before!);
            var anchorMatch = Regex.Match(anchorPath, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]$");
            if (!anchorMatch.Success ||
                int.Parse(anchorMatch.Groups[1].Value) != slideIdx ||
                int.Parse(anchorMatch.Groups[2].Value) != tableIdx)
            {
                throw new ArgumentException(
                    $"Move row anchor must be a row in the same table: /slide[{slideIdx}]/table[{tableIdx}]/tr[N]. Got: {anchorPath}");
            }
            var anchorRowIdx = int.Parse(anchorMatch.Groups[3].Value); // 1-based
            // Self-anchor is a no-op
            if (anchorRowIdx == rowIdx)
                return $"/slide[{slideIdx}]/table[{tableIdx}]/tr[{rowIdx}]";
            targetIdx = position.After != null ? anchorRowIdx : anchorRowIdx - 1; // 0-based
            // Compensate when removing the source shifts later siblings up
            if (rowIdx < anchorRowIdx) targetIdx -= 1;
        }
        else if (position?.Index.HasValue == true)
        {
            targetIdx = position.Index.Value;
        }

        row.Remove();
        var remaining = table.Elements<Drawing.TableRow>().ToList();
        if (targetIdx.HasValue && targetIdx.Value >= 0 && targetIdx.Value < remaining.Count)
            remaining[targetIdx.Value].InsertBeforeSelf(row);
        else
            table.AppendChild(row);

        GetSlide(slidePart).Save();
        var newRows = table.Elements<Drawing.TableRow>().ToList();
        var newRowIdx = newRows.IndexOf(row) + 1;
        return $"/slide[{slideIdx}]/table[{tableIdx}]/tr[{newRowIdx}]";
    }

    /// <summary>
    /// Clone a table row inside the same table (or duplicate-in-place when no
    /// target supplied). Cross-table copies are out of scope to keep grid
    /// width semantics unambiguous.
    /// </summary>
    private string CopyTableRow(Match trMatch, InsertPosition? position, string? targetParentPath)
    {
        var slideIdx = int.Parse(trMatch.Groups[1].Value);
        var tableIdx = int.Parse(trMatch.Groups[2].Value);
        var rowIdx = int.Parse(trMatch.Groups[3].Value);

        if (!string.IsNullOrEmpty(targetParentPath))
        {
            var expected = $"/slide[{slideIdx}]/table[{tableIdx}]";
            if (!string.Equals(targetParentPath, expected, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    $"Cross-table row copy is not supported. Source row's table is {expected}; target was {targetParentPath}.");
        }

        var (slidePart, table) = ResolveTable(slideIdx, tableIdx);
        var rows = table.Elements<Drawing.TableRow>().ToList();
        if (rowIdx < 1 || rowIdx > rows.Count)
            throw new ArgumentException($"Row {rowIdx} not found (total: {rows.Count})");

        var clone = (Drawing.TableRow)rows[rowIdx - 1].CloneNode(true);

        // Resolve --before/--after anchor first (relative to current sibling order).
        int? targetIdx = null;
        if (position?.After != null || position?.Before != null)
        {
            var anchorPath = ResolveIdPath(position.After ?? position.Before!);
            var anchorMatch = Regex.Match(anchorPath, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]$");
            if (!anchorMatch.Success ||
                int.Parse(anchorMatch.Groups[1].Value) != slideIdx ||
                int.Parse(anchorMatch.Groups[2].Value) != tableIdx)
            {
                throw new ArgumentException(
                    $"Copy row anchor must be a row in the same table: /slide[{slideIdx}]/table[{tableIdx}]/tr[N]. Got: {anchorPath}");
            }
            var anchorRowIdx = int.Parse(anchorMatch.Groups[3].Value);
            targetIdx = position.After != null ? anchorRowIdx : anchorRowIdx - 1; // 0-based
        }
        else if (position?.Index.HasValue == true)
        {
            targetIdx = position.Index.Value;
        }

        var siblings = table.Elements<Drawing.TableRow>().ToList();
        if (targetIdx.HasValue && targetIdx.Value >= 0 && targetIdx.Value < siblings.Count)
            siblings[targetIdx.Value].InsertBeforeSelf(clone);
        else
            table.AppendChild(clone);

        GetSlide(slidePart).Save();
        var newRows = table.Elements<Drawing.TableRow>().ToList();
        var newRowIdx = newRows.IndexOf(clone) + 1;
        return $"/slide[{slideIdx}]/table[{tableIdx}]/tr[{newRowIdx}]";
    }

    /// <summary>
    /// Clone a single table cell within its row (same-row only). Mirrors
    /// CopyTableRow: target must be the source row (or null = "duplicate in
    /// place"), --before/--after must point at a sibling tc in the same row.
    /// Cross-row / cross-table cell copy is rejected — the receiving row
    /// would have a different column count than its peers, breaking the
    /// table's grid invariant.
    /// </summary>
    private string CopyTableCell(Match tcMatch, InsertPosition? position, string? targetParentPath)
    {
        var slideIdx = int.Parse(tcMatch.Groups[1].Value);
        var tableIdx = int.Parse(tcMatch.Groups[2].Value);
        var rowIdx = int.Parse(tcMatch.Groups[3].Value);
        var cellIdx = int.Parse(tcMatch.Groups[4].Value);

        if (!string.IsNullOrEmpty(targetParentPath))
        {
            var expected = $"/slide[{slideIdx}]/table[{tableIdx}]/tr[{rowIdx}]";
            if (!string.Equals(targetParentPath, expected, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    $"Cross-row/cross-table cell copy is not supported. Source cell's row is {expected}; target was {targetParentPath}.");
        }

        var (slidePart, table) = ResolveTable(slideIdx, tableIdx);
        var rows = table.Elements<Drawing.TableRow>().ToList();
        if (rowIdx < 1 || rowIdx > rows.Count)
            throw new ArgumentException($"Row {rowIdx} not found (total: {rows.Count})");
        var row = rows[rowIdx - 1];
        var cells = row.Elements<Drawing.TableCell>().ToList();
        if (cellIdx < 1 || cellIdx > cells.Count)
            throw new ArgumentException($"Cell {cellIdx} not found (total: {cells.Count})");

        var clone = (Drawing.TableCell)cells[cellIdx - 1].CloneNode(true);

        int? targetIdx = null;
        if (position?.After != null || position?.Before != null)
        {
            var anchorPath = ResolveIdPath(position.After ?? position.Before!);
            var anchorMatch = Regex.Match(anchorPath, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]/tc\[(\d+)\]$");
            if (!anchorMatch.Success ||
                int.Parse(anchorMatch.Groups[1].Value) != slideIdx ||
                int.Parse(anchorMatch.Groups[2].Value) != tableIdx ||
                int.Parse(anchorMatch.Groups[3].Value) != rowIdx)
            {
                throw new ArgumentException(
                    $"Copy cell anchor must be a cell in the same row: /slide[{slideIdx}]/table[{tableIdx}]/tr[{rowIdx}]/tc[N]. Got: {anchorPath}");
            }
            var anchorCellIdx = int.Parse(anchorMatch.Groups[4].Value);
            targetIdx = position.After != null ? anchorCellIdx : anchorCellIdx - 1; // 0-based
        }
        else if (position?.Index.HasValue == true)
        {
            targetIdx = position.Index.Value;
        }

        var siblings = row.Elements<Drawing.TableCell>().ToList();
        if (targetIdx.HasValue && targetIdx.Value >= 0 && targetIdx.Value < siblings.Count)
            siblings[targetIdx.Value].InsertBeforeSelf(clone);
        else
            row.AppendChild(clone);

        GetSlide(slidePart).Save();
        var newCells = row.Elements<Drawing.TableCell>().ToList();
        var newCellIdx = newCells.IndexOf(clone) + 1;
        return $"/slide[{slideIdx}]/table[{tableIdx}]/tr[{rowIdx}]/tc[{newCellIdx}]";
    }

    /// <summary>
    /// Resolve a column-anchor path against the same table. Returns the
    /// requested 0-based target column index (insertion slot in gridCol /
    /// per-row tc lists), or null if no anchor or anchor was self-referential.
    /// </summary>
    private int? ResolveColumnAnchorIndex(InsertPosition? position, int slideIdx, int tableIdx, int? sourceColIdx)
    {
        if (position?.After == null && position?.Before == null)
        {
            return position?.Index;
        }
        var anchorPath = ResolveIdPath(position.After ?? position.Before!);
        var anchorMatch = Regex.Match(anchorPath, @"^/slide\[(\d+)\]/table\[(\d+)\]/col\[(\d+)\]$");
        if (!anchorMatch.Success ||
            int.Parse(anchorMatch.Groups[1].Value) != slideIdx ||
            int.Parse(anchorMatch.Groups[2].Value) != tableIdx)
        {
            throw new ArgumentException(
                $"Column anchor must be a column in the same table: /slide[{slideIdx}]/table[{tableIdx}]/col[N]. Got: {anchorPath}");
        }
        var anchorColIdx = int.Parse(anchorMatch.Groups[3].Value); // 1-based
        if (sourceColIdx.HasValue && anchorColIdx == sourceColIdx.Value)
            return -1; // self-anchor sentinel
        var target = position.After != null ? anchorColIdx : anchorColIdx - 1; // 0-based
        // Compensate when removing the source shifts later siblings left
        if (sourceColIdx.HasValue && sourceColIdx.Value < anchorColIdx) target -= 1;
        return target;
    }

    /// <summary>
    /// Move a table column within its table by --before/--after/--index. Same
    /// table only — cross-table moves are ambiguous (grid widths differ).
    /// Mirrors MoveTableRow's compensation logic for delete-then-insert order.
    /// </summary>
    private string MoveTableColumn(Match colMatch, InsertPosition? position, string? targetParentPath)
    {
        var slideIdx = int.Parse(colMatch.Groups[1].Value);
        var tableIdx = int.Parse(colMatch.Groups[2].Value);
        var colIdx = int.Parse(colMatch.Groups[3].Value);

        if (!string.IsNullOrEmpty(targetParentPath))
        {
            var expected = $"/slide[{slideIdx}]/table[{tableIdx}]";
            if (!string.Equals(targetParentPath, expected, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    $"Cross-table column move is not supported. Source column's table is {expected}; target was {targetParentPath}.");
        }

        var (slidePart, table) = ResolveTable(slideIdx, tableIdx);
        var grid = table.GetFirstChild<Drawing.TableGrid>()
            ?? throw new InvalidOperationException("Table has no grid");
        var gridCols = grid.Elements<Drawing.GridColumn>().ToList();
        if (colIdx < 1 || colIdx > gridCols.Count)
            throw new ArgumentException($"Column {colIdx} not found (total: {gridCols.Count})");

        var targetIdx = ResolveColumnAnchorIndex(position, slideIdx, tableIdx, colIdx);
        if (targetIdx == -1) // self-anchor
            return $"/slide[{slideIdx}]/table[{tableIdx}]/col[{colIdx}]";

        // Detach gridCol + per-row tc
        var movingGridCol = gridCols[colIdx - 1];
        movingGridCol.Remove();
        var movingCells = new List<Drawing.TableCell>();
        foreach (var row in table.Elements<Drawing.TableRow>())
        {
            var cells = row.Elements<Drawing.TableCell>().ToList();
            if (colIdx <= cells.Count)
            {
                movingCells.Add(cells[colIdx - 1]);
                cells[colIdx - 1].Remove();
            }
            else
            {
                movingCells.Add(new Drawing.TableCell()); // pad if asymmetric
            }
        }

        // Insert gridCol at targetIdx
        var remainingGridCols = grid.Elements<Drawing.GridColumn>().ToList();
        if (targetIdx.HasValue && targetIdx.Value >= 0 && targetIdx.Value < remainingGridCols.Count)
            remainingGridCols[targetIdx.Value].InsertBeforeSelf(movingGridCol);
        else
            grid.AppendChild(movingGridCol);

        // Insert tc into each row at the same position
        int rowIdx2 = 0;
        foreach (var row in table.Elements<Drawing.TableRow>())
        {
            var rowCells = row.Elements<Drawing.TableCell>().ToList();
            var movingCell = movingCells[rowIdx2++];
            if (targetIdx.HasValue && targetIdx.Value >= 0 && targetIdx.Value < rowCells.Count)
                rowCells[targetIdx.Value].InsertBeforeSelf(movingCell);
            else
                row.AppendChild(movingCell);
        }

        GetSlide(slidePart).Save();
        var newGridCols = grid.Elements<Drawing.GridColumn>().ToList();
        var newColIdx = newGridCols.IndexOf(movingGridCol) + 1;
        return $"/slide[{slideIdx}]/table[{tableIdx}]/col[{newColIdx}]";
    }

    /// <summary>
    /// Clone a table column (gridCol + per-row tc) inside the same table.
    /// </summary>
    private string CopyTableColumn(Match colMatch, InsertPosition? position, string? targetParentPath)
    {
        var slideIdx = int.Parse(colMatch.Groups[1].Value);
        var tableIdx = int.Parse(colMatch.Groups[2].Value);
        var colIdx = int.Parse(colMatch.Groups[3].Value);

        if (!string.IsNullOrEmpty(targetParentPath))
        {
            var expected = $"/slide[{slideIdx}]/table[{tableIdx}]";
            if (!string.Equals(targetParentPath, expected, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    $"Cross-table column copy is not supported. Source column's table is {expected}; target was {targetParentPath}.");
        }

        var (slidePart, table) = ResolveTable(slideIdx, tableIdx);
        var grid = table.GetFirstChild<Drawing.TableGrid>()
            ?? throw new InvalidOperationException("Table has no grid");
        var gridCols = grid.Elements<Drawing.GridColumn>().ToList();
        if (colIdx < 1 || colIdx > gridCols.Count)
            throw new ArgumentException($"Column {colIdx} not found (total: {gridCols.Count})");

        // No source removal here, so don't pass sourceColIdx (no compensation needed).
        var targetIdx = ResolveColumnAnchorIndex(position, slideIdx, tableIdx, sourceColIdx: null);

        var clonedGridCol = (Drawing.GridColumn)gridCols[colIdx - 1].CloneNode(true);
        var clonedCells = new List<Drawing.TableCell>();
        foreach (var row in table.Elements<Drawing.TableRow>())
        {
            var cells = row.Elements<Drawing.TableCell>().ToList();
            clonedCells.Add(colIdx <= cells.Count
                ? (Drawing.TableCell)cells[colIdx - 1].CloneNode(true)
                : new Drawing.TableCell());
        }

        var siblingsGrid = grid.Elements<Drawing.GridColumn>().ToList();
        if (targetIdx.HasValue && targetIdx.Value >= 0 && targetIdx.Value < siblingsGrid.Count)
            siblingsGrid[targetIdx.Value].InsertBeforeSelf(clonedGridCol);
        else
            grid.AppendChild(clonedGridCol);

        int rowIdx2 = 0;
        foreach (var row in table.Elements<Drawing.TableRow>())
        {
            var rowCells = row.Elements<Drawing.TableCell>().ToList();
            var clone = clonedCells[rowIdx2++];
            if (targetIdx.HasValue && targetIdx.Value >= 0 && targetIdx.Value < rowCells.Count)
                rowCells[targetIdx.Value].InsertBeforeSelf(clone);
            else
                row.AppendChild(clone);
        }

        // Update GraphicFrame container width to match new total grid width
        var graphicFrame = table.Ancestors<GraphicFrame>().FirstOrDefault();
        if (graphicFrame?.Transform?.Extents != null)
        {
            long totalColWidth = grid.Elements<Drawing.GridColumn>()
                .Sum(gc => gc.Width?.Value ?? 914400);
            graphicFrame.Transform.Extents.Cx = totalColWidth;
        }

        GetSlide(slidePart).Save();
        var newGridCols = grid.Elements<Drawing.GridColumn>().ToList();
        var newColIdx = newGridCols.IndexOf(clonedGridCol) + 1;
        return $"/slide[{slideIdx}]/table[{tableIdx}]/col[{newColIdx}]";
    }

    /// <summary>
    /// Clone an entire slide with all its content, relationships (images, charts, media),
    /// layout link, background, notes, and transitions.
    /// Pattern follows POI's createSlide(layout) + importContent(srcSlide).
    /// </summary>
    private string CloneSlide(Match slideMatch, List<SlidePart> slideParts, int? index)
    {
        var srcSlideIdx = int.Parse(slideMatch.Groups[1].Value);
        if (srcSlideIdx < 1 || srcSlideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {srcSlideIdx} not found (total: {slideParts.Count})");

        var srcSlidePart = slideParts[srcSlideIdx - 1];
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation not found");
        var presentation = presentationPart.Presentation
            ?? throw new InvalidOperationException("No presentation");

        // 1. Create new SlidePart
        var newSlidePart = presentationPart.AddNewPart<SlidePart>();

        // 2. Copy slide layout relationship (link to same layout as source)
        var srcLayoutPart = srcSlidePart.SlideLayoutPart;
        if (srcLayoutPart != null)
            newSlidePart.AddPart(srcLayoutPart);

        // 3. Deep-clone the Slide XML
        var srcSlide = GetSlide(srcSlidePart);
        newSlidePart.Slide = (Slide)srcSlide.CloneNode(true);

        // 4. Copy all referenced parts (images, charts, embedded objects, media)
        CopySlideParts(srcSlidePart, newSlidePart);

        // 5. Copy notes slide if present
        if (srcSlidePart.NotesSlidePart != null)
        {
            var srcNotesPart = srcSlidePart.NotesSlidePart;
            var newNotesPart = newSlidePart.AddNewPart<NotesSlidePart>();
            newNotesPart.NotesSlide = srcNotesPart.NotesSlide != null
                ? (NotesSlide)srcNotesPart.NotesSlide.CloneNode(true)
                : new NotesSlide();
            // Link notes to the new slide
            newNotesPart.AddPart(newSlidePart);
        }

        newSlidePart.Slide.Save();

        // 6. Register in SlideIdList at the correct position
        var slideIdList = presentation.GetFirstChild<SlideIdList>()
            ?? presentation.AppendChild(new SlideIdList());
        var maxId = slideIdList.Elements<SlideId>().Any()
            ? slideIdList.Elements<SlideId>().Max(s => s.Id?.Value ?? 255) + 1
            : 256;
        var relId = presentationPart.GetIdOfPart(newSlidePart);
        var newSlideId = new SlideId { Id = maxId, RelationshipId = relId };

        if (index.HasValue && index.Value < slideIdList.Elements<SlideId>().Count())
        {
            var refSlide = slideIdList.Elements<SlideId>().ElementAtOrDefault(index.Value);
            if (refSlide != null)
                slideIdList.InsertBefore(newSlideId, refSlide);
            else
                slideIdList.AppendChild(newSlideId);
        }
        else
        {
            slideIdList.AppendChild(newSlideId);
        }

        presentation.Save();

        var slideIds = slideIdList.Elements<SlideId>().ToList();
        var insertedIdx = slideIds.FindIndex(s => s.RelationshipId?.Value == relId) + 1;
        return $"/slide[{insertedIdx}]";
    }

    /// <summary>
    /// Copy all sub-parts (images, charts, media, etc.) from source to target slide,
    /// remapping relationship IDs in the cloned XML.
    /// </summary>
    private static void CopySlideParts(SlidePart source, SlidePart target)
    {
        // Build a map of old rId → new rId for all parts that need copying
        var rIdMap = new Dictionary<string, string>();

        foreach (var part in source.Parts)
        {
            // Skip SlideLayoutPart (already linked above)
            if (part.OpenXmlPart is SlideLayoutPart) continue;
            // Skip NotesSlidePart (handled separately)
            if (part.OpenXmlPart is NotesSlidePart) continue;

            try
            {
                // Try to add the same part (shares the underlying data)
                var newRelId = target.CreateRelationshipToPart(part.OpenXmlPart);
                if (newRelId != part.RelationshipId)
                    rIdMap[part.RelationshipId] = newRelId;
            }
            catch
            {
                // If sharing fails, deep-copy the part data
                try
                {
                    var newPart = target.AddNewPart<OpenXmlPart>(part.OpenXmlPart.ContentType, part.RelationshipId);
                    using var stream = part.OpenXmlPart.GetStream();
                    newPart.FeedData(stream);
                }
                catch { /* Best effort — some parts may not be copyable */ }
            }
        }

        // Also copy external relationships (hyperlinks, media links)
        foreach (var extRel in source.ExternalRelationships)
        {
            try
            {
                target.AddExternalRelationship(extRel.RelationshipType, extRel.Uri, extRel.Id);
            }
            catch { }
        }
        foreach (var hyperRel in source.HyperlinkRelationships)
        {
            try
            {
                target.AddHyperlinkRelationship(hyperRel.Uri, hyperRel.IsExternal, hyperRel.Id);
            }
            catch { }
        }

        // Remap any changed relationship IDs in the slide XML
        if (rIdMap.Count > 0 && target.Slide != null)
        {
            RemapRelationshipIds(target.Slide, rIdMap);
            target.Slide.Save();
        }
    }

    /// <summary>
    /// Update all r:id references in the XML tree when relationship IDs changed during copy.
    /// </summary>
    private static void RemapRelationshipIds(OpenXmlElement root, Dictionary<string, string> rIdMap)
    {
        var rNsUri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        foreach (var el in root.Descendants().Prepend(root).ToList())
        {
            foreach (var attr in el.GetAttributes().ToList())
            {
                if (attr.NamespaceUri != rNsUri || attr.Value == null) continue;
                if (rIdMap.TryGetValue(attr.Value, out var newId))
                {
                    el.SetAttribute(new OpenXmlAttribute(attr.Prefix, attr.LocalName, attr.NamespaceUri, newId));
                }
            }
        }
    }

    private (SlidePart slidePart, OpenXmlElement element) ResolveSlideElement(string path, List<SlidePart> slideParts)
    {
        var match = Regex.Match(path, @"^/slide\[(\d+)\]/(\w+)\[(\d+)\]$");
        if (!match.Success)
            throw new ArgumentException($"Invalid element path: {path}. Expected /slide[N]/element[M]");

        var slideIdx = int.Parse(match.Groups[1].Value);
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");

        var elementType = match.Groups[2].Value;
        var elementIdx = int.Parse(match.Groups[3].Value);

        OpenXmlElement element = elementType switch
        {
            "shape" => shapeTree.Elements<Shape>().ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Shape {elementIdx} not found"),
            "picture" or "pic" => shapeTree.Elements<Picture>().ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Picture {elementIdx} not found"),
            "connector" or "connection" => shapeTree.Elements<ConnectionShape>().ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Connector {elementIdx} not found"),
            "table" => shapeTree.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<Drawing.Table>().Any()).ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Table {elementIdx} not found"),
            "chart" => shapeTree.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<C.ChartReference>().Any()).ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Chart {elementIdx} not found"),
            "ole" or "object" or "embed" => shapeTree.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<DocumentFormat.OpenXml.Presentation.OleObject>().Any())
                .ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"OLE object {elementIdx} not found"),
            "group" => shapeTree.Elements<GroupShape>().ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Group {elementIdx} not found"),
            _ => shapeTree.ChildElements
                .Where(e => e.LocalName.Equals(elementType, StringComparison.OrdinalIgnoreCase))
                .ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"{elementType} {elementIdx} not found")
        };

        return (slidePart, element);
    }

    private static void CopyRelationships(OpenXmlElement element, SlidePart sourcePart, SlidePart targetPart)
    {
        var rNsUri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var allElements = element.Descendants().Prepend(element);

        foreach (var el in allElements.ToList())
        {
            foreach (var attr in el.GetAttributes().ToList())
            {
                if (attr.NamespaceUri != rNsUri) continue;

                var oldRelId = attr.Value;
                if (string.IsNullOrEmpty(oldRelId)) continue;

                // Try part-based relationships first
                bool handled = false;
                try
                {
                    var referencedPart = sourcePart.GetPartById(oldRelId);
                    string newRelId;
                    try
                    {
                        newRelId = targetPart.GetIdOfPart(referencedPart);
                    }
                    catch (ArgumentException)
                    {
                        newRelId = targetPart.CreateRelationshipToPart(referencedPart);
                    }

                    if (newRelId != oldRelId)
                    {
                        el.SetAttribute(new OpenXmlAttribute(attr.Prefix, attr.LocalName, attr.NamespaceUri, newRelId));
                    }
                    handled = true;
                }
                catch (ArgumentOutOfRangeException) { /* Not a part-based relationship */ }

                if (!handled)
                {
                    // Try hyperlink relationships (external, not part-based)
                    var hyperlinkRel = sourcePart.HyperlinkRelationships.FirstOrDefault(r => r.Id == oldRelId);
                    if (hyperlinkRel != null)
                    {
                        var existingTarget = targetPart.HyperlinkRelationships.FirstOrDefault(r => r.Uri == hyperlinkRel.Uri);
                        var newHRelId = existingTarget?.Id
                            ?? targetPart.AddHyperlinkRelationship(hyperlinkRel.Uri, hyperlinkRel.IsExternal).Id;
                        if (newHRelId != oldRelId)
                        {
                            el.SetAttribute(new OpenXmlAttribute(attr.Prefix, attr.LocalName, attr.NamespaceUri, newHRelId));
                        }
                    }
                    else
                    {
                        // Try other external relationships
                        var externalRel = sourcePart.ExternalRelationships.FirstOrDefault(r => r.Id == oldRelId);
                        if (externalRel != null)
                        {
                            var existing = targetPart.ExternalRelationships
                                .FirstOrDefault(r => r.Uri == externalRel.Uri && r.RelationshipType == externalRel.RelationshipType);
                            var newERelId = existing?.Id ?? targetPart.AddExternalRelationship(externalRel.RelationshipType, externalRel.Uri).Id;
                            if (newERelId != oldRelId)
                            {
                                el.SetAttribute(new OpenXmlAttribute(attr.Prefix, attr.LocalName, attr.NamespaceUri, newERelId));
                            }
                        }
                    }
                }
            }
        }
    }

    private static void InsertAtPosition(OpenXmlElement parent, OpenXmlElement element, int? index)
    {
        if (index.HasValue && parent is ShapeTree)
        {
            // Skip structural elements (nvGrpSpPr, grpSpPr) that must stay at the beginning
            var contentChildren = parent.ChildElements
                .Where(e => e is not NonVisualGroupShapeProperties && e is not GroupShapeProperties)
                .ToList();
            if (index.Value >= 0 && index.Value < contentChildren.Count)
                contentChildren[index.Value].InsertBeforeSelf(element);
            else if (contentChildren.Count > 0)
                contentChildren.Last().InsertAfterSelf(element);
            else
                parent.AppendChild(element);
        }
        else if (index.HasValue)
        {
            var children = parent.ChildElements.ToList();
            if (index.Value >= 0 && index.Value < children.Count)
                children[index.Value].InsertBeforeSelf(element);
            else
                parent.AppendChild(element);
        }
        else
        {
            parent.AppendChild(element);
        }
    }

    private static string ComputeElementPath(string parentPath, OpenXmlElement element, ShapeTree shapeTree)
    {
        // Map back to semantic type names
        string typeName;
        int typeIdx;
        if (element is Shape)
        {
            typeName = "shape";
            typeIdx = shapeTree.Elements<Shape>().ToList().IndexOf((Shape)element) + 1;
        }
        else if (element is Picture)
        {
            typeName = "picture";
            typeIdx = shapeTree.Elements<Picture>().ToList().IndexOf((Picture)element) + 1;
        }
        else if (element is ConnectionShape)
        {
            typeName = "connector";
            typeIdx = shapeTree.Elements<ConnectionShape>().ToList().IndexOf((ConnectionShape)element) + 1;
        }
        else if (element is GroupShape)
        {
            typeName = "group";
            typeIdx = shapeTree.Elements<GroupShape>().ToList().IndexOf((GroupShape)element) + 1;
        }
        else if (element is GraphicFrame gf)
        {
            if (gf.Descendants<Drawing.Table>().Any())
            {
                typeName = "table";
                typeIdx = shapeTree.Elements<GraphicFrame>()
                    .Where(f => f.Descendants<Drawing.Table>().Any())
                    .ToList().IndexOf(gf) + 1;
            }
            else if (gf.Descendants<C.ChartReference>().Any())
            {
                typeName = "chart";
                typeIdx = shapeTree.Elements<GraphicFrame>()
                    .Where(f => f.Descendants<C.ChartReference>().Any())
                    .ToList().IndexOf(gf) + 1;
            }
            else if (gf.Descendants<DocumentFormat.OpenXml.Presentation.OleObject>().Any())
            {
                typeName = "ole";
                typeIdx = shapeTree.Elements<GraphicFrame>()
                    .Where(f => f.Descendants<DocumentFormat.OpenXml.Presentation.OleObject>().Any())
                    .ToList().IndexOf(gf) + 1;
            }
            else
            {
                typeName = element.LocalName;
                typeIdx = shapeTree.ChildElements
                    .Where(e => e.LocalName == element.LocalName)
                    .ToList().IndexOf(element) + 1;
            }
        }
        else
        {
            typeName = element.LocalName;
            typeIdx = shapeTree.ChildElements
                .Where(e => e.LocalName == element.LocalName)
                .ToList().IndexOf(element) + 1;
        }
        return $"{parentPath}/{BuildElementPathSegment(typeName, element, typeIdx)}";
    }

    // CONSISTENCY(container-remove-guard): hardcoded list of pptx container
    // paths that must never be removed. Mirrors schema entries marked
    // `"container": true` under schemas/help/pptx/*.json (presentation,
    // theme, slidemaster, slidelayout). Removing the backing part of any
    // of these breaks the deck beyond recovery.
    private static readonly HashSet<string> ProtectedPptxContainerPaths = new(StringComparer.OrdinalIgnoreCase)
    {
        "/presentation",
        "/slidemaster",
        "/slidelayout",
        "/theme",
    };

    private static bool IsProtectedPptxContainerPath(string path)
    {
        if (string.IsNullOrEmpty(path)) return false;
        return ProtectedPptxContainerPaths.Contains(path.TrimEnd('/'));
    }
}
