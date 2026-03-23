// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private string AddConnector(string parentPath, int? index, Dictionary<string, string> properties)
    {
                var cxnSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!cxnSlideMatch.Success)
                    throw new ArgumentException("Connectors must be added to a slide: /slide[N]");

                var cxnSlideIdx = int.Parse(cxnSlideMatch.Groups[1].Value);
                var cxnSlideParts = GetSlideParts().ToList();
                if (cxnSlideIdx < 1 || cxnSlideIdx > cxnSlideParts.Count)
                    throw new ArgumentException($"Slide {cxnSlideIdx} not found (total: {cxnSlideParts.Count})");

                var cxnSlidePart = cxnSlideParts[cxnSlideIdx - 1];
                var cxnShapeTree = GetSlide(cxnSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                var cxnId = (uint)(cxnShapeTree.ChildElements.Count + 2);
                var cxnName = properties.GetValueOrDefault("name", $"Connector {cxnId}");

                // Position: x1,y1 → x2,y2 or x,y,width,height
                long cxnX = (properties.TryGetValue("x", out var cx1) || properties.TryGetValue("left", out cx1)) ? ParseEmu(cx1) : 2000000;
                long cxnY = (properties.TryGetValue("y", out var cy1) || properties.TryGetValue("top", out cy1)) ? ParseEmu(cy1) : 3000000;
                long cxnCx = properties.TryGetValue("width", out var cw) ? ParseEmu(cw) : 4000000;
                long cxnCy = properties.TryGetValue("height", out var ch) ? ParseEmu(ch) : 0;

                var connector = new ConnectionShape();
                var cxnNvProps = new NonVisualConnectionShapeProperties(
                    new NonVisualDrawingProperties { Id = cxnId, Name = cxnName },
                    new NonVisualConnectorShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                );

                // Connect to shapes if specified
                var cxnDrawProps = cxnNvProps.NonVisualConnectorShapeDrawingProperties!;
                if (properties.TryGetValue("startshape", out var startId))
                {
                    if (!uint.TryParse(startId, out var startIdVal))
                        throw new ArgumentException($"Invalid 'startshape' value: '{startId}'. Expected a positive integer (shape ID).");
                    cxnDrawProps.StartConnection = new Drawing.StartConnection { Id = startIdVal, Index = 0 };
                }
                if (properties.TryGetValue("endshape", out var endId))
                {
                    if (!uint.TryParse(endId, out var endIdVal))
                        throw new ArgumentException($"Invalid 'endshape' value: '{endId}'. Expected a positive integer (shape ID).");
                    cxnDrawProps.EndConnection = new Drawing.EndConnection { Id = endIdVal, Index = 0 };
                }

                connector.NonVisualConnectionShapeProperties = cxnNvProps;
                connector.ShapeProperties = new ShapeProperties(
                    new Drawing.Transform2D(
                        new Drawing.Offset { X = cxnX, Y = cxnY },
                        new Drawing.Extents { Cx = cxnCx, Cy = cxnCy }
                    ),
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList())
                    {
                        Preset = properties.GetValueOrDefault("preset", "straightConnector1").ToLowerInvariant() switch
                        {
                            "straight" or "straightconnector1" => Drawing.ShapeTypeValues.StraightConnector1,
                            "elbow" or "bentconnector3" => Drawing.ShapeTypeValues.BentConnector3,
                            "curve" or "curvedconnector3" => Drawing.ShapeTypeValues.CurvedConnector3,
                            _ => throw new ArgumentException($"Invalid connector preset: '{properties.GetValueOrDefault("preset", "straightConnector1")}'. Valid values: straight, elbow, curve.")
                        }
                    }
                );

                // Line style
                var cxnOutline = new Drawing.Outline { Width = 12700 }; // 1pt default
                if (properties.TryGetValue("lineColor", out var cxnColor2) || properties.TryGetValue("linecolor", out cxnColor2)
                    || properties.TryGetValue("line", out cxnColor2) || properties.TryGetValue("color", out cxnColor2)
                    || properties.TryGetValue("line.color", out cxnColor2))
                    cxnOutline.AppendChild(BuildSolidFill(cxnColor2));
                else
                    cxnOutline.AppendChild(BuildSolidFill("000000"));
                if (properties.TryGetValue("linewidth", out var lwVal) || properties.TryGetValue("lineWidth", out lwVal)
                    || properties.TryGetValue("line.width", out lwVal))
                    cxnOutline.Width = Core.EmuConverter.ParseLineWidth(lwVal);
                if (properties.TryGetValue("lineDash", out var cxnDash) || properties.TryGetValue("linedash", out cxnDash))
                {
                    cxnOutline.AppendChild(new Drawing.PresetDash
                    {
                        Val = cxnDash.ToLowerInvariant() switch
                        {
                            "solid" => Drawing.PresetLineDashValues.Solid,
                            "dot" => Drawing.PresetLineDashValues.Dot,
                            "dash" => Drawing.PresetLineDashValues.Dash,
                            "dashdot" => Drawing.PresetLineDashValues.DashDot,
                            "longdash" => Drawing.PresetLineDashValues.LargeDash,
                            "longdashdot" => Drawing.PresetLineDashValues.LargeDashDot,
                            "sysdot" => Drawing.PresetLineDashValues.SystemDot,
                            "sysdash" => Drawing.PresetLineDashValues.SystemDash,
                            _ => Drawing.PresetLineDashValues.Solid
                        }
                    });
                }
                // Arrow head/tail
                if (properties.TryGetValue("headEnd", out var headVal) || properties.TryGetValue("headend", out headVal))
                {
                    cxnOutline.AppendChild(new Drawing.HeadEnd { Type = ParseLineEndType(headVal) });
                }
                if (properties.TryGetValue("tailEnd", out var tailVal) || properties.TryGetValue("tailend", out tailVal))
                {
                    cxnOutline.AppendChild(new Drawing.TailEnd { Type = ParseLineEndType(tailVal) });
                }

                if (properties.TryGetValue("rotation", out var cxnRot))
                {
                    if (int.TryParse(cxnRot, out var rotDeg))
                        connector.ShapeProperties.Transform2D!.Rotation = rotDeg * 60000;
                }
                connector.ShapeProperties.AppendChild(cxnOutline);

                cxnShapeTree.AppendChild(connector);
                GetSlide(cxnSlidePart).Save();

                var cxnCount = cxnShapeTree.Elements<ConnectionShape>().Count();
                return $"/slide[{cxnSlideIdx}]/connector[{cxnCount}]";
    }


    private string AddGroup(string parentPath, int? index, Dictionary<string, string> properties)
    {
                var grpSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!grpSlideMatch.Success)
                    throw new ArgumentException("Groups must be added to a slide: /slide[N]");

                var grpSlideIdx = int.Parse(grpSlideMatch.Groups[1].Value);
                var grpSlideParts = GetSlideParts().ToList();
                if (grpSlideIdx < 1 || grpSlideIdx > grpSlideParts.Count)
                    throw new ArgumentException($"Slide {grpSlideIdx} not found (total: {grpSlideParts.Count})");

                var grpSlidePart = grpSlideParts[grpSlideIdx - 1];
                var grpShapeTree = GetSlide(grpSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                var grpId = (uint)(grpShapeTree.ChildElements.Count + 2);
                var grpName = properties.GetValueOrDefault("name", $"Group {grpId}");

                // Parse shape paths to group: shapes="1,2,3" (shape indices)
                if (!properties.TryGetValue("shapes", out var shapesStr))
                    throw new ArgumentException("'shapes' property required: comma-separated shape indices to group (e.g. shapes=1,2,3)");

                var shapeParts = shapesStr.Split(',');
                var shapeIndices = new List<int>();
                foreach (var sp in shapeParts)
                {
                    if (!int.TryParse(sp.Trim(), out var idx))
                        throw new ArgumentException($"Invalid 'shapes' value: '{sp.Trim()}' is not a valid integer. Expected comma-separated shape indices (e.g. shapes=1,2,3).");
                    shapeIndices.Add(idx);
                }
                var allShapes = grpShapeTree.Elements<Shape>().ToList();

                // Collect shapes to group (in reverse order to maintain indices during removal)
                var toGroup = new List<Shape>();
                foreach (var si in shapeIndices.OrderBy(i => i))
                {
                    if (si < 1 || si > allShapes.Count)
                        throw new ArgumentException($"Shape {si} not found (total: {allShapes.Count})");
                    toGroup.Add(allShapes[si - 1]);
                }

                // Calculate bounding box
                long minX = long.MaxValue, minY = long.MaxValue, maxX = long.MinValue, maxY = long.MinValue;
                bool hasTransform = false;
                foreach (var s in toGroup)
                {
                    var xfrm = s.ShapeProperties?.Transform2D;
                    if (xfrm?.Offset == null || xfrm.Extents == null) continue;
                    hasTransform = true;
                    long sx = xfrm.Offset.X ?? 0;
                    long sy = xfrm.Offset.Y ?? 0;
                    long scx = xfrm.Extents.Cx ?? 0;
                    long scy = xfrm.Extents.Cy ?? 0;
                    if (sx < minX) minX = sx;
                    if (sy < minY) minY = sy;
                    if (sx + scx > maxX) maxX = sx + scx;
                    if (sy + scy > maxY) maxY = sy + scy;
                }
                if (!hasTransform) { minX = 0; minY = 0; maxX = 0; maxY = 0; }

                var groupShape = new GroupShape();
                groupShape.NonVisualGroupShapeProperties = new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = grpId, Name = grpName },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                );
                groupShape.GroupShapeProperties = new GroupShapeProperties(
                    new Drawing.TransformGroup(
                        new Drawing.Offset { X = minX, Y = minY },
                        new Drawing.Extents { Cx = maxX - minX, Cy = maxY - minY },
                        new Drawing.ChildOffset { X = minX, Y = minY },
                        new Drawing.ChildExtents { Cx = maxX - minX, Cy = maxY - minY }
                    )
                );

                // Move shapes into group
                foreach (var s in toGroup)
                {
                    s.Remove();
                    groupShape.AppendChild(s);
                }

                grpShapeTree.AppendChild(groupShape);
                GetSlide(grpSlidePart).Save();

                var grpCount = grpShapeTree.Elements<GroupShape>().Count();
                return $"/slide[{grpSlideIdx}]/group[{grpCount}]";
    }


    private string AddAnimation(string parentPath, int? index, Dictionary<string, string> properties)
    {
                // Add animation to a shape: parentPath must be /slide[N]/shape[M]
                var animMatch = System.Text.RegularExpressions.Regex.Match(parentPath, @"^/slide\[(\d+)\]/shape\[(\d+)\]$");
                if (!animMatch.Success)
                    throw new ArgumentException("Animations must be added to a shape: /slide[N]/shape[M]");

                var animSlideIdx = int.Parse(animMatch.Groups[1].Value);
                var animShapeIdx = int.Parse(animMatch.Groups[2].Value);
                var (animSlidePart, animShape) = ResolveShape(animSlideIdx, animShapeIdx);

                // Build animation value string from properties
                var effect = properties.GetValueOrDefault("effect", "fade");
                var cls = properties.GetValueOrDefault("class", "entrance");
                var duration = properties.GetValueOrDefault("duration", "500");
                var trigger = properties.GetValueOrDefault("trigger", "onclick");

                // Map trigger property to animation format
                var triggerPart = trigger.ToLowerInvariant() switch
                {
                    "onclick" or "click" => "click",
                    "after" or "afterprevious" => "after",
                    "with" or "withprevious" => "with",
                    _ => throw new ArgumentException($"Invalid animation trigger: '{trigger}'. Valid values: onclick, click, after, afterprevious, with, withprevious.")
                };

                var animValue = $"{effect}-{cls}-{duration}-{triggerPart}";

                // Append delay/easing properties if specified
                if (properties.TryGetValue("delay", out var delay))
                    animValue += $"-delay={delay}";
                if (properties.TryGetValue("easein", out var easein))
                    animValue += $"-easein={easein}";
                if (properties.TryGetValue("easeout", out var easeout))
                    animValue += $"-easeout={easeout}";
                if (properties.TryGetValue("easing", out var easing))
                    animValue += $"-easing={easing}";
                if (properties.TryGetValue("direction", out var dir))
                    animValue += $"-{dir}";

                ApplyShapeAnimation(animSlidePart, animShape, animValue);
                GetSlide(animSlidePart).Save();

                // Count animations on this shape
                var animShapeId = animShape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0;
                var timing = GetSlide(animSlidePart).GetFirstChild<Timing>();
                var animCount = timing?.Descendants<ShapeTarget>()
                    .Count(st => st.ShapeId?.Value == animShapeId.ToString()) ?? 0;
                return $"{parentPath}/animation[{animCount}]";
    }


    private string AddZoom(string parentPath, int? index, Dictionary<string, string> properties)
    {
                var zmSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!zmSlideMatch.Success)
                    throw new ArgumentException("Zoom must be added to a slide: /slide[N]");

                // Target slide (required)
                if (!properties.TryGetValue("target", out var targetStr) && !properties.TryGetValue("slide", out targetStr))
                    throw new ArgumentException("'target' property required for zoom type (target slide number, e.g. target=2)");
                if (!int.TryParse(targetStr, out var targetSlideNum))
                    throw new ArgumentException($"Invalid 'target' value: '{targetStr}'. Expected a slide number.");

                var zmSlideIdx = int.Parse(zmSlideMatch.Groups[1].Value);
                var zmSlideParts = GetSlideParts().ToList();
                if (zmSlideIdx < 1 || zmSlideIdx > zmSlideParts.Count)
                    throw new ArgumentException($"Slide {zmSlideIdx} not found (total: {zmSlideParts.Count})");
                if (targetSlideNum < 1 || targetSlideNum > zmSlideParts.Count)
                    throw new ArgumentException($"Target slide {targetSlideNum} not found (total: {zmSlideParts.Count})");

                var zmSlidePart = zmSlideParts[zmSlideIdx - 1];
                var zmShapeTree = GetSlide(zmSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");
                var targetSlidePart = zmSlideParts[targetSlideNum - 1];

                // Get target slide's SlideId from presentation.xml
                var zmPresentation = _doc.PresentationPart?.Presentation
                    ?? throw new InvalidOperationException("No presentation");
                var zmSlideIdList = zmPresentation.GetFirstChild<SlideIdList>()
                    ?? throw new InvalidOperationException("No slides");
                var zmSlideIds = zmSlideIdList.Elements<SlideId>().ToList();
                var targetSldId = zmSlideIds[targetSlideNum - 1].Id!.Value;

                // Position and size (default: 8cm x 4.5cm, centered)
                long zmCx = 3048000; // ~8cm
                long zmCy = 1714500; // ~4.5cm
                if (properties.TryGetValue("width", out var zmW)) zmCx = ParseEmu(zmW);
                if (properties.TryGetValue("height", out var zmH)) zmCy = ParseEmu(zmH);
                var (zmSlideW, zmSlideH) = GetSlideSize();
                long zmX = (zmSlideW - zmCx) / 2;
                long zmY = (zmSlideH - zmCy) / 2;
                if (properties.TryGetValue("x", out var zmXStr)) zmX = ParseEmu(zmXStr);
                if (properties.TryGetValue("y", out var zmYStr)) zmY = ParseEmu(zmYStr);

                var returnToParent = properties.TryGetValue("returntoparent", out var rtp) && IsTruthy(rtp) ? "1" : "0";
                var transitionDur = properties.GetValueOrDefault("transitiondur", "1000");

                // Generate shape IDs
                var zmShapeId = (uint)(zmShapeTree.ChildElements.Count + 2);
                var zmName = properties.GetValueOrDefault("name", $"Slide Zoom {zmShapeId}");
                var zmGuid = Guid.NewGuid().ToString("B").ToUpperInvariant();
                var zmCreationId = Guid.NewGuid().ToString("B").ToUpperInvariant();

                // Create a minimal 1x1 gray placeholder PNG (PowerPoint regenerates the thumbnail on open)
                byte[] placeholderPng = GenerateZoomPlaceholderPng();
                var zmImagePart = zmSlidePart.AddImagePart(ImagePartType.Png);
                using (var ms = new MemoryStream(placeholderPng))
                    zmImagePart.FeedData(ms);
                var zmImageRelId = zmSlidePart.GetIdOfPart(zmImagePart);

                // Create slide-to-slide relationship for fallback hyperlink
                var zmSlideRelId = zmSlidePart.CreateRelationshipToPart(targetSlidePart);

                // Build mc:AlternateContent programmatically (same pattern as morph transition)
                var mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
                var pNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
                var aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
                var rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                var pslzNs = "http://schemas.microsoft.com/office/powerpoint/2016/slidezoom";
                var p166Ns = "http://schemas.microsoft.com/office/powerpoint/2016/6/main";
                var a16Ns = "http://schemas.microsoft.com/office/drawing/2014/main";

                var acElement = new OpenXmlUnknownElement("mc", "AlternateContent", mcNs);

                // === mc:Choice (for clients that support Slide Zoom) ===
                var choiceElement = new OpenXmlUnknownElement("mc", "Choice", mcNs);
                choiceElement.SetAttribute(new OpenXmlAttribute("", "Requires", null!, "pslz"));
                choiceElement.AddNamespaceDeclaration("pslz", pslzNs);

                var gfElement = new OpenXmlUnknownElement("p", "graphicFrame", pNs);
                gfElement.AddNamespaceDeclaration("a", aNs);
                gfElement.AddNamespaceDeclaration("r", rNs);

                // nvGraphicFramePr
                var nvGfPr = new OpenXmlUnknownElement("p", "nvGraphicFramePr", pNs);
                var cNvPr = new OpenXmlUnknownElement("p", "cNvPr", pNs);
                cNvPr.SetAttribute(new OpenXmlAttribute("", "id", null!, zmShapeId.ToString()));
                cNvPr.SetAttribute(new OpenXmlAttribute("", "name", null!, zmName));
                // creationId extension
                var extLst = new OpenXmlUnknownElement("a", "extLst", aNs);
                var ext = new OpenXmlUnknownElement("a", "ext", aNs);
                ext.SetAttribute(new OpenXmlAttribute("", "uri", null!, "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"));
                var creationId = new OpenXmlUnknownElement("a16", "creationId", a16Ns);
                creationId.SetAttribute(new OpenXmlAttribute("", "id", null!, zmCreationId));
                ext.AppendChild(creationId);
                extLst.AppendChild(ext);
                cNvPr.AppendChild(extLst);
                nvGfPr.AppendChild(cNvPr);

                var cNvGfSpPr = new OpenXmlUnknownElement("p", "cNvGraphicFramePr", pNs);
                var gfLocks = new OpenXmlUnknownElement("a", "graphicFrameLocks", aNs);
                gfLocks.SetAttribute(new OpenXmlAttribute("", "noChangeAspect", null!, "1"));
                cNvGfSpPr.AppendChild(gfLocks);
                nvGfPr.AppendChild(cNvGfSpPr);
                nvGfPr.AppendChild(new OpenXmlUnknownElement("p", "nvPr", pNs));
                gfElement.AppendChild(nvGfPr);

                // xfrm (position/size)
                var gfXfrm = new OpenXmlUnknownElement("p", "xfrm", pNs);
                var gfOff = new OpenXmlUnknownElement("a", "off", aNs);
                gfOff.SetAttribute(new OpenXmlAttribute("", "x", null!, zmX.ToString()));
                gfOff.SetAttribute(new OpenXmlAttribute("", "y", null!, zmY.ToString()));
                var gfExt = new OpenXmlUnknownElement("a", "ext", aNs);
                gfExt.SetAttribute(new OpenXmlAttribute("", "cx", null!, zmCx.ToString()));
                gfExt.SetAttribute(new OpenXmlAttribute("", "cy", null!, zmCy.ToString()));
                gfXfrm.AppendChild(gfOff);
                gfXfrm.AppendChild(gfExt);
                gfElement.AppendChild(gfXfrm);

                // graphic > graphicData > pslz:sldZm
                var graphic = new OpenXmlUnknownElement("a", "graphic", aNs);
                var graphicData = new OpenXmlUnknownElement("a", "graphicData", aNs);
                graphicData.SetAttribute(new OpenXmlAttribute("", "uri", null!, pslzNs));

                var sldZm = new OpenXmlUnknownElement("pslz", "sldZm", pslzNs);
                var sldZmObj = new OpenXmlUnknownElement("pslz", "sldZmObj", pslzNs);
                sldZmObj.SetAttribute(new OpenXmlAttribute("", "sldId", null!, targetSldId.ToString()));
                sldZmObj.SetAttribute(new OpenXmlAttribute("", "cId", null!, "0"));

                var zmPr = new OpenXmlUnknownElement("pslz", "zmPr", pslzNs);
                zmPr.AddNamespaceDeclaration("p166", p166Ns);
                zmPr.SetAttribute(new OpenXmlAttribute("", "id", null!, zmGuid));
                zmPr.SetAttribute(new OpenXmlAttribute("", "returnToParent", null!, returnToParent));
                zmPr.SetAttribute(new OpenXmlAttribute("", "transitionDur", null!, transitionDur));

                // blipFill (thumbnail)
                var blipFill = new OpenXmlUnknownElement("p166", "blipFill", p166Ns);
                var blip = new OpenXmlUnknownElement("a", "blip", aNs);
                blip.SetAttribute(new OpenXmlAttribute("r", "embed", rNs, zmImageRelId));
                blipFill.AppendChild(blip);
                var stretch = new OpenXmlUnknownElement("a", "stretch", aNs);
                stretch.AppendChild(new OpenXmlUnknownElement("a", "fillRect", aNs));
                blipFill.AppendChild(stretch);
                zmPr.AppendChild(blipFill);

                // spPr (shape properties inside zoom)
                var zmSpPr = new OpenXmlUnknownElement("p166", "spPr", p166Ns);
                var zmSpXfrm = new OpenXmlUnknownElement("a", "xfrm", aNs);
                var zmSpOff = new OpenXmlUnknownElement("a", "off", aNs);
                zmSpOff.SetAttribute(new OpenXmlAttribute("", "x", null!, "0"));
                zmSpOff.SetAttribute(new OpenXmlAttribute("", "y", null!, "0"));
                var zmSpExt = new OpenXmlUnknownElement("a", "ext", aNs);
                zmSpExt.SetAttribute(new OpenXmlAttribute("", "cx", null!, zmCx.ToString()));
                zmSpExt.SetAttribute(new OpenXmlAttribute("", "cy", null!, zmCy.ToString()));
                zmSpXfrm.AppendChild(zmSpOff);
                zmSpXfrm.AppendChild(zmSpExt);
                zmSpPr.AppendChild(zmSpXfrm);
                var prstGeom = new OpenXmlUnknownElement("a", "prstGeom", aNs);
                prstGeom.SetAttribute(new OpenXmlAttribute("", "prst", null!, "rect"));
                prstGeom.AppendChild(new OpenXmlUnknownElement("a", "avLst", aNs));
                zmSpPr.AppendChild(prstGeom);
                var zmLn = new OpenXmlUnknownElement("a", "ln", aNs);
                zmLn.SetAttribute(new OpenXmlAttribute("", "w", null!, "3175"));
                var zmLnFill = new OpenXmlUnknownElement("a", "solidFill", aNs);
                var zmLnClr = new OpenXmlUnknownElement("a", "prstClr", aNs);
                zmLnClr.SetAttribute(new OpenXmlAttribute("", "val", null!, "ltGray"));
                zmLnFill.AppendChild(zmLnClr);
                zmLn.AppendChild(zmLnFill);
                zmSpPr.AppendChild(zmLn);
                zmPr.AppendChild(zmSpPr);

                sldZmObj.AppendChild(zmPr);
                sldZm.AppendChild(sldZmObj);
                graphicData.AppendChild(sldZm);
                graphic.AppendChild(graphicData);
                gfElement.AppendChild(graphic);
                choiceElement.AppendChild(gfElement);

                // === mc:Fallback (pic + hyperlink for older clients) ===
                var fallbackElement = new OpenXmlUnknownElement("mc", "Fallback", mcNs);
                var fbPic = new OpenXmlUnknownElement("p", "pic", pNs);
                fbPic.AddNamespaceDeclaration("a", aNs);
                fbPic.AddNamespaceDeclaration("r", rNs);

                var fbNvPicPr = new OpenXmlUnknownElement("p", "nvPicPr", pNs);
                var fbCNvPr = new OpenXmlUnknownElement("p", "cNvPr", pNs);
                fbCNvPr.SetAttribute(new OpenXmlAttribute("", "id", null!, zmShapeId.ToString()));
                fbCNvPr.SetAttribute(new OpenXmlAttribute("", "name", null!, zmName));
                var hlinkClick = new OpenXmlUnknownElement("a", "hlinkClick", aNs);
                hlinkClick.SetAttribute(new OpenXmlAttribute("r", "id", rNs, zmSlideRelId));
                hlinkClick.SetAttribute(new OpenXmlAttribute("", "action", null!, "ppaction://hlinksldjump"));
                fbCNvPr.AppendChild(hlinkClick);
                // Same creationId
                var fbExtLst = new OpenXmlUnknownElement("a", "extLst", aNs);
                var fbExt = new OpenXmlUnknownElement("a", "ext", aNs);
                fbExt.SetAttribute(new OpenXmlAttribute("", "uri", null!, "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"));
                var fbCreationId = new OpenXmlUnknownElement("a16", "creationId", a16Ns);
                fbCreationId.SetAttribute(new OpenXmlAttribute("", "id", null!, zmCreationId));
                fbExt.AppendChild(fbCreationId);
                fbExtLst.AppendChild(fbExt);
                fbCNvPr.AppendChild(fbExtLst);
                fbNvPicPr.AppendChild(fbCNvPr);

                var fbCNvPicPr = new OpenXmlUnknownElement("p", "cNvPicPr", pNs);
                var picLocks = new OpenXmlUnknownElement("a", "picLocks", aNs);
                foreach (var lockAttr in new[] { "noGrp", "noRot", "noChangeAspect", "noMove", "noResize",
                    "noEditPoints", "noAdjustHandles", "noChangeArrowheads", "noChangeShapeType" })
                    picLocks.SetAttribute(new OpenXmlAttribute("", lockAttr, null!, "1"));
                fbCNvPicPr.AppendChild(picLocks);
                fbNvPicPr.AppendChild(fbCNvPicPr);
                fbNvPicPr.AppendChild(new OpenXmlUnknownElement("p", "nvPr", pNs));
                fbPic.AppendChild(fbNvPicPr);

                // Fallback blipFill
                var fbBlipFill = new OpenXmlUnknownElement("p", "blipFill", pNs);
                var fbBlip = new OpenXmlUnknownElement("a", "blip", aNs);
                fbBlip.SetAttribute(new OpenXmlAttribute("r", "embed", rNs, zmImageRelId));
                fbBlipFill.AppendChild(fbBlip);
                var fbStretch = new OpenXmlUnknownElement("a", "stretch", aNs);
                fbStretch.AppendChild(new OpenXmlUnknownElement("a", "fillRect", aNs));
                fbBlipFill.AppendChild(fbStretch);
                fbPic.AppendChild(fbBlipFill);

                // Fallback spPr
                var fbSpPr = new OpenXmlUnknownElement("p", "spPr", pNs);
                var fbXfrm = new OpenXmlUnknownElement("a", "xfrm", aNs);
                var fbOff = new OpenXmlUnknownElement("a", "off", aNs);
                fbOff.SetAttribute(new OpenXmlAttribute("", "x", null!, zmX.ToString()));
                fbOff.SetAttribute(new OpenXmlAttribute("", "y", null!, zmY.ToString()));
                var fbExtSz = new OpenXmlUnknownElement("a", "ext", aNs);
                fbExtSz.SetAttribute(new OpenXmlAttribute("", "cx", null!, zmCx.ToString()));
                fbExtSz.SetAttribute(new OpenXmlAttribute("", "cy", null!, zmCy.ToString()));
                fbXfrm.AppendChild(fbOff);
                fbXfrm.AppendChild(fbExtSz);
                fbSpPr.AppendChild(fbXfrm);
                var fbGeom = new OpenXmlUnknownElement("a", "prstGeom", aNs);
                fbGeom.SetAttribute(new OpenXmlAttribute("", "prst", null!, "rect"));
                fbGeom.AppendChild(new OpenXmlUnknownElement("a", "avLst", aNs));
                fbSpPr.AppendChild(fbGeom);
                var fbLn = new OpenXmlUnknownElement("a", "ln", aNs);
                fbLn.SetAttribute(new OpenXmlAttribute("", "w", null!, "3175"));
                var fbLnFill = new OpenXmlUnknownElement("a", "solidFill", aNs);
                var fbLnClr = new OpenXmlUnknownElement("a", "prstClr", aNs);
                fbLnClr.SetAttribute(new OpenXmlAttribute("", "val", null!, "ltGray"));
                fbLnFill.AppendChild(fbLnClr);
                fbLn.AppendChild(fbLnFill);
                fbSpPr.AppendChild(fbLn);
                fbPic.AppendChild(fbSpPr);

                fallbackElement.AppendChild(fbPic);

                acElement.AppendChild(choiceElement);
                acElement.AppendChild(fallbackElement);
                zmShapeTree.AppendChild(acElement);
                GetSlide(zmSlidePart).Save();

                var zmCount = zmShapeTree.Elements<OpenXmlUnknownElement>()
                    .Count(e => e.LocalName == "AlternateContent");
                return $"/slide[{zmSlideIdx}]/zoom[{zmCount}]";
    }


    private string AddDefault(string parentPath, int? index, Dictionary<string, string> properties, string type)
    {
                // Try resolving logical paths (table/placeholder) first
                var logicalResult = ResolveLogicalPath(parentPath);
                SlidePart fbSlidePart;
                OpenXmlElement fbParent;

                if (logicalResult.HasValue)
                {
                    fbSlidePart = logicalResult.Value.slidePart;
                    fbParent = logicalResult.Value.element;
                }
                else
                {
                    // Generic fallback: navigate by XML localName
                    var allSegments = GenericXmlQuery.ParsePathSegments(parentPath);
                    if (allSegments.Count == 0 || !allSegments[0].Name.Equals("slide", StringComparison.OrdinalIgnoreCase) || !allSegments[0].Index.HasValue)
                        throw new ArgumentException($"Generic add requires a path starting with /slide[N]: {parentPath}");

                    var fbSlideIdx = allSegments[0].Index!.Value;
                    var fbSlideParts = GetSlideParts().ToList();
                    if (fbSlideIdx < 1 || fbSlideIdx > fbSlideParts.Count)
                        throw new ArgumentException($"Slide {fbSlideIdx} not found (total: {fbSlideParts.Count})");

                    fbSlidePart = fbSlideParts[fbSlideIdx - 1];
                    fbParent = GetSlide(fbSlidePart);
                    var remaining = allSegments.Skip(1).ToList();
                    if (remaining.Count > 0)
                    {
                        fbParent = GenericXmlQuery.NavigateByPath(fbParent, remaining)
                            ?? throw new ArgumentException($"Parent element not found: {parentPath}");
                    }
                }

                var created = GenericXmlQuery.TryCreateTypedElement(fbParent, type, properties, index);
                if (created == null)
                    throw new ArgumentException($"Unknown element type '{type}' for {parentPath}. " +
                        "Valid types: slide, shape, textbox, picture, table, chart, paragraph, run, connector, group, video, audio, equation, notes, zoom. " +
                        "Use 'officecli pptx add' for details.");

                GetSlide(fbSlidePart).Save();

                // Build result path
                var siblings = fbParent.ChildElements.Where(e => e.LocalName == created.LocalName).ToList();
                var createdIdx = siblings.IndexOf(created) + 1;
                return $"{parentPath}/{created.LocalName}[{createdIdx}]";
    }

}
