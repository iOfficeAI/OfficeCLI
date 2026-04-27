// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

// Per-element-type Set helpers for picture / media / OLE / 3D model / zoom paths.
// Mechanically extracted from the original god-method Set(); each helper
// owns one path-pattern's full handling. No behavior change.
public partial class PowerPointHandler
{
    private List<string> SetPictureByPath(Match picSetMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(picSetMatch.Groups[1].Value);
        var picIdx = int.Parse(picSetMatch.Groups[2].Value);

        var slideParts3 = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts3.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts3.Count})");

        var slidePart = slideParts3[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");
        var pics = shapeTree.Elements<Picture>().ToList();
        if (picIdx < 1 || picIdx > pics.Count)
            throw new ArgumentException($"Picture {picIdx} not found (total: {pics.Count})");

        var pic = pics[picIdx - 1];
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "alt":
                    var nvPicPr = pic.NonVisualPictureProperties?.NonVisualDrawingProperties;
                    if (nvPicPr != null) nvPicPr.Description = value;
                    break;
                case "x" or "y" or "width" or "height":
                {
                    var spPr = pic.ShapeProperties ?? (pic.ShapeProperties = new ShapeProperties());
                    var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
                    TryApplyPositionSize(key.ToLowerInvariant(), value,
                        xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset()),
                        xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents()));
                    break;
                }
                case "path" or "src":
                {
                    // Replace image source
                    var blipFill = pic.BlipFill;
                    var blip = blipFill?.GetFirstChild<Drawing.Blip>();
                    if (blip == null) { unsupported.Add(key); break; }
                    var (imgStream, imgType) = OfficeCli.Core.ImageSource.Resolve(value);
                    using var imgStreamDispose2 = imgStream;
                    // Remove old image part(s) to avoid storage bloat,
                    // including the asvg:svgBlip-referenced SVG part
                    // when the previous image was SVG.
                    var oldEmbedId = blip.Embed?.Value;
                    if (oldEmbedId != null)
                    {
                        try { slidePart.DeletePart(oldEmbedId); } catch { }
                    }
                    var oldPicSvgRelId = OfficeCli.Core.SvgImageHelper.GetSvgRelId(blip);
                    if (oldPicSvgRelId != null)
                    {
                        try { slidePart.DeletePart(oldPicSvgRelId); } catch { }
                    }

                    if (imgType == ImagePartType.Svg)
                    {
                        using var newSvgBuf = new MemoryStream();
                        imgStream.CopyTo(newSvgBuf);
                        newSvgBuf.Position = 0;
                        var newSvgPart = slidePart.AddImagePart(ImagePartType.Svg);
                        newSvgPart.FeedData(newSvgBuf);
                        var newPicSvgRelId = slidePart.GetIdOfPart(newSvgPart);

                        var pngFb = slidePart.AddImagePart(ImagePartType.Png);
                        pngFb.FeedData(new MemoryStream(
                            OfficeCli.Core.SvgImageHelper.TransparentPng1x1, writable: false));
                        blip.Embed = slidePart.GetIdOfPart(pngFb);
                        OfficeCli.Core.SvgImageHelper.AppendSvgExtension(blip, newPicSvgRelId);
                    }
                    else
                    {
                        var newImgPart = slidePart.AddImagePart(imgType);
                        newImgPart.FeedData(imgStream);
                        blip.Embed = slidePart.GetIdOfPart(newImgPart);
                        if (oldPicSvgRelId != null)
                        {
                            var extLst = blip.GetFirstChild<Drawing.BlipExtensionList>();
                            if (extLst != null)
                            {
                                foreach (var ext in extLst.Elements<Drawing.BlipExtension>().ToList())
                                {
                                    if (string.Equals(ext.Uri?.Value,
                                        OfficeCli.Core.SvgImageHelper.SvgExtensionUri,
                                        StringComparison.OrdinalIgnoreCase))
                                        ext.Remove();
                                }
                                if (!extLst.Elements<Drawing.BlipExtension>().Any())
                                    extLst.Remove();
                            }
                        }
                    }
                    break;
                }
                case "rotation" or "rotate":
                {
                    var spPr = pic.ShapeProperties ?? (pic.ShapeProperties = new ShapeProperties());
                    var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
                    xfrm.Rotation = (int)(ParseHelpers.SafeParseDouble(value, "rotation") * 60000);
                    break;
                }
                case "crop" or "cropleft" or "cropright" or "croptop" or "cropbottom":
                {
                    // R10: tolerate trailing '%' on crop values — error message
                    // already says "Expected a percentage (0-100)", so the % literal
                    // is the natural input form.
                    static string StripPct(string s)
                    {
                        var t = s.Trim();
                        return t.EndsWith("%", StringComparison.Ordinal) ? t[..^1].Trim() : t;
                    }
                    var blipFill = pic.BlipFill;
                    if (blipFill == null) { unsupported.Add(key); break; }
                    var srcRect = blipFill.GetFirstChild<Drawing.SourceRectangle>();
                    if (srcRect == null)
                    {
                        srcRect = new Drawing.SourceRectangle();
                        // CONSISTENCY(ooxml-element-order): in CT_BlipFillProperties
                        // srcRect must precede the fill-mode element (stretch/tile).
                        // PowerPoint silently ignores out-of-order srcRect.
                        var fillMode = (OpenXmlElement?)blipFill.GetFirstChild<Drawing.Stretch>()
                            ?? blipFill.GetFirstChild<Drawing.Tile>();
                        if (fillMode != null)
                            blipFill.InsertBefore(srcRect, fillMode);
                        else
                            blipFill.AppendChild(srcRect);
                    }

                    if (key.Equals("crop", StringComparison.OrdinalIgnoreCase))
                    {
                        // Single value: "left,top,right,bottom" as percentages (0-100)
                        var parts = value.Split(',');
                        if (parts.Length == 4)
                        {
                            var cropVals = new double[4];
                            for (int ci = 0; ci < 4; ci++)
                            {
                                cropVals[ci] = ParseHelpers.SafeParseDouble(StripPct(parts[ci]), "crop");
                                if (cropVals[ci] < 0 || cropVals[ci] > 100)
                                    throw new ArgumentException($"Invalid 'crop' value: '{parts[ci].Trim()}'. Crop percentage must be between 0 and 100.");
                            }
                            srcRect.Left = (int)(cropVals[0] * 1000);
                            srcRect.Top = (int)(cropVals[1] * 1000);
                            srcRect.Right = (int)(cropVals[2] * 1000);
                            srcRect.Bottom = (int)(cropVals[3] * 1000);
                        }
                        else if (parts.Length == 2)
                        {
                            // 2-value: vertical,horizontal (top/bottom, left/right)
                            var vCrop = ParseHelpers.SafeParseDouble(StripPct(parts[0]), "crop");
                            var hCrop = ParseHelpers.SafeParseDouble(StripPct(parts[1]), "crop");
                            if (vCrop < 0 || vCrop > 100 || hCrop < 0 || hCrop > 100)
                                throw new ArgumentException($"Invalid 'crop' value: '{value}'. Crop percentages must be between 0 and 100.");
                            srcRect.Top = (int)(vCrop * 1000); srcRect.Bottom = (int)(vCrop * 1000);
                            srcRect.Left = (int)(hCrop * 1000); srcRect.Right = (int)(hCrop * 1000);
                        }
                        else if (parts.Length == 1)
                        {
                            if (!double.TryParse(StripPct(value), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var cropVal))
                                throw new ArgumentException($"Invalid 'crop' value: '{value}'. Expected a percentage (e.g. 10 = 10% from each edge).");
                            if (cropVal < 0 || cropVal > 100)
                                throw new ArgumentException($"Invalid 'crop' value: '{value}'. Crop percentage must be between 0 and 100.");
                            var cropPct = (int)(cropVal * 1000);
                            srcRect.Left = cropPct; srcRect.Top = cropPct; srcRect.Right = cropPct; srcRect.Bottom = cropPct;
                        }
                        else
                        {
                            throw new ArgumentException($"Invalid 'crop' value: '{value}'. Expected 1 value (symmetric), 2 values (vertical,horizontal), or 4 values (left,top,right,bottom).");
                        }
                    }
                    else
                    {
                        if (!double.TryParse(StripPct(value), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var cropSingle))
                            throw new ArgumentException($"Invalid '{key}' value: '{value}'. Expected a percentage (0-100).");
                        if (cropSingle < 0 || cropSingle > 100)
                            throw new ArgumentException($"Invalid '{key}' value: '{value}'. Crop percentage must be between 0 and 100.");
                        var pct = (int)(cropSingle * 1000); // percent (0-100) → 1/1000ths
                        switch (key.ToLowerInvariant())
                        {
                            case "cropleft": srcRect.Left = pct; break;
                            case "croptop": srcRect.Top = pct; break;
                            case "cropright": srcRect.Right = pct; break;
                            case "cropbottom": srcRect.Bottom = pct; break;
                        }
                    }
                    // Reset semantics: if all four sides are zero (or unset),
                    // drop the srcRect entirely so the XML is clean.
                    int L = srcRect.Left?.Value ?? 0;
                    int T = srcRect.Top?.Value ?? 0;
                    int R = srcRect.Right?.Value ?? 0;
                    int B = srcRect.Bottom?.Value ?? 0;
                    if (L == 0 && T == 0 && R == 0 && B == 0)
                        srcRect.Remove();
                    break;
                }
                case "opacity":
                {
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var opacityVal)
                        || double.IsNaN(opacityVal) || double.IsInfinity(opacityVal))
                        throw new ArgumentException($"Invalid 'opacity' value: '{value}'. Expected a finite decimal 0.0-1.0.");
                    if (opacityVal > 1.0) opacityVal /= 100.0;
                    var blip = pic.BlipFill?.GetFirstChild<Drawing.Blip>();
                    if (blip != null)
                    {
                        blip.RemoveAllChildren<Drawing.AlphaModulationFixed>();
                        var alphaVal = (int)(opacityVal * 100000); // 0.0-1.0 → 0-100000
                        blip.AppendChild(new Drawing.AlphaModulationFixed { Amount = alphaVal });
                    }
                    break;
                }
                case "name":
                {
                    var nvPr = pic.NonVisualPictureProperties?.NonVisualDrawingProperties;
                    if (nvPr != null) nvPr.Name = value;
                    break;
                }
                default:
                    if (unsupported.Count == 0)
                        unsupported.Add($"{key} (valid picture props: path, src, x, y, width, height, rotation, opacity, name, crop, cropleft, croptop, cropright, cropbottom)");
                    else
                        unsupported.Add(key);
                    break;
            }
        }
        GetSlide(slidePart).Save();
        return unsupported;
    }

    private List<string> SetZoomByPath(Match zoomSetMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(zoomSetMatch.Groups[1].Value);
        var zmIdx = int.Parse(zoomSetMatch.Groups[2].Value);
        var zmSlideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > zmSlideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {zmSlideParts.Count})");
        var zmSlidePart = zmSlideParts[slideIdx - 1];
        var zmShapeTree = GetSlide(zmSlidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");
        var zoomElements = GetZoomElements(zmShapeTree);
        if (zmIdx < 1 || zmIdx > zoomElements.Count)
            throw new ArgumentException($"Zoom {zmIdx} not found (total: {zoomElements.Count})");

        var acElement = zoomElements[zmIdx - 1];
        var choice = acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Choice");
        var fallback = acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Fallback");
        var gf = choice?.ChildElements.FirstOrDefault(e => e.LocalName == "graphicFrame");
        var sldZmObj = acElement.Descendants().FirstOrDefault(d => d.LocalName == "sldZmObj");
        var zmPr = acElement.Descendants().FirstOrDefault(d => d.LocalName == "zmPr");

        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "target" or "slide":
                {
                    if (!int.TryParse(value, out var targetNum))
                        throw new ArgumentException($"Invalid target value: '{value}'. Expected a slide number.");
                    if (targetNum < 1 || targetNum > zmSlideParts.Count)
                        throw new ArgumentException($"Target slide {targetNum} not found (total: {zmSlideParts.Count})");
                    var zmPresentation = _doc.PresentationPart?.Presentation
                        ?? throw new InvalidOperationException("No presentation");
                    var zmSlideIds = zmPresentation.GetFirstChild<SlideIdList>()
                        ?.Elements<SlideId>().ToList()
                        ?? throw new InvalidOperationException("No slides");
                    var newSldId = zmSlideIds[targetNum - 1].Id!.Value;
                    sldZmObj?.SetAttribute(new OpenXmlAttribute("", "sldId", null!, newSldId.ToString()));

                    // Update fallback hyperlink relationship
                    var fbPic = fallback?.ChildElements.FirstOrDefault(e => e.LocalName == "pic");
                    var fbCNvPr = fbPic?.Descendants().FirstOrDefault(d => d.LocalName == "cNvPr");
                    var hlinkClick = fbCNvPr?.ChildElements.FirstOrDefault(e => e.LocalName == "hlinkClick");
                    if (hlinkClick != null)
                    {
                        var rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                        var targetSlidePart = zmSlideParts[targetNum - 1];
                        var newRelId = zmSlidePart.CreateRelationshipToPart(targetSlidePart);
                        hlinkClick.SetAttribute(new OpenXmlAttribute("r", "id", rNs, newRelId));
                    }
                    break;
                }
                case "returntoparent":
                    zmPr?.SetAttribute(new OpenXmlAttribute("", "returnToParent", null!, IsTruthy(value) ? "1" : "0"));
                    break;
                case "transitiondur":
                    zmPr?.SetAttribute(new OpenXmlAttribute("", "transitionDur", null!, value));
                    break;
                case "x" or "y" or "width" or "height":
                {
                    var emu = ParseEmu(value);
                    // Update graphicFrame xfrm
                    var gfXfrm = gf?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
                    if (gfXfrm != null)
                    {
                        if (key.ToLowerInvariant() is "x" or "y")
                        {
                            var off = gfXfrm.ChildElements.FirstOrDefault(e => e.LocalName == "off");
                            off?.SetAttribute(new OpenXmlAttribute("", key.ToLowerInvariant(), null!, emu.ToString()));
                        }
                        else
                        {
                            var ext = gfXfrm.ChildElements.FirstOrDefault(e => e.LocalName == "ext");
                            var attrName = key.ToLowerInvariant() == "width" ? "cx" : "cy";
                            ext?.SetAttribute(new OpenXmlAttribute("", attrName, null!, emu.ToString()));
                        }
                    }
                    // Update fallback spPr xfrm
                    var fbPic = fallback?.ChildElements.FirstOrDefault(e => e.LocalName == "pic");
                    var fbSpPr = fbPic?.ChildElements.FirstOrDefault(e => e.LocalName == "spPr");
                    var fbXfrm = fbSpPr?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
                    if (fbXfrm != null)
                    {
                        if (key.ToLowerInvariant() is "x" or "y")
                        {
                            var off = fbXfrm.ChildElements.FirstOrDefault(e => e.LocalName == "off");
                            off?.SetAttribute(new OpenXmlAttribute("", key.ToLowerInvariant(), null!, emu.ToString()));
                        }
                        else
                        {
                            var ext = fbXfrm.ChildElements.FirstOrDefault(e => e.LocalName == "ext");
                            var attrName = key.ToLowerInvariant() == "width" ? "cx" : "cy";
                            ext?.SetAttribute(new OpenXmlAttribute("", attrName, null!, emu.ToString()));
                        }
                    }
                    // Update inner zmPr > spPr > xfrm (only for width/height)
                    if (key.ToLowerInvariant() is "width" or "height")
                    {
                        var p166Ns = "http://schemas.microsoft.com/office/powerpoint/2016/6/main";
                        var zmSpPr = zmPr?.ChildElements.FirstOrDefault(e => e.LocalName == "spPr" && e.NamespaceUri == p166Ns);
                        var zmSpXfrm = zmSpPr?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
                        var zmSpExt = zmSpXfrm?.ChildElements.FirstOrDefault(e => e.LocalName == "ext");
                        var attrName = key.ToLowerInvariant() == "width" ? "cx" : "cy";
                        zmSpExt?.SetAttribute(new OpenXmlAttribute("", attrName, null!, emu.ToString()));
                    }
                    break;
                }
                case "name":
                {
                    // Update cNvPr name in Choice
                    var nvGfPr = gf?.ChildElements.FirstOrDefault(e => e.LocalName == "nvGraphicFramePr");
                    var choiceCNvPr = nvGfPr?.ChildElements.FirstOrDefault(e => e.LocalName == "cNvPr");
                    choiceCNvPr?.SetAttribute(new OpenXmlAttribute("", "name", null!, value));
                    // Update cNvPr name in Fallback
                    var fbPic = fallback?.ChildElements.FirstOrDefault(e => e.LocalName == "pic");
                    var fbNvPicPr = fbPic?.ChildElements.FirstOrDefault(e => e.LocalName == "nvPicPr");
                    var fbCNvPr = fbNvPicPr?.ChildElements.FirstOrDefault(e => e.LocalName == "cNvPr");
                    fbCNvPr?.SetAttribute(new OpenXmlAttribute("", "name", null!, value));
                    break;
                }
                case "image" or "path" or "src" or "cover":
                {
                    var (zmImgStream, zmImgPartType) = OfficeCli.Core.ImageSource.Resolve(value);
                    using var zmImgDispose = zmImgStream;
                    // Add new image part
                    var newImagePart = zmSlidePart.AddImagePart(zmImgPartType);
                    newImagePart.FeedData(zmImgStream);
                    var newImgRelId = zmSlidePart.GetIdOfPart(newImagePart);
                    var rNs2 = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                    // Update blip in zmPr > blipFill
                    var zmBlip = zmPr?.Descendants().FirstOrDefault(d => d.LocalName == "blip");
                    zmBlip?.SetAttribute(new OpenXmlAttribute("r", "embed", rNs2, newImgRelId));
                    // Update blip in fallback > blipFill
                    var fbBlipFill = fallback?.Descendants().FirstOrDefault(d => d.LocalName == "blipFill");
                    var fbBlip = fbBlipFill?.ChildElements.FirstOrDefault(e => e.LocalName == "blip");
                    fbBlip?.SetAttribute(new OpenXmlAttribute("r", "embed", rNs2, newImgRelId));
                    // Set imageType to "cover" so PowerPoint uses our image instead of auto-preview
                    zmPr?.SetAttribute(new OpenXmlAttribute("", "imageType", null!, "cover"));
                    break;
                }
                case "imagetype":
                    zmPr?.SetAttribute(new OpenXmlAttribute("", "imageType", null!, value));
                    break;
                default:
                    if (unsupported.Count == 0)
                        unsupported.Add($"{key} (valid zoom props: target, image, src, path, imagetype, x, y, width, height)");
                    else
                        unsupported.Add(key);
                    break;
            }
        }
        GetSlide(zmSlidePart).Save();
        return unsupported;
    }

    private List<string> SetModel3DByPath(Match model3dSetMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(model3dSetMatch.Groups[1].Value);
        var m3dIdx = int.Parse(model3dSetMatch.Groups[2].Value);
        var m3dSlideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > m3dSlideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {m3dSlideParts.Count})");
        var m3dSlidePart = m3dSlideParts[slideIdx - 1];
        var m3dShapeTree = GetSlide(m3dSlidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");
        var model3dElements = GetModel3DElements(m3dShapeTree);
        if (m3dIdx < 1 || m3dIdx > model3dElements.Count)
            throw new ArgumentException($"3D model {m3dIdx} not found (total: {model3dElements.Count})");

        var acElement = model3dElements[m3dIdx - 1];
        var choice = acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Choice");
        var fallback = acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Fallback");
        var sp = choice?.ChildElements.FirstOrDefault(e => e.LocalName == "graphicFrame")
              ?? choice?.ChildElements.FirstOrDefault(e => e.LocalName == "sp");

        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "x" or "y" or "width" or "height":
                {
                    var emu = ParseEmu(value);
                    // Update xfrm (graphicFrame level or spPr level)
                    var xfrmEl = sp?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
                    if (xfrmEl == null)
                    {
                        var spPr = sp?.ChildElements.FirstOrDefault(e => e.LocalName == "spPr");
                        xfrmEl = spPr?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
                    }
                    if (xfrmEl != null)
                    {
                        if (key.ToLowerInvariant() is "x" or "y")
                        {
                            var off = xfrmEl.ChildElements.FirstOrDefault(e => e.LocalName == "off");
                            off?.SetAttribute(new OpenXmlAttribute("", key.ToLowerInvariant(), null!, emu.ToString()));
                        }
                        else
                        {
                            var attrName = key.ToLowerInvariant() == "width" ? "cx" : "cy";
                            var ext = xfrmEl.ChildElements.FirstOrDefault(e => e.LocalName == "ext");
                            ext?.SetAttribute(new OpenXmlAttribute("", attrName, null!, emu.ToString()));
                        }
                    }
                    // Also update fallback pic spPr
                    var fbPic = fallback?.ChildElements.FirstOrDefault(e => e.LocalName == "pic");
                    var fbSpPr = fbPic?.ChildElements.FirstOrDefault(e => e.LocalName == "spPr");
                    var fbXfrm = fbSpPr?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
                    if (fbXfrm != null)
                    {
                        if (key.ToLowerInvariant() is "x" or "y")
                        {
                            var off = fbXfrm.ChildElements.FirstOrDefault(e => e.LocalName == "off");
                            off?.SetAttribute(new OpenXmlAttribute("", key.ToLowerInvariant(), null!, emu.ToString()));
                        }
                        else
                        {
                            var attrName = key.ToLowerInvariant() == "width" ? "cx" : "cy";
                            var ext = fbXfrm.ChildElements.FirstOrDefault(e => e.LocalName == "ext");
                            ext?.SetAttribute(new OpenXmlAttribute("", attrName, null!, emu.ToString()));
                        }
                    }
                    break;
                }
                case "name":
                {
                    var nvSpPr = sp?.ChildElements.FirstOrDefault(e => e.LocalName == "nvGraphicFramePr")
                              ?? sp?.ChildElements.FirstOrDefault(e => e.LocalName == "nvSpPr");
                    var cNvPr = nvSpPr?.ChildElements.FirstOrDefault(e => e.LocalName == "cNvPr");
                    cNvPr?.SetAttribute(new OpenXmlAttribute("", "name", null!, value));
                    // Also update fallback name
                    var fbPic = fallback?.ChildElements.FirstOrDefault(e => e.LocalName == "pic");
                    var fbCNvPr = fbPic?.Descendants().FirstOrDefault(d => d.LocalName == "cNvPr");
                    fbCNvPr?.SetAttribute(new OpenXmlAttribute("", "name", null!, value));
                    break;
                }
                case "rotx" or "roty" or "rotz":
                {
                    var model3dEl = acElement.Descendants().FirstOrDefault(d => d.LocalName == "model3d");
                    var trans = model3dEl?.ChildElements.FirstOrDefault(e => e.LocalName == "trans");
                    if (trans != null)
                    {
                        var rot = trans.ChildElements.FirstOrDefault(e => e.LocalName == "rot");
                        if (rot == null)
                        {
                            rot = new OpenXmlUnknownElement("am3d", "rot", Am3dNs);
                            trans.AppendChild(rot);
                        }
                        var attrName = key.ToLowerInvariant() switch { "rotx" => "ax", "roty" => "ay", _ => "az" };
                        rot.SetAttribute(new OpenXmlAttribute("", attrName, null!, ParseAngle60k(value).ToString()));
                    }
                    break;
                }
                default:
                    unsupported.Add(key);
                    break;
            }
        }
        GetSlide(m3dSlidePart).Save();
        return unsupported;
    }

    private List<string> SetOleByPath(Match oleSetMatch, Dictionary<string, string> properties)
    {
        var oleSlideIdx = int.Parse(oleSetMatch.Groups[1].Value);
        var oleEntryIdx = int.Parse(oleSetMatch.Groups[2].Value);
        var oleSlideParts = GetSlideParts().ToList();
        if (oleSlideIdx < 1 || oleSlideIdx > oleSlideParts.Count)
            throw new ArgumentException($"Slide {oleSlideIdx} not found (total: {oleSlideParts.Count})");
        var oleSlidePart = oleSlideParts[oleSlideIdx - 1];
        var oleShapeTree = GetSlide(oleSlidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");
        var oleFrames = oleShapeTree.Elements<GraphicFrame>()
            .Where(gf => gf.Descendants<DocumentFormat.OpenXml.Presentation.OleObject>().Any())
            .ToList();
        if (oleEntryIdx < 1 || oleEntryIdx > oleFrames.Count)
            throw new ArgumentException($"OLE object {oleEntryIdx} not found (total: {oleFrames.Count})");
        var oleFrame = oleFrames[oleEntryIdx - 1];
        var oleEl = oleFrame.Descendants<DocumentFormat.OpenXml.Presentation.OleObject>().First();
        var oleUnsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "path" or "src":
                {
                    // Delete old payload part and attach the new one.
                    if (oleEl.Id?.Value is string oldRel && !string.IsNullOrEmpty(oldRel))
                    {
                        try { oleSlidePart.DeletePart(oldRel); } catch { }
                    }
                    var (newRel, _) = OfficeCli.Core.OleHelper.AddEmbeddedPart(oleSlidePart, value, _filePath);
                    oleEl.Id = newRel;
                    // Auto-refresh progId from the new extension unless
                    // the caller explicitly pinned one in the same call.
                    if (!properties.ContainsKey("progId") && !properties.ContainsKey("progid"))
                    {
                        var autoProgId = OfficeCli.Core.OleHelper.DetectProgId(value);
                        OfficeCli.Core.OleHelper.ValidateProgId(autoProgId);
                        oleEl.ProgId = autoProgId;
                    }
                    break;
                }
                case "progid":
                    OfficeCli.Core.OleHelper.ValidateProgId(value);
                    oleEl.ProgId = value;
                    break;
                case "name":
                    oleEl.Name = value;
                    break;
                case "display":
                {
                    // Strict: only "icon" or "content" are accepted —
                    // see OleHelper.NormalizeOleDisplay.
                    var oleDisp = OfficeCli.Core.OleHelper.NormalizeOleDisplay(value);
                    oleEl.ShowAsIcon = oleDisp != "content";
                    break;
                }
                case "x" or "y" or "width" or "height":
                {
                    var xfrm = oleFrame.Transform ?? (oleFrame.Transform = new Transform());
                    var off = xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset { X = 0, Y = 0 });
                    var ext = xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents { Cx = 0, Cy = 0 });
                    var emu = ParseEmu(value);
                    var k = key.ToLowerInvariant();
                    // CONSISTENCY(ole-nonnegative-size): width/height are
                    // OOXML positive-sized types (ST_PositiveCoordinate).
                    // Silently storing a negative EMU breaks the shape
                    // frame and opens unpredictably in PowerPoint. Reject
                    // it explicitly; x/y may legitimately be negative
                    // (off-slide anchors) so they pass through.
                    if ((k == "width" || k == "height") && emu < 0)
                        throw new ArgumentException($"{k} must be non-negative");
                    switch (k)
                    {
                        case "x": off.X = emu; break;
                        case "y": off.Y = emu; break;
                        case "width": ext.Cx = emu; break;
                        case "height": ext.Cy = emu; break;
                    }
                    break;
                }
                default:
                    oleUnsupported.Add(key);
                    break;
            }
        }
        GetSlide(oleSlidePart).Save();
        return oleUnsupported;
    }

    private List<string> SetMediaByPath(Match mediaSetMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(mediaSetMatch.Groups[1].Value);
        var mediaType = mediaSetMatch.Groups[2].Value;
        var mediaIdx = int.Parse(mediaSetMatch.Groups[3].Value);

        var slideParts4 = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts4.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts4.Count})");

        var slidePart = slideParts4[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");

        var mediaPics = shapeTree.Elements<Picture>()
            .Where(p =>
            {
                var nvPr = p.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
                return mediaType == "video"
                    ? nvPr?.GetFirstChild<Drawing.VideoFromFile>() != null
                    : nvPr?.GetFirstChild<Drawing.AudioFromFile>() != null;
            }).ToList();
        if (mediaIdx < 1 || mediaIdx > mediaPics.Count)
            throw new ArgumentException($"{mediaType} {mediaIdx} not found (total: {mediaPics.Count})");

        var pic = mediaPics[mediaIdx - 1];
        var shapeId = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value;
        var unsupported = new List<string>();

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "volume":
                {
                    if (shapeId == null) { unsupported.Add(key); break; }
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var volVal)
                        || double.IsNaN(volVal) || double.IsInfinity(volVal))
                        throw new ArgumentException($"Invalid volume value: '{value}'. Expected a finite number (0-100).");
                    var vol = (int)(volVal * 1000); // 0-100 → 0-100000
                    var mediaNode = FindMediaTimingNode(slidePart, shapeId.Value);
                    if (mediaNode != null) mediaNode.Volume = vol;
                    break;
                }
                case "autoplay":
                {
                    if (shapeId == null) { unsupported.Add(key); break; }
                    var mediaNode = FindMediaTimingNode(slidePart, shapeId.Value);
                    var cTn = mediaNode?.CommonTimeNode;
                    var startCond = cTn?.StartConditionList?.GetFirstChild<Condition>();
                    if (startCond != null)
                        startCond.Delay = IsTruthy(value) ? "0" : "indefinite";
                    break;
                }
                case "trimstart":
                {
                    var nvPr = pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
                    var p14Media = nvPr?.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().FirstOrDefault();
                    if (p14Media != null)
                    {
                        var trim = p14Media.MediaTrim ?? (p14Media.MediaTrim = new DocumentFormat.OpenXml.Office2010.PowerPoint.MediaTrim());
                        trim.Start = value;
                    }
                    break;
                }
                case "trimend":
                {
                    var nvPr = pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
                    var p14Media = nvPr?.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().FirstOrDefault();
                    if (p14Media != null)
                    {
                        var trim = p14Media.MediaTrim ?? (p14Media.MediaTrim = new DocumentFormat.OpenXml.Office2010.PowerPoint.MediaTrim());
                        trim.End = value;
                    }
                    break;
                }
                case "x" or "y" or "width" or "height":
                {
                    var spPr = pic.ShapeProperties ?? (pic.ShapeProperties = new ShapeProperties());
                    var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
                    TryApplyPositionSize(key.ToLowerInvariant(), value,
                        xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset()),
                        xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents()));
                    break;
                }
                default:
                    if (unsupported.Count == 0)
                        unsupported.Add($"{key} (valid media props: volume, autoplay, trimstart, trimend, x, y, width, height)");
                    else
                        unsupported.Add(key);
                    break;
            }
        }
        GetSlide(slidePart).Save();
        return unsupported;
    }

}
