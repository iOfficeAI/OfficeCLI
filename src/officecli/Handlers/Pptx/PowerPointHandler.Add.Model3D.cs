// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private const string Am3dNs = "http://schemas.microsoft.com/office/drawing/2017/model3d";
    private const string Model3dRelType = "http://schemas.microsoft.com/office/2017/06/relationships/model3d";
    // PowerPoint uses "model/gltf.binary" (dot, not dash)
    private const string GlbContentType = "model/gltf.binary";

    private string AddModel3D(string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("path", out var modelPath) &&
            !properties.TryGetValue("src", out modelPath))
            throw new ArgumentException("'src' property is required for 3dmodel type");

        var slideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
        if (!slideMatch.Success)
            throw new ArgumentException("3D models must be added to a slide: /slide[N]");

        var slideIdx = int.Parse(slideMatch.Groups[1].Value);
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");

        // Resolve source (local path, HTTP URL, or data URI)
        var (modelStream, fileExt) = OfficeCli.Core.FileSource.Resolve(modelPath);
        using var modelStreamDispose = modelStream;
        if (fileExt != ".glb")
            throw new ArgumentException($"Unsupported 3D model format: {fileExt}. Only .glb (glTF-Binary) is supported.");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");

        // Parse GLB bounding box for centering
        var glbBounds = ParseGlbBoundingBox(modelStream);

        // Embed .glb file as an extended part
        modelStream.Position = 0;
        var modelPart = slidePart.AddExtendedPart(Model3dRelType, GlbContentType, ".glb");
        modelPart.FeedData(modelStream);
        var modelRelId = slidePart.GetIdOfPart(modelPart);

        // Create fallback placeholder image
        byte[] placeholderPng = GenerateZoomPlaceholderPng();
        var imagePart = slidePart.AddImagePart(ImagePartType.Png);
        using (var ms = new MemoryStream(placeholderPng))
            imagePart.FeedData(ms);
        var imageRelId = slidePart.GetIdOfPart(imagePart);

        // Position and size (default: 10cm x 10cm, centered)
        long cx = 3600000; // ~10cm
        long cy = 3600000;
        if (properties.TryGetValue("width", out var w)) cx = ParseEmu(w);
        if (properties.TryGetValue("height", out var h)) cy = ParseEmu(h);
        var (slideW, slideH) = GetSlideSize();
        long x = (slideW - cx) / 2;
        long y = (slideH - cy) / 2;
        if (properties.TryGetValue("x", out var xs) || properties.TryGetValue("left", out xs)) x = ParseEmu(xs);
        if (properties.TryGetValue("y", out var ys) || properties.TryGetValue("top", out ys)) y = ParseEmu(ys);

        var shapeId = GenerateUniqueShapeId(shapeTree);
        var shapeName = properties.GetValueOrDefault("name", $"3D Model {GetModel3DElements(shapeTree).Count + 1}");

        // Namespaces
        var mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        var pNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
        var aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var a16Ns = "http://schemas.microsoft.com/office/drawing/2014/main";

        var creationGuid = Guid.NewGuid().ToString("B").ToUpperInvariant();

        // Build mc:AlternateContent
        var acElement = new OpenXmlUnknownElement("mc", "AlternateContent", mcNs);

        // === mc:Choice (for clients that support 3D models) ===
        var choiceElement = new OpenXmlUnknownElement("mc", "Choice", mcNs);
        choiceElement.SetAttribute(new OpenXmlAttribute("", "Requires", null!, "am3d"));
        choiceElement.AddNamespaceDeclaration("am3d", Am3dNs);

        // Use p:graphicFrame (NOT p:sp) — same as zoom and native PowerPoint
        var gf = new OpenXmlUnknownElement("p", "graphicFrame", pNs);
        gf.AddNamespaceDeclaration("a", aNs);
        gf.AddNamespaceDeclaration("r", rNs);

        // nvGraphicFramePr
        var nvGfPr = new OpenXmlUnknownElement("p", "nvGraphicFramePr", pNs);
        var cNvPr = new OpenXmlUnknownElement("p", "cNvPr", pNs);
        cNvPr.SetAttribute(new OpenXmlAttribute("", "id", null!, shapeId.ToString()));
        cNvPr.SetAttribute(new OpenXmlAttribute("", "name", null!, shapeName));
        // creationId extension
        var extLst = new OpenXmlUnknownElement("a", "extLst", aNs);
        var ext = new OpenXmlUnknownElement("a", "ext", aNs);
        ext.SetAttribute(new OpenXmlAttribute("", "uri", null!, "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"));
        var creationId = new OpenXmlUnknownElement("a16", "creationId", a16Ns);
        creationId.SetAttribute(new OpenXmlAttribute("", "id", null!, creationGuid));
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
        gf.AppendChild(nvGfPr);

        // xfrm (position/size on the graphicFrame level)
        var gfXfrm = new OpenXmlUnknownElement("p", "xfrm", pNs);
        var gfOff = new OpenXmlUnknownElement("a", "off", aNs);
        gfOff.SetAttribute(new OpenXmlAttribute("", "x", null!, x.ToString()));
        gfOff.SetAttribute(new OpenXmlAttribute("", "y", null!, y.ToString()));
        var gfExt = new OpenXmlUnknownElement("a", "ext", aNs);
        gfExt.SetAttribute(new OpenXmlAttribute("", "cx", null!, cx.ToString()));
        gfExt.SetAttribute(new OpenXmlAttribute("", "cy", null!, cy.ToString()));
        gfXfrm.AppendChild(gfOff);
        gfXfrm.AppendChild(gfExt);
        gf.AppendChild(gfXfrm);

        // a:graphic > a:graphicData[uri=am3d] > am3d:model3d
        var graphic = new OpenXmlUnknownElement("a", "graphic", aNs);
        var graphicData = new OpenXmlUnknownElement("a", "graphicData", aNs);
        graphicData.SetAttribute(new OpenXmlAttribute("", "uri", null!, Am3dNs));

        var model3d = BuildModel3DElement(modelRelId, imageRelId, cx, cy, properties, glbBounds);
        graphicData.AppendChild(model3d);
        graphic.AppendChild(graphicData);
        gf.AppendChild(graphic);

        choiceElement.AppendChild(gf);

        // === mc:Fallback (static image for older clients) ===
        var fallbackElement = new OpenXmlUnknownElement("mc", "Fallback", mcNs);
        var fbPic = new OpenXmlUnknownElement("p", "pic", pNs);
        fbPic.AddNamespaceDeclaration("a", aNs);
        fbPic.AddNamespaceDeclaration("r", rNs);

        var fbNvPicPr = new OpenXmlUnknownElement("p", "nvPicPr", pNs);
        var fbCNvPr = new OpenXmlUnknownElement("p", "cNvPr", pNs);
        fbCNvPr.SetAttribute(new OpenXmlAttribute("", "id", null!, shapeId.ToString()));
        fbCNvPr.SetAttribute(new OpenXmlAttribute("", "name", null!, shapeName));
        // Same creationId
        var fbExtLst = new OpenXmlUnknownElement("a", "extLst", aNs);
        var fbExt = new OpenXmlUnknownElement("a", "ext", aNs);
        fbExt.SetAttribute(new OpenXmlAttribute("", "uri", null!, "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"));
        var fbCreationId = new OpenXmlUnknownElement("a16", "creationId", a16Ns);
        fbCreationId.SetAttribute(new OpenXmlAttribute("", "id", null!, creationGuid));
        fbExt.AppendChild(fbCreationId);
        fbExtLst.AppendChild(fbExt);
        fbCNvPr.AppendChild(fbExtLst);
        fbNvPicPr.AppendChild(fbCNvPr);

        var fbCNvPicPr = new OpenXmlUnknownElement("p", "cNvPicPr", pNs);
        var picLocks = new OpenXmlUnknownElement("a", "picLocks", aNs);
        foreach (var lockAttr in new[] { "noGrp", "noRot", "noChangeAspect", "noMove", "noResize",
            "noEditPoints", "noAdjustHandles", "noChangeArrowheads", "noChangeShapeType", "noCrop" })
            picLocks.SetAttribute(new OpenXmlAttribute("", lockAttr, null!, "1"));
        fbCNvPicPr.AppendChild(picLocks);
        fbNvPicPr.AppendChild(fbCNvPicPr);
        fbNvPicPr.AppendChild(new OpenXmlUnknownElement("p", "nvPr", pNs));
        fbPic.AppendChild(fbNvPicPr);

        // Fallback blipFill
        var fbBlipFill = new OpenXmlUnknownElement("p", "blipFill", pNs);
        var fbBlip = new OpenXmlUnknownElement("a", "blip", aNs);
        fbBlip.SetAttribute(new OpenXmlAttribute("r", "embed", rNs, imageRelId));
        fbBlipFill.AppendChild(fbBlip);
        var fbStretch = new OpenXmlUnknownElement("a", "stretch", aNs);
        fbStretch.AppendChild(new OpenXmlUnknownElement("a", "fillRect", aNs));
        fbBlipFill.AppendChild(fbStretch);
        fbPic.AppendChild(fbBlipFill);

        // Fallback spPr
        var fbSpPr = new OpenXmlUnknownElement("p", "spPr", pNs);
        var fbXfrm = new OpenXmlUnknownElement("a", "xfrm", aNs);
        var fbOff = new OpenXmlUnknownElement("a", "off", aNs);
        fbOff.SetAttribute(new OpenXmlAttribute("", "x", null!, x.ToString()));
        fbOff.SetAttribute(new OpenXmlAttribute("", "y", null!, y.ToString()));
        var fbExtSz = new OpenXmlUnknownElement("a", "ext", aNs);
        fbExtSz.SetAttribute(new OpenXmlAttribute("", "cx", null!, cx.ToString()));
        fbExtSz.SetAttribute(new OpenXmlAttribute("", "cy", null!, cy.ToString()));
        fbXfrm.AppendChild(fbOff);
        fbXfrm.AppendChild(fbExtSz);
        fbSpPr.AppendChild(fbXfrm);
        var fbGeom = new OpenXmlUnknownElement("a", "prstGeom", aNs);
        fbGeom.SetAttribute(new OpenXmlAttribute("", "prst", null!, "rect"));
        fbGeom.AppendChild(new OpenXmlUnknownElement("a", "avLst", aNs));
        fbSpPr.AppendChild(fbGeom);
        fbPic.AppendChild(fbSpPr);

        fallbackElement.AppendChild(fbPic);

        acElement.AppendChild(choiceElement);
        acElement.AppendChild(fallbackElement);
        InsertAtPosition(shapeTree, acElement, index);

        // Ensure am3d namespace is declared on slide root
        var slide = GetSlide(slidePart);
        try { slide.AddNamespaceDeclaration("am3d", Am3dNs); } catch { }
        try { slide.AddNamespaceDeclaration("mc", mcNs); } catch { }
        var ignorable = slide.MCAttributes?.Ignorable?.Value;
        if (ignorable == null || !ignorable.Contains("am3d"))
        {
            slide.MCAttributes ??= new MarkupCompatibilityAttributes();
            slide.MCAttributes.Ignorable = string.IsNullOrEmpty(ignorable) ? "am3d" : $"{ignorable} am3d";
        }
        slide.Save();

        var model3dCount = GetModel3DElements(shapeTree).Count;
        return $"/slide[{slideIdx}]/model3d[{model3dCount}]";
    }

    /// <summary>
    /// Build the am3d:model3d element with camera, transform, viewport, and lighting.
    /// Follows the native PowerPoint XML structure exactly.
    /// </summary>
    private OpenXmlUnknownElement BuildModel3DElement(
        string modelRelId, string imageRelId, long cx, long cy,
        Dictionary<string, string> properties, GlbBoundingBox bounds)
    {
        var aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        var model3d = new OpenXmlUnknownElement("am3d", "model3d", Am3dNs);
        model3d.SetAttribute(new OpenXmlAttribute("r", "embed", rNs, modelRelId));

        // mpu = 1 / effectiveMaxExtent
        // effectiveMaxExtent = rawMaxExtent × nodeScale (from GLB root node transform)
        var mpuVal = bounds.MaxExtent > 0 ? 1.0 / bounds.MaxExtent : 0.5;

        // Half-extents (already in am3d coordinates from ParseGlbBoundingBox)
        var halfExtX = bounds.ExtentX / 2.0;
        var halfExtY = bounds.ExtentY / 2.0;
        var halfExtZ = bounds.ExtentZ / 2.0;

        // Radius for camera distance: normFactor * ‖halfExtents‖
        // normFactor internally = 1/(2*maxHalfExt), but mpu may differ due to mpuFactor
        var maxHalfExt = Math.Max(halfExtX, Math.Max(halfExtY, halfExtZ));
        var normFactor = maxHalfExt > 0 ? 1.0 / (2.0 * maxHalfExt) : 0.5;
        var radius = normFactor * Math.Sqrt(halfExtX * halfExtX + halfExtY * halfExtY + halfExtZ * halfExtZ);
        if (radius == 0) radius = 1.0;

        // FOV (default 45°)
        const int fov60k = 2700000;
        var fovHalfRad = fov60k / 60000.0 * Math.PI / 180.0 * 0.5;

        // Camera Z distance (perspective mode)
        var cameraZ = radius / Math.Sin(fovHalfRad);

        // viewportSz: PPT computes this via tight-wrap 3D rendering.
        // Without a renderer, use max(cx,cy) which gives ≤6% error vs PPT native.
        var viewportSize = Math.Max(cx, cy);

        // 1. spPr (internal shape properties for the 3D model viewport)
        var spPr = new OpenXmlUnknownElement("am3d", "spPr", Am3dNs);
        var xfrm = new OpenXmlUnknownElement("a", "xfrm", aNs);
        var off = new OpenXmlUnknownElement("a", "off", aNs);
        off.SetAttribute(new OpenXmlAttribute("", "x", null!, "0"));
        off.SetAttribute(new OpenXmlAttribute("", "y", null!, "0"));
        var ext = new OpenXmlUnknownElement("a", "ext", aNs);
        ext.SetAttribute(new OpenXmlAttribute("", "cx", null!, cx.ToString()));
        ext.SetAttribute(new OpenXmlAttribute("", "cy", null!, cy.ToString()));
        xfrm.AppendChild(off);
        xfrm.AppendChild(ext);
        spPr.AppendChild(xfrm);
        var prstGeom = new OpenXmlUnknownElement("a", "prstGeom", aNs);
        prstGeom.SetAttribute(new OpenXmlAttribute("", "prst", null!, "rect"));
        prstGeom.AppendChild(new OpenXmlUnknownElement("a", "avLst", aNs));
        spPr.AppendChild(prstGeom);
        model3d.AppendChild(spPr);

        // 2. camera — perspective, looking at origin from z-axis
        var computedCamZ = (long)(cameraZ * 36000000.0);
        var camPosX = properties.GetValueOrDefault("camerax", "0");
        var camPosY = properties.GetValueOrDefault("cameray", "0");
        var camPosZ = properties.GetValueOrDefault("cameraz", computedCamZ.ToString());

        var camera = new OpenXmlUnknownElement("am3d", "camera", Am3dNs);
        var camPos = new OpenXmlUnknownElement("am3d", "pos", Am3dNs);
        camPos.SetAttribute(new OpenXmlAttribute("", "x", null!, camPosX));
        camPos.SetAttribute(new OpenXmlAttribute("", "y", null!, camPosY));
        camPos.SetAttribute(new OpenXmlAttribute("", "z", null!, camPosZ));
        camera.AppendChild(camPos);
        var camUp = new OpenXmlUnknownElement("am3d", "up", Am3dNs);
        camUp.SetAttribute(new OpenXmlAttribute("", "dx", null!, "0"));
        camUp.SetAttribute(new OpenXmlAttribute("", "dy", null!, "36000000"));
        camUp.SetAttribute(new OpenXmlAttribute("", "dz", null!, "0"));
        camera.AppendChild(camUp);
        var camLookAt = new OpenXmlUnknownElement("am3d", "lookAt", Am3dNs);
        camLookAt.SetAttribute(new OpenXmlAttribute("", "x", null!, "0"));
        camLookAt.SetAttribute(new OpenXmlAttribute("", "y", null!, "0"));
        camLookAt.SetAttribute(new OpenXmlAttribute("", "z", null!, "0"));
        camera.AppendChild(camLookAt);
        var perspective = new OpenXmlUnknownElement("am3d", "perspective", Am3dNs);
        perspective.SetAttribute(new OpenXmlAttribute("", "fov", null!, fov60k.ToString()));
        camera.AppendChild(perspective);
        model3d.AppendChild(camera);

        // 3. trans — mpu, preTrans, scale, rot, postTrans
        var trans = new OpenXmlUnknownElement("am3d", "trans", Am3dNs);

        // mpu = normFactor = 1/fullMaxExtent, stored as PosRatio n/1000000
        var mpuN = (long)(mpuVal * 1000000);
        var mpu = new OpenXmlUnknownElement("am3d", "meterPerModelUnit", Am3dNs);
        mpu.SetAttribute(new OpenXmlAttribute("", "n", null!, mpuN.ToString()));
        mpu.SetAttribute(new OpenXmlAttribute("", "d", null!, "1000000"));
        trans.AppendChild(mpu);

        // preTrans: center model at origin. bounds.Center* is already in am3d coordinates.
        var preTransScale = mpuVal * 36000000.0;
        var preTrans = new OpenXmlUnknownElement("am3d", "preTrans", Am3dNs);
        preTrans.SetAttribute(new OpenXmlAttribute("", "dx", null!, ((long)(-bounds.CenterX * preTransScale)).ToString()));
        preTrans.SetAttribute(new OpenXmlAttribute("", "dy", null!, ((long)(-bounds.CenterY * preTransScale)).ToString()));
        preTrans.SetAttribute(new OpenXmlAttribute("", "dz", null!, ((long)(-bounds.CenterZ * preTransScale)).ToString()));
        trans.AppendChild(preTrans);

        // scale (default 1:1:1)
        var scale = new OpenXmlUnknownElement("am3d", "scale", Am3dNs);
        foreach (var axis in new[] { "sx", "sy", "sz" })
        {
            var s = new OpenXmlUnknownElement("am3d", axis, Am3dNs);
            s.SetAttribute(new OpenXmlAttribute("", "n", null!, "1000000"));
            s.SetAttribute(new OpenXmlAttribute("", "d", null!, "1000000"));
            scale.AppendChild(s);
        }
        trans.AppendChild(scale);

        // rot
        var rot = new OpenXmlUnknownElement("am3d", "rot", Am3dNs);
        var rotXVal = "0"; var rotYVal = "0"; var rotZVal = "0";
        if (properties.TryGetValue("rotx", out var rx)) rotXVal = ParseAngle60k(rx).ToString();
        if (properties.TryGetValue("roty", out var ry)) rotYVal = ParseAngle60k(ry).ToString();
        if (properties.TryGetValue("rotz", out var rz)) rotZVal = ParseAngle60k(rz).ToString();
        rot.SetAttribute(new OpenXmlAttribute("", "ax", null!, rotXVal));
        rot.SetAttribute(new OpenXmlAttribute("", "ay", null!, rotYVal));
        rot.SetAttribute(new OpenXmlAttribute("", "az", null!, rotZVal));
        trans.AppendChild(rot);

        // postTrans
        var postTrans = new OpenXmlUnknownElement("am3d", "postTrans", Am3dNs);
        postTrans.SetAttribute(new OpenXmlAttribute("", "dx", null!, "0"));
        postTrans.SetAttribute(new OpenXmlAttribute("", "dy", null!, "0"));
        postTrans.SetAttribute(new OpenXmlAttribute("", "dz", null!, "0"));
        trans.AppendChild(postTrans);

        model3d.AppendChild(trans);

        // 4. raster (cached rendering) — use am3d:blip (not a:blip)
        var raster = new OpenXmlUnknownElement("am3d", "raster", Am3dNs);
        raster.SetAttribute(new OpenXmlAttribute("", "rName", null!, "Office3DRenderer"));
        raster.SetAttribute(new OpenXmlAttribute("", "rVer", null!, "16.0.8326"));
        var rasterBlip = new OpenXmlUnknownElement("am3d", "blip", Am3dNs);
        rasterBlip.SetAttribute(new OpenXmlAttribute("r", "embed", rNs, imageRelId));
        raster.AppendChild(rasterBlip);
        model3d.AppendChild(raster);

        // 5. objViewport — matches the shape size
        var viewport = new OpenXmlUnknownElement("am3d", "objViewport", Am3dNs);
        viewport.SetAttribute(new OpenXmlAttribute("", "viewportSz", null!, viewportSize.ToString()));
        model3d.AppendChild(viewport);

        // 6. ambientLight — use scrgbClr like native PowerPoint
        var ambient = new OpenXmlUnknownElement("am3d", "ambientLight", Am3dNs);
        var ambClr = new OpenXmlUnknownElement("am3d", "clr", Am3dNs);
        var ambScrgb = new OpenXmlUnknownElement("a", "scrgbClr", aNs);
        ambScrgb.SetAttribute(new OpenXmlAttribute("", "r", null!, "50000"));
        ambScrgb.SetAttribute(new OpenXmlAttribute("", "g", null!, "50000"));
        ambScrgb.SetAttribute(new OpenXmlAttribute("", "b", null!, "50000"));
        ambClr.AppendChild(ambScrgb);
        ambient.AppendChild(ambClr);
        var ambIll = new OpenXmlUnknownElement("am3d", "illuminance", Am3dNs);
        ambIll.SetAttribute(new OpenXmlAttribute("", "n", null!, "500000"));
        ambIll.SetAttribute(new OpenXmlAttribute("", "d", null!, "1000000"));
        ambient.AppendChild(ambIll);
        model3d.AppendChild(ambient);

        // 7. ptLight — three point lights (matching native PowerPoint)
        AddPointLight(model3d, aNs, "100000", "75000", "50000", "9765625", "21959998", "70920001", "16344003");
        AddPointLight(model3d, aNs, "40000", "60000", "95000", "12250000", "-37964106", "51130435", "57631972");
        AddPointLight(model3d, aNs, "86837", "72700", "100000", "3125000", "-37739122", "58056624", "-34769649");

        return model3d;
    }

    private static void AddPointLight(OpenXmlUnknownElement parent, string aNs,
        string r, string g, string b, string intensity,
        string posX, string posY, string posZ)
    {
        var ptLight = new OpenXmlUnknownElement("am3d", "ptLight", Am3dNs);
        ptLight.SetAttribute(new OpenXmlAttribute("", "rad", null!, "0"));
        var ptClr = new OpenXmlUnknownElement("am3d", "clr", Am3dNs);
        var ptScrgb = new OpenXmlUnknownElement("a", "scrgbClr", aNs);
        ptScrgb.SetAttribute(new OpenXmlAttribute("", "r", null!, r));
        ptScrgb.SetAttribute(new OpenXmlAttribute("", "g", null!, g));
        ptScrgb.SetAttribute(new OpenXmlAttribute("", "b", null!, b));
        ptClr.AppendChild(ptScrgb);
        ptLight.AppendChild(ptClr);
        var ptInt = new OpenXmlUnknownElement("am3d", "intensity", Am3dNs);
        ptInt.SetAttribute(new OpenXmlAttribute("", "n", null!, intensity));
        ptInt.SetAttribute(new OpenXmlAttribute("", "d", null!, "1000000"));
        ptLight.AppendChild(ptInt);
        var ptPos = new OpenXmlUnknownElement("am3d", "pos", Am3dNs);
        ptPos.SetAttribute(new OpenXmlAttribute("", "x", null!, posX));
        ptPos.SetAttribute(new OpenXmlAttribute("", "y", null!, posY));
        ptPos.SetAttribute(new OpenXmlAttribute("", "z", null!, posZ));
        ptLight.AppendChild(ptPos);
        parent.AppendChild(ptLight);
    }

    /// <summary>
    /// Parse degrees to 60000ths-of-a-degree for am3d rotation attributes.
    /// </summary>
    private static int ParseAngle60k(string value)
    {
        if (!double.TryParse(value, System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out var deg))
            return 0;
        return (int)(deg * 60000);
    }

    /// <summary>
    /// Bounding box info extracted from a GLB file.
    /// Extents and center are in effective (scene-transformed) coordinates.
    /// RawMaxExtent is before node scale, NodeScale is the root node scale factor.
    /// </summary>
    private record GlbBoundingBox(
        double CenterX, double CenterY, double CenterZ,
        double ExtentX, double ExtentY, double ExtentZ,
        double MaxExtent, double MeterPerModelUnit,
        double RawMaxExtent, double NodeScale);

    /// <summary>
    /// Parse a GLB file and compute world-space AABB by traversing the scene graph,
    /// matching OSpectre's bounding box calculation.
    /// </summary>
    private static GlbBoundingBox ParseGlbBoundingBox(Stream glbStream)
    {
        try
        {
            glbStream.Position = 0;
            using var reader = new BinaryReader(glbStream, System.Text.Encoding.UTF8, leaveOpen: true);

            var magic = reader.ReadUInt32();
            var version = reader.ReadUInt32();
            var totalLen = reader.ReadUInt32();
            var chunkLen = reader.ReadUInt32();
            var chunkType = reader.ReadUInt32();
            var jsonBytes = reader.ReadBytes((int)chunkLen);
            var json = System.Text.Encoding.UTF8.GetString(jsonBytes);
            var doc = System.Text.Json.JsonDocument.Parse(json);
            var root = doc.RootElement;

            // 1. Build per-mesh local AABBs from accessors
            //    meshBounds[meshIndex] = (min, max) in local mesh space
            var meshBounds = new Dictionary<int, (double[] min, double[] max)>();
            if (root.TryGetProperty("meshes", out var meshes) &&
                root.TryGetProperty("accessors", out var accessors))
            {
                for (int mi = 0; mi < meshes.GetArrayLength(); mi++)
                {
                    double[] lMin = { double.MaxValue, double.MaxValue, double.MaxValue };
                    double[] lMax = { double.MinValue, double.MinValue, double.MinValue };
                    bool hasBounds = false;

                    var mesh = meshes[mi];
                    if (mesh.TryGetProperty("primitives", out var prims))
                    {
                        foreach (var prim in prims.EnumerateArray())
                        {
                            if (!prim.TryGetProperty("attributes", out var attrs)) continue;
                            if (!attrs.TryGetProperty("POSITION", out var posIdx)) continue;
                            var acc = accessors[posIdx.GetInt32()];
                            if (acc.TryGetProperty("min", out var mn) && acc.TryGetProperty("max", out var mx)
                                && mn.GetArrayLength() >= 3 && mx.GetArrayLength() >= 3)
                            {
                                hasBounds = true;
                                for (int i = 0; i < 3; i++)
                                {
                                    var lo = mn[i].GetDouble(); var hi = mx[i].GetDouble();
                                    if (lo < lMin[i]) lMin[i] = lo;
                                    if (hi > lMax[i]) lMax[i] = hi;
                                }
                            }
                        }
                    }
                    if (hasBounds)
                        meshBounds[mi] = (lMin, lMax);
                }
            }

            if (meshBounds.Count == 0)
                return new GlbBoundingBox(0, 0, 0, 1, 1, 1, 1, 0.5, 1, 1.0);

            // 2. Parse node transforms and traverse scene graph
            var nodesArr = root.TryGetProperty("nodes", out var nodesEl) ? nodesEl : default;
            int nodeCount = nodesArr.ValueKind == System.Text.Json.JsonValueKind.Array ? nodesArr.GetArrayLength() : 0;

            // World-space AABB accumulator
            double wMinX = double.MaxValue, wMinY = double.MaxValue, wMinZ = double.MaxValue;
            double wMaxX = double.MinValue, wMaxY = double.MinValue, wMaxZ = double.MinValue;
            bool hasWorldBounds = false;

            void TraverseNode(int nodeIdx, double[] parentMatrix)
            {
                if (nodeIdx < 0 || nodeIdx >= nodeCount) return;
                var node = nodesArr[nodeIdx];

                // Compute this node's local transform matrix (4x4 column-major)
                var local = GetNodeMatrix(node);
                var world = MultiplyMatrix4x4(parentMatrix, local);

                // If node has a mesh, transform its AABB corners to world space
                if (node.TryGetProperty("mesh", out var meshIdx) && meshBounds.TryGetValue(meshIdx.GetInt32(), out var mb))
                {
                    var (lMin, lMax) = mb;
                    // Transform 8 AABB corners
                    for (int cx = 0; cx < 2; cx++)
                    for (int cy = 0; cy < 2; cy++)
                    for (int cz = 0; cz < 2; cz++)
                    {
                        double px = cx == 0 ? lMin[0] : lMax[0];
                        double py = cy == 0 ? lMin[1] : lMax[1];
                        double pz = cz == 0 ? lMin[2] : lMax[2];
                        // Apply 4x4 column-major transform: result = M * [px,py,pz,1]
                        double wx = world[0] * px + world[4] * py + world[8]  * pz + world[12];
                        double wy = world[1] * px + world[5] * py + world[9]  * pz + world[13];
                        double wz = world[2] * px + world[6] * py + world[10] * pz + world[14];
                        if (wx < wMinX) wMinX = wx; if (wx > wMaxX) wMaxX = wx;
                        if (wy < wMinY) wMinY = wy; if (wy > wMaxY) wMaxY = wy;
                        if (wz < wMinZ) wMinZ = wz; if (wz > wMaxZ) wMaxZ = wz;
                        hasWorldBounds = true;
                    }
                }

                // Recurse into children
                if (node.TryGetProperty("children", out var children))
                    foreach (var child in children.EnumerateArray())
                        TraverseNode(child.GetInt32(), world);
            }

            // Identity matrix
            var identity = new double[] { 1,0,0,0, 0,1,0,0, 0,0,1,0, 0,0,0,1 };

            if (root.TryGetProperty("scenes", out var scenes) && scenes.GetArrayLength() > 0)
            {
                var scene = scenes[0];
                if (scene.TryGetProperty("nodes", out var sceneNodes))
                    foreach (var ni in sceneNodes.EnumerateArray())
                        TraverseNode(ni.GetInt32(), identity);
            }

            if (!hasWorldBounds)
                return new GlbBoundingBox(0, 0, 0, 1, 1, 1, 1, 0.5, 1, 1.0);

            // Use glTF world-space coordinates directly (no axis transform needed —
            // the 3D engine handles coordinate system conversion at render time)
            var ecx = (wMinX + wMaxX) / 2;
            var ecy = (wMinY + wMaxY) / 2;
            var ecz = (wMinZ + wMaxZ) / 2;
            var eex = wMaxX - wMinX;
            var eey = wMaxY - wMinY;
            var eez = wMaxZ - wMinZ;
            var maxExt = Math.Max(eex, Math.Max(eey, eez));
            var mpu = maxExt > 0 ? 1.0 / maxExt : 0.5;

            // RawMaxExtent/NodeScale kept for backward compat but not used in new formula
            var rawMaxExt = Math.Max(wMaxX - wMinX, Math.Max(wMaxY - wMinY, wMaxZ - wMinZ));
            double nodeScale = maxExt > 0 && rawMaxExt > 0 ? maxExt / rawMaxExt : 1.0;

            return new GlbBoundingBox(ecx, ecy, ecz, eex, eey, eez, maxExt, mpu, rawMaxExt, nodeScale);
        }
        catch
        {
            return new GlbBoundingBox(0, 0, 0, 1, 1, 1, 1, 0.5, 1, 1.0);
        }
    }

    /// <summary>
    /// Get the 4x4 column-major transform matrix from a glTF node.
    /// Supports "matrix", "scale"/"rotation"/"translation" (TRS), or identity.
    /// </summary>
    private static double[] GetNodeMatrix(System.Text.Json.JsonElement node)
    {
        if (node.TryGetProperty("matrix", out var mat) && mat.GetArrayLength() == 16)
        {
            var m = new double[16];
            for (int i = 0; i < 16; i++) m[i] = mat[i].GetDouble();
            return m;
        }

        // TRS decomposition → 4x4 column-major
        double tx = 0, ty = 0, tz = 0;
        double qx = 0, qy = 0, qz = 0, qw = 1;
        double sx = 1, sy = 1, sz = 1;

        if (node.TryGetProperty("translation", out var t) && t.GetArrayLength() == 3)
        { tx = t[0].GetDouble(); ty = t[1].GetDouble(); tz = t[2].GetDouble(); }
        if (node.TryGetProperty("rotation", out var r) && r.GetArrayLength() == 4)
        { qx = r[0].GetDouble(); qy = r[1].GetDouble(); qz = r[2].GetDouble(); qw = r[3].GetDouble(); }
        if (node.TryGetProperty("scale", out var s) && s.GetArrayLength() == 3)
        { sx = s[0].GetDouble(); sy = s[1].GetDouble(); sz = s[2].GetDouble(); }

        // Quaternion to rotation matrix, then apply scale and translation
        double x2 = qx + qx, y2 = qy + qy, z2 = qz + qz;
        double xx = qx * x2, xy = qx * y2, xz = qx * z2;
        double yy = qy * y2, yz = qy * z2, zz = qz * z2;
        double wx = qw * x2, wy = qw * y2, wz = qw * z2;

        return new[]
        {
            (1 - yy - zz) * sx, (xy + wz) * sx,     (xz - wy) * sx,     0,
            (xy - wz) * sy,     (1 - xx - zz) * sy,  (yz + wx) * sy,     0,
            (xz + wy) * sz,     (yz - wx) * sz,      (1 - xx - yy) * sz, 0,
            tx,                  ty,                   tz,                  1
        };
    }

    /// <summary>
    /// Multiply two 4x4 column-major matrices: result = A * B.
    /// </summary>
    private static double[] MultiplyMatrix4x4(double[] a, double[] b)
    {
        var r = new double[16];
        for (int col = 0; col < 4; col++)
        for (int row = 0; row < 4; row++)
            r[col * 4 + row] = a[row] * b[col * 4] + a[4 + row] * b[col * 4 + 1]
                              + a[8 + row] * b[col * 4 + 2] + a[12 + row] * b[col * 4 + 3];
        return r;
    }

}
