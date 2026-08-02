// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using RichData = DocumentFormat.OpenXml.Office2019.Excel.RichData;

namespace OfficeCli.Handlers;

/// <summary>
/// In-cell images ("Place in Cell", Excel 365 richValue storage). The image is
/// the CELL'S VALUE — not a floating drawing: the cell is stored as
/// <c>&lt;c t="e" vm="N"&gt;&lt;v&gt;#VALUE!&lt;/v&gt;&lt;/c&gt;</c> where
/// <c>vm</c> points through <c>xl/metadata.xml</c> (XLRICHVALUE valueMetadata)
/// into the <c>xl/richData</c> part family:
///
///   rdrichvalue.xml          rv records — [localImageRelIndex, CalcOrigin(, alt)]
///   rdrichvaluestructure.xml s t="_localImage" key layout the rv values follow
///   richValueRel.xml         rel-index → r:id indirection into its own .rels
///   rdRichValueTypes.xml     static reserved-key flags blob
///
/// Cell-level surface (NOT the picture element): `set /Sheet/A2 --prop
/// image=photo.png` mirrors the hyperlink precedent — full picture element in
/// the drawing layer, value-semantic variant as a cell prop (like `link=`).
/// Down-level Excel shows the #VALUE! fallback; Excel 365 shows the image.
/// </summary>
public partial class ExcelHandler
{
    private const string RichValueRelRelType = "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel";
    private const string RichValueRelContentType = "application/vnd.ms-excel.richvaluerel+xml";
    private static readonly XNamespace RichDataNs = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
    private static readonly XNamespace RichValueRelNs = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel";
    private static readonly XNamespace OoxmlRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string XlRichValueExtUri = "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}";
    private const string LocalImageKey = "_rvRel:LocalImageIdentifier";

    // Reserved-key flags blob, captured verbatim from an Excel 365 in-cell
    // image workbook. Static for every workbook.
    private const string RdRichValueTypesXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
        "<rvTypesInfo xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2\" " +
        "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x\" " +
        "xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
        "<global><keyFlags>" +
        "<key name=\"_Self\"><flag name=\"ExcludeFromFile\" value=\"1\"/><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>" +
        "<key name=\"_DisplayString\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>" +
        "<key name=\"_Flags\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>" +
        "<key name=\"_Format\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>" +
        "<key name=\"_SubLabel\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>" +
        "<key name=\"_Attribution\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>" +
        "<key name=\"_Icon\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>" +
        "<key name=\"_Display\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>" +
        "<key name=\"_CanonicalPropertyNames\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>" +
        "<key name=\"_ClassificationId\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>" +
        "</keyFlags></global></rvTypesInfo>";

    internal readonly record struct InCellImageInfo(
        string ContentType, long FileSize, string? Alt, string RelId, OpenXmlPart Part);

    // ==================== write ====================

    /// <summary>
    /// Make <paramref name="cell"/> an in-cell image cell. Replaces any prior
    /// value/formula. <paramref name="alt"/> lands in the rv "Text" key
    /// (Excel's alt-text for in-cell images).
    /// </summary>
    private void SetInCellImage(Cell cell, string src, string? alt)
    {
        var wb = _doc.WorkbookPart ?? throw new InvalidOperationException("Workbook not found");
        var (imgStream, imgPartType) = Core.ImageSource.Resolve(src);
        using var dispose = imgStream;
        if (imgPartType == ImagePartType.Svg)
            throw new ArgumentException(
                "In-cell images do not support SVG (Excel's _localImage richValue takes raster formats only). Convert to PNG first, or use `add --type picture` for a floating SVG.");

        // 1) richValueRel part (no typed SDK class — ExtendedPart by rel type)
        //    + the image part hanging off its .rels.
        var relPart = wb.Parts.Select(p => p.OpenXmlPart).OfType<ExtendedPart>()
            .FirstOrDefault(p => p.RelationshipType == RichValueRelRelType);
        XDocument relDoc;
        if (relPart == null)
        {
            relPart = wb.AddExtendedPart(RichValueRelRelType, RichValueRelContentType, "xml");
            relDoc = new XDocument(new XElement(RichValueRelNs + "richValueRels",
                new XAttribute(XNamespace.Xmlns + "r", OoxmlRelNs.NamespaceName)));
        }
        else
        {
            relDoc = LoadPartXml(relPart);
        }
        var imagePart = relPart.AddNewPart<ImagePart>(imgPartType.ContentType);
        imagePart.FeedData(imgStream);
        var relRoot = relDoc.Root!;
        relRoot.Add(new XElement(RichValueRelNs + "rel",
            new XAttribute(OoxmlRelNs + "id", relPart.GetIdOfPart(imagePart))));
        var relIndex = relRoot.Elements(RichValueRelNs + "rel").Count() - 1;
        SavePartXml(relPart, relDoc);

        // 2) structure record: _localImage with the exact key layout our rv
        //    values will follow. Reuse an existing matching structure.
        var keys = alt != null
            ? new[] { (LocalImageKey, "i"), ("CalcOrigin", "i"), ("Text", "s") }
            : new[] { (LocalImageKey, "i"), ("CalcOrigin", "i") };
        var structPart = wb.GetPartsOfType<RdRichValueStructurePart>().FirstOrDefault();
        XDocument structDoc;
        if (structPart == null)
        {
            structPart = wb.AddNewPart<RdRichValueStructurePart>();
            structDoc = new XDocument(new XElement(RichDataNs + "rvStructures", new XAttribute("count", 0)));
        }
        else
        {
            structDoc = LoadPartXml(structPart);
        }
        var structures = structDoc.Root!.Elements(RichDataNs + "s").ToList();
        int structIndex = structures.FindIndex(s =>
            (string?)s.Attribute("t") == "_localImage"
            && s.Elements(RichDataNs + "k").Select(k => ((string?)k.Attribute("n"), (string?)k.Attribute("t")))
                .SequenceEqual(keys.Select(k => ((string?)k.Item1, (string?)k.Item2))));
        if (structIndex < 0)
        {
            structDoc.Root!.Add(new XElement(RichDataNs + "s", new XAttribute("t", "_localImage"),
                keys.Select(k => new XElement(RichDataNs + "k",
                    new XAttribute("n", k.Item1), new XAttribute("t", k.Item2)))));
            structIndex = structures.Count;
        }
        structDoc.Root!.SetAttributeValue("count", structDoc.Root!.Elements(RichDataNs + "s").Count());
        SavePartXml(structPart, structDoc);

        // 3) rv record. CalcOrigin=5 mirrors what Excel/xlsxwriter write for a
        //    user-inserted local image.
        var rvPart = wb.GetPartsOfType<RdRichValuePart>().FirstOrDefault();
        XDocument rvDoc;
        if (rvPart == null)
        {
            rvPart = wb.AddNewPart<RdRichValuePart>();
            rvDoc = new XDocument(new XElement(RichDataNs + "rvData", new XAttribute("count", 0)));
        }
        else
        {
            rvDoc = LoadPartXml(rvPart);
        }
        var values = new List<string> { relIndex.ToString(), "5" };
        if (alt != null) values.Add(alt);
        rvDoc.Root!.Add(new XElement(RichDataNs + "rv", new XAttribute("s", structIndex),
            values.Select(v => new XElement(RichDataNs + "v", v))));
        var rvIndex = rvDoc.Root!.Elements(RichDataNs + "rv").Count() - 1;
        rvDoc.Root!.SetAttributeValue("count", rvIndex + 1);
        SavePartXml(rvPart, rvDoc);

        // 4) static reserved-key flags part
        if (!wb.GetPartsOfType<RdRichValueTypesPart>().Any())
        {
            var typesPart = wb.AddNewPart<RdRichValueTypesPart>();
            using var ts = typesPart.GetStream(FileMode.Create);
            using var tw = new StreamWriter(ts);
            tw.Write(RdRichValueTypesXml);
        }

        // 5) metadata.xml XLRICHVALUE record → vm index
        var vm = EnsureXlRichValueMetadata(rvIndex);

        // 6) the cell itself: error-typed #VALUE! fallback + vm pointer
        cell.CellFormula = null;
        cell.RemoveAllChildren<InlineString>();
        cell.DataType = new EnumValue<CellValues>(CellValues.Error);
        cell.CellValue = new CellValue("#VALUE!");
        cell.ValueMetaIndex = vm;
    }

    /// <summary>
    /// Clear the in-cell image: drop the vm pointer and the #VALUE! fallback.
    /// The rv/rel/media records stay behind as unreferenced garbage — same
    /// stance Excel itself takes on delete (vacuuming richData is a rewrite
    /// of every vm index in the workbook).
    /// </summary>
    private static void RemoveInCellImage(Cell cell)
    {
        if (cell.ValueMetaIndex == null) return;
        cell.ValueMetaIndex = null;
        if (cell.DataType?.Value == CellValues.Error && cell.CellValue?.Text == "#VALUE!")
        {
            cell.CellValue = null;
            cell.DataType = null;
        }
    }

    /// <summary>
    /// Ensure xl/metadata.xml carries the XLRICHVALUE metadataType, a
    /// futureMetadata bk pointing at <paramref name="rvIndex"/>, and a
    /// valueMetadata bk referencing both. Returns the 1-based vm index for the
    /// cell. Merges into an existing metadata part (XLDAPR dynamic-array
    /// records coexist — cellMetadata and valueMetadata are separate lists).
    /// </summary>
    private uint EnsureXlRichValueMetadata(int rvIndex)
    {
        var wb = _doc.WorkbookPart ?? throw new InvalidOperationException("Workbook not found");
        var part = wb.GetPartsOfType<CellMetadataPart>().FirstOrDefault()
            ?? wb.AddNewPart<CellMetadataPart>();
        var md = part.Metadata ??= new Metadata();

        // metadataTypes / XLRICHVALUE type entry (1-based index)
        var types = md.GetFirstChild<MetadataTypes>();
        if (types == null)
        {
            types = new MetadataTypes();
            md.InsertAt(types, 0);
        }
        var typeList = types.Elements<MetadataType>().ToList();
        int typeIndex = typeList.FindIndex(t => t.Name?.Value == "XLRICHVALUE") + 1;
        if (typeIndex == 0)
        {
            types.AppendChild(new MetadataType
            {
                Name = "XLRICHVALUE",
                MinSupportedVersion = 120000U,
                Copy = true, PasteAll = true, PasteValues = true, Merge = true,
                SplitFirst = true, RowColumnShift = true, ClearFormats = true,
                ClearComments = true, Assign = true, Coerce = true
            });
            typeIndex = typeList.Count + 1;
        }
        types.Count = (uint)types.Elements<MetadataType>().Count();

        // futureMetadata name="XLRICHVALUE" bk → xlrd:rvb i=rvIndex
        var fm = md.Elements<FutureMetadata>().FirstOrDefault(f => f.Name?.Value == "XLRICHVALUE");
        if (fm == null)
        {
            fm = new FutureMetadata { Name = "XLRICHVALUE" };
            // Schema order: futureMetadata precedes cellMetadata/valueMetadata.
            var anchor = (OpenXmlElement?)md.GetFirstChild<CellMetadata>() ?? md.GetFirstChild<ValueMetadata>();
            if (anchor != null) md.InsertBefore(fm, anchor);
            else md.AppendChild(fm);
        }
        fm.AppendChild(new FutureMetadataBlock(new ExtensionList(
            new Extension(new RichData.RichValueBlock { I = (uint)rvIndex }) { Uri = XlRichValueExtUri })));
        var fmIndex = fm.Elements<FutureMetadataBlock>().Count() - 1;
        fm.Count = (uint)(fmIndex + 1);

        // valueMetadata bk rc t=typeIndex v=fmIndex → cell vm (1-based)
        var vmEl = md.GetFirstChild<ValueMetadata>();
        if (vmEl == null)
        {
            vmEl = new ValueMetadata();
            md.AppendChild(vmEl);
        }
        vmEl.AppendChild(new MetadataBlock(new MetadataRecord
        {
            TypeIndex = (uint)typeIndex,
            Val = (uint)fmIndex
        }));
        var vm = (uint)vmEl.Elements<MetadataBlock>().Count();
        vmEl.Count = vm;
        return vm;
    }

    // ==================== read ====================

    /// <summary>
    /// Resolve a cell's <c>vm</c> pointer to an in-cell image, walking
    /// metadata.xml → rdrichvalue → rdrichvaluestructure → richValueRel →
    /// image part. Defensive at every hop — a broken chain (foreign producer,
    /// hand-edited file) reports false, never throws.
    /// </summary>
    private bool TryGetInCellImage(Cell cell, out InCellImageInfo info)
    {
        info = default;
        var wb = _doc.WorkbookPart;
        if (wb == null || cell.ValueMetaIndex?.Value is not { } vm || vm < 1) return false;
        try
        {
            var md = wb.GetPartsOfType<CellMetadataPart>().FirstOrDefault()?.Metadata;
            var rec = md?.GetFirstChild<ValueMetadata>()?.Elements<MetadataBlock>()
                .ElementAtOrDefault((int)vm - 1)?.GetFirstChild<MetadataRecord>();
            if (md == null || rec?.TypeIndex?.Value is not { } t || rec.Val?.Value is not { } v) return false;
            var mdType = md.GetFirstChild<MetadataTypes>()?.Elements<MetadataType>().ElementAtOrDefault((int)t - 1);
            if (mdType?.Name?.Value != "XLRICHVALUE") return false;
            // xlrd:rvb is not reliably typed on reload (unknown element under
            // x:ext) — match by namespace + local name instead of class.
            var rvbEl = md.Elements<FutureMetadata>().FirstOrDefault(f => f.Name?.Value == "XLRICHVALUE")
                ?.Elements<FutureMetadataBlock>().ElementAtOrDefault((int)v)
                ?.Descendants().FirstOrDefault(e =>
                    e.LocalName == "rvb" && e.NamespaceUri == RichDataNs.NamespaceName);
            if (rvbEl == null || !uint.TryParse(
                    rvbEl.GetAttributes().FirstOrDefault(a => a.LocalName == "i").Value, out var rvIndex))
                return false;

            var rvPart = wb.GetPartsOfType<RdRichValuePart>().FirstOrDefault();
            var structPart = wb.GetPartsOfType<RdRichValueStructurePart>().FirstOrDefault();
            if (rvPart == null || structPart == null) return false;
            var rv = LoadPartXml(rvPart).Root?.Elements(RichDataNs + "rv").ElementAtOrDefault((int)rvIndex);
            if (rv == null || !int.TryParse((string?)rv.Attribute("s"), out var sIndex)) return false;
            var st = LoadPartXml(structPart).Root?.Elements(RichDataNs + "s").ElementAtOrDefault(sIndex);
            if (st == null || (string?)st.Attribute("t") != "_localImage") return false;
            var keyNames = st.Elements(RichDataNs + "k").Select(k => (string?)k.Attribute("n")).ToList();
            var rvValues = rv.Elements(RichDataNs + "v").Select(x => x.Value).ToList();
            var imgKey = keyNames.IndexOf(LocalImageKey);
            if (imgKey < 0 || imgKey >= rvValues.Count || !int.TryParse(rvValues[imgKey], out var relIndex)) return false;
            var textKey = keyNames.IndexOf("Text");
            var alt = textKey >= 0 && textKey < rvValues.Count ? rvValues[textKey] : null;

            var relPart = wb.Parts.Select(p => p.OpenXmlPart).OfType<ExtendedPart>()
                .FirstOrDefault(p => p.RelationshipType == RichValueRelRelType);
            var relId = relPart == null ? null
                : (string?)LoadPartXml(relPart).Root?.Elements(RichValueRelNs + "rel")
                    .ElementAtOrDefault(relIndex)?.Attribute(OoxmlRelNs + "id");
            if (relPart == null || relId == null) return false;
            // On a freshly-opened package the media part under the ExtendedPart
            // is NOT typed ImagePart (unknown parent → untyped children) —
            // probe by content-type, not by part class.
            var imgPart = relPart.GetPartById(relId);
            if (!imgPart.ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) return false;

            long size;
            using (var s = imgPart.GetStream(FileMode.Open, FileAccess.Read)) size = s.Length;
            info = new InCellImageInfo(imgPart.ContentType, size, alt, relId, imgPart);
            return true;
        }
        catch (Exception)
        {
            return false;
        }
    }

    private static XDocument LoadPartXml(OpenXmlPart part)
    {
        using var s = part.GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(s);
    }

    private static void SavePartXml(OpenXmlPart part, XDocument doc)
    {
        using var s = part.GetStream(FileMode.Create, FileAccess.Write);
        doc.Save(s);
    }
}
