// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using Vml = DocumentFormat.OpenXml.Vml;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{

    /// <summary>
    /// Ensure Columns exists in SectionProperties in correct schema order.
    /// CT_SectPr child sequence is strict (cols is rank 11, after lnNumType
    /// rank 9). Delegate to InsertSectPrChildInOrder so later children
    /// (lnNumType, pgNumType, …) that the old hardcoded "after PageMargin"
    /// branch landed cols ahead of are now sorted correctly.
    /// </summary>
    private static Columns EnsureColumns(SectionProperties sectPr)
    {
        var existing = sectPr.GetFirstChild<Columns>();
        if (existing != null) return existing;

        var cols = new Columns();
        InsertSectPrChildInOrder(sectPr, cols);
        return cols;
    }

    /// <summary>
    /// Ensure PageSize exists in SectionProperties in correct schema order.
    /// CT_SectPr's child sequence is strict (headerReference*, footerReference*,
    /// footnotePr, endnotePr, type, pgSz, pgMar, …), so a newly created pgSz is
    /// placed by rank via InsertSectPrChildInOrder rather than by a hand-rolled
    /// position. The previous code fell back to PrependChild when the section
    /// had neither a w:type nor a header/footer reference, which put pgSz ahead
    /// of an existing footnotePr/endnotePr and made the document fail
    /// validation.
    /// </summary>
    private static PageSize EnsureSectPrPageSize(SectionProperties sectPr)
    {
        var existing = sectPr.GetFirstChild<PageSize>();
        if (existing != null) return existing;

        var ps = new PageSize();
        InsertSectPrChildInOrder(sectPr, ps);
        return ps;
    }

    /// <summary>
    /// Ensure PageMargin exists in SectionProperties in correct schema order.
    /// Same rank-based placement as <see cref="EnsureSectPrPageSize"/> — pgMar
    /// follows headerReference/footerReference/footnotePr/endnotePr/type/pgSz.
    /// </summary>
    private static PageMargin EnsureSectPrPageMargin(SectionProperties sectPr)
    {
        var existing = sectPr.GetFirstChild<PageMargin>();
        if (existing != null) return existing;

        var pm = WordPageDefaults.CreatePageMargin();
        InsertSectPrChildInOrder(sectPr, pm);
        return pm;
    }

    // ==================== sectPr schema-order insertion ====================

    /// <summary>
    /// Canonical CT_SectPr child schema order (subset, in document order):
    ///   headerReference*, footerReference*, footnotePr, endnotePr, type, pgSz,
    ///   pgMar, paperSrc, pgBorders, lnNumType, pgNumType, cols, formProt,
    ///   vAlign, noEndnote, titlePg, textDirection, bidi, rtlGutter, docGrid,
    ///   printerSettings, sectPrChange.
    /// Used to map a child element to its schema-order rank for ordered insertion.
    /// </summary>
    private static int SectPrChildOrder(OpenXmlElement el) => el switch
    {
        HeaderReference => 0,
        FooterReference => 1,
        FootnoteProperties => 2,
        EndnoteProperties => 3,
        SectionType => 4,
        PageSize => 5,
        PageMargin => 6,
        PaperSource => 7,
        PageBorders => 8,
        LineNumberType => 9,
        PageNumberType => 10,
        Columns => 11,
        FormProtection => 12,
        VerticalTextAlignmentOnPage => 13,
        NoEndnote => 14,
        TitlePage => 15,
        TextDirection => 16,
        BiDi => 17,
        GutterOnRight => 18,
        DocGrid => 19,
        PrinterSettingsReference => 20,
        SectionPropertiesChange => 21,
        _ => 99,
    };

    /// <summary>
    /// Insert <paramref name="newChild"/> into <paramref name="sectPr"/> at the
    /// position dictated by CT_SectPr schema order. Required for elements like
    /// &lt;w:bidi/&gt; which Word's schema validator rejects when appended after
    /// &lt;w:docGrid/&gt;. Mirrors the InsertRunPropInSchemaOrder pattern used
    /// for run properties.
    /// </summary>
    private static void InsertSectPrChildInOrder(SectionProperties sectPr, OpenXmlElement newChild)
    {
        var newRank = SectPrChildOrder(newChild);
        OpenXmlElement? successor = null;
        foreach (var child in sectPr.ChildElements)
        {
            if (SectPrChildOrder(child) > newRank)
            {
                successor = child;
                break;
            }
        }
        if (successor != null)
            successor.InsertBeforeSelf(newChild);
        else
            sectPr.AppendChild(newChild);
    }

    /// <summary>
    /// CT_TblPrBase schema order:
    ///   tblStyle, tblpPr, tblOverlap, bidiVisual, tblStyleRowBandSize,
    ///   tblStyleColBandSize, tblW, jc, tblCellSpacing, tblInd, tblBorders,
    ///   shd, tblLayout, tblCellMar, tblLook, tblCaption, tblDescription,
    ///   tblPrChange.
    /// </summary>
    private static int TblPrChildOrder(OpenXmlElement el) => el switch
    {
        TableStyle => 0,
        TablePositionProperties => 1,
        TableOverlap => 2,
        BiDiVisual => 3,
        TableStyleRowBandSize => 4,
        TableStyleColumnBandSize => 5,
        TableWidth => 6,
        TableJustification => 7,
        TableCellSpacing => 8,
        TableIndentation => 9,
        TableBorders => 10,
        Shading => 11,
        TableLayout => 12,
        TableCellMarginDefault => 13,
        TableLook => 14,
        TableCaption => 15,
        TableDescription => 16,
        TablePropertiesChange => 17,
        _ => 99,
    };

    /// <summary>
    /// Insert <paramref name="newChild"/> into <paramref name="tblPr"/> at the
    /// position dictated by CT_TblPrBase schema order. Required for elements
    /// like &lt;w:bidiVisual/&gt; which Word's schema validator rejects when
    /// appended after &lt;w:tblBorders/&gt;.
    /// </summary>
    private static void InsertTblPrChildInOrder(TableProperties tblPr, OpenXmlElement newChild)
    {
        var newRank = TblPrChildOrder(newChild);
        OpenXmlElement? successor = null;
        foreach (var child in tblPr.ChildElements)
        {
            if (TblPrChildOrder(child) > newRank)
            {
                successor = child;
                break;
            }
        }
        if (successor != null)
            successor.InsertBeforeSelf(newChild);
        else
            tblPr.AppendChild(newChild);
    }

    /// <summary>
    /// Get-or-create <w:tblCellMar/> on the given tblPr in CT_TblPrBase schema
    /// order. Prevents the "argv-order produces schema-invalid tblCellMar
    /// position" class of bug — see InsertTblPrChildInOrder docstring.
    /// </summary>
    private static TableCellMarginDefault EnsureTableCellMarginDefault(TableProperties tblPr)
    {
        var cm = tblPr.TableCellMarginDefault;
        if (cm == null)
        {
            cm = new TableCellMarginDefault();
            InsertTblPrChildInOrder(tblPr, cm);
        }
        return cm;
    }

    /// <summary>
    /// dump→batch: cells wrapped by a CELL-LEVEL content control — a
    /// <c>&lt;w:sdt&gt;</c> that is a direct <c>&lt;w:tr&gt;</c> child whose
    /// sdtContent holds the <c>&lt;w:tc&gt;</c> (Word's dropdown-in-a-cell
    /// shape). Navigation flattens these to plain cells for the typed emit, so
    /// EmitTable patches them back via raw-set replace. Returns
    /// (1-based cell ordinal counting flattened cells, verbatim sdt XML) for
    /// every single-cell wrapper in the row; multi-cell wrappers are skipped
    /// (the existing flatten behavior stands for those).
    /// </summary>
    internal List<(int CellOrdinal, string SdtXml)> GetSdtWrappedCellsOfRow(string rowPath)
    {
        var result = new List<(int, string)>();
        OpenXmlElement? element;
        try { element = NavigateToElement(ParsePath(rowPath)); }
        catch { return result; }
        if (element is not TableRow row) return result;

        int ordinal = 0;
        foreach (var child in row.ChildElements)
        {
            if (child is TableCell)
            {
                ordinal++;
            }
            else if (child is SdtElement sdt)
            {
                var wrappedCells = sdt.Descendants<TableCell>()
                    .Where(tc => ReferenceEquals(tc.Ancestors<TableRow>().FirstOrDefault(), row))
                    .ToList();
                if (wrappedCells.Count == 1)
                {
                    ordinal++;
                    result.Add((ordinal, sdt.OuterXml));
                }
                else
                {
                    ordinal += wrappedCells.Count;
                }
            }
        }
        return result;
    }

    /// <summary>
    /// All cells of a row in document order, INCLUDING cells wrapped by a
    /// cell-level content control (<c>&lt;w:sdt&gt;&lt;w:sdtContent&gt;&lt;w:tc&gt;</c>
    /// as a direct <c>&lt;w:tr&gt;</c> child — Word's dropdown-bound-cell shape).
    /// <c>Elements&lt;TableCell&gt;()</c> sees only direct children, so a row
    /// with wrapped cells under-counted its columns: paths mis-resolved,
    /// Get/Query reported missing cells, and dump→batch rebuilt the row short.
    /// </summary>
    internal static List<TableCell> GetRowCellsFlattened(TableRow row)
    {
        var cells = new List<TableCell>();
        foreach (var child in row.ChildElements)
        {
            if (child is TableCell tc)
                cells.Add(tc);
            else if (child is SdtElement sdt)
                cells.AddRange(sdt.Descendants<TableCell>()
                    .Where(c => ReferenceEquals(c.Ancestors<TableRow>().FirstOrDefault(), row)));
        }
        return cells;
    }

    /// <summary>
    /// dump→batch: rows wrapped by a ROW-LEVEL content control — a
    /// <c>&lt;w:sdt&gt;</c> (SdtRow) that is a direct <c>&lt;w:tbl&gt;</c> child
    /// whose sdtContent holds the <c>&lt;w:tr&gt;</c> (Word's locked-row shape,
    /// the table analog of <see cref="GetSdtWrappedCellsOfRow"/>). Navigation
    /// flattens these to plain rows for the typed emit, so EmitTable patches the
    /// wrapper (and its lock) back via raw-set replace. Returns (1-based row
    /// ordinal counting flattened rows, verbatim sdt XML) for every single-row
    /// wrapper in the table; multi-row wrappers are skipped (the flatten
    /// behavior stands — those rows survive un-wrapped).
    /// </summary>
    internal List<(int RowOrdinal, string SdtXml)> GetSdtWrappedRowsOfTable(string tablePath)
    {
        var result = new List<(int, string)>();
        OpenXmlElement? element;
        try { element = NavigateToElement(ParsePath(tablePath)); }
        catch { return result; }
        if (element is not Table table) return result;

        int ordinal = 0;
        foreach (var child in table.ChildElements)
        {
            if (child is TableRow)
            {
                ordinal++;
            }
            else if (child is SdtElement sdt)
            {
                var wrappedRows = sdt.Descendants<TableRow>()
                    .Where(r => ReferenceEquals(
                        r.Ancestors<Table>().FirstOrDefault(), table))
                    .ToList();
                if (wrappedRows.Count == 1)
                {
                    ordinal++;
                    result.Add((ordinal, sdt.OuterXml));
                }
                else
                {
                    ordinal += wrappedRows.Count;
                }
            }
        }
        return result;
    }

    /// <summary>
    /// All rows of a table in document order, INCLUDING rows wrapped by a
    /// row-level content control (<c>&lt;w:sdt&gt;&lt;w:sdtContent&gt;&lt;w:tr&gt;</c>
    /// as a direct <c>&lt;w:tbl&gt;</c> child — Word's locked-row shape, the
    /// table analog of <see cref="GetRowCellsFlattened"/>). CT_Tbl's content
    /// model permits an SDT (SdtRow) around one or more rows; locked
    /// government forms use it to make an entire row read-only.
    /// <c>Elements&lt;TableRow&gt;()</c> sees only direct children, so a table
    /// with SDT-wrapped rows under-counted its rows: Get/Query/dump rebuilt
    /// the table short, dropping the wrapped rows, their content and their
    /// locks. Mirrors the cell-flatten contract.
    /// </summary>
    internal static List<TableRow> GetTableRowsFlattened(Table table)
    {
        var rows = new List<TableRow>();
        foreach (var child in table.ChildElements)
        {
            if (child is TableRow tr)
                rows.Add(tr);
            else if (child is SdtElement sdt)
                rows.AddRange(sdt.Descendants<TableRow>()
                    .Where(r => ReferenceEquals(
                        r.Ancestors<Table>().FirstOrDefault(), table)));
        }
        return rows;
    }
}
