// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    private static string HeaderFooterTypeName(HeaderFooterValues v)
    {
        if (v == HeaderFooterValues.First) return "first";
        if (v == HeaderFooterValues.Even) return "even";
        return "default";
    }

    // BUG-R5B(BUG1): on dump→batch replay, a section whose break is encoded as
    // an inline w:pPr/w:sectPr can be re-materialized verbatim by a different
    // emit path (EmitCrossParagraphFieldMember raw-sets the whole <w:p>,
    // including its <w:headerReference>/<w:footerReference>) while
    // EmitHeadersFooters independently emits `add header|footer parent=/section[N]`
    // for the same section. The reference copied in by the raw-set points at an
    // r:id that has no matching part on the rebuilt doc (the part was never
    // recreated) — an orphan. AddHeader/AddFooter's duplicate-of-type guard saw
    // the orphan and threw "already exists", aborting the add and losing the
    // footer/header CONTENT for that section.
    //
    // Distinguish a genuine duplicate (the existing reference resolves to a real
    // part — a true user error) from an orphaned reference (its r:id resolves to
    // nothing — round-trip leftover). For the orphan we repurpose the add:
    // create the part, rewire the reference's r:id to it, and skip prepending a
    // second reference. Returns the existing reference to repoint, or null when
    // none of this type / it is a genuine non-orphan duplicate.
    private HeaderReference? FindOrphanedHeaderReference(SectionProperties? sectPr, HeaderFooterValues type)
    {
        var mainPart = _doc.MainDocumentPart;
        if (sectPr == null || mainPart == null) return null;
        foreach (var hr in sectPr.Elements<HeaderReference>())
        {
            if (hr.Type == null || hr.Type.Value != type) continue;
            if (string.IsNullOrEmpty(hr.Id?.Value)) return hr; // missing r:id ⇒ orphan
            if (!ReferenceResolvesToPart(mainPart, hr.Id!.Value!)) return hr;
        }
        return null;
    }

    private FooterReference? FindOrphanedFooterReference(SectionProperties? sectPr, HeaderFooterValues type)
    {
        var mainPart = _doc.MainDocumentPart;
        if (sectPr == null || mainPart == null) return null;
        foreach (var fr in sectPr.Elements<FooterReference>())
        {
            if (fr.Type == null || fr.Type.Value != type) continue;
            if (string.IsNullOrEmpty(fr.Id?.Value)) return fr; // missing r:id ⇒ orphan
            if (!ReferenceResolvesToPart(mainPart, fr.Id!.Value!)) return fr;
        }
        return null;
    }

    private static bool ReferenceResolvesToPart(MainDocumentPart mainPart, string relationshipId)
    {
        try
        {
            // GetPartById throws (KeyNotFoundException / ArgumentOutOfRange) when
            // the relationship id is unknown to the part — that's exactly the
            // orphan case we want to detect.
            return mainPart.GetPartById(relationshipId) != null;
        }
        catch
        {
            return false;
        }
    }

    private string AddSection(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        // BUG-DUMP-R31-1: a childless mid-document <w:sectPr/> round-trip. The
        // dump emits `empty=true` (with no `type` and no geometry keys) when the
        // source's inline sectPr had NO children — the author deferred the break
        // kind and page geometry to Word's defaults, exactly like an empty FINAL
        // body sectPr. In that case build a bare sectPr: no fabricated
        // <w:type>, no default pgSz/pgMar. A mid sectPr WITH a real type or
        // geometry forwards those keys and takes the normal path below.
        bool bareSection = properties.TryGetValue("empty", out var emptyFlag) && IsTruthy(emptyFlag);

        // BUG-DUMP-SECT-TYPEINJECT: a source sectPr that OMITTED <w:type>
        // (deferring to the OOXML default, nextPage) must round-trip WITHOUT
        // one — fabricating <w:type w:val="nextPage"/> turns an implicit break
        // into an EXPLICIT section page break that Word honors as a hard page
        // boundary → +1 page (NAR1 27→28). The `empty` flag only covered a
        // fully childless sectPr; a sectPr with geometry (pgSz/docGrid) but no
        // <w:type> fell through to the default-stamp. The dump signals this with
        // an explicit `notype=true` (emitted ONLY when the source sectPr had
        // children but no <w:type>); on that signal we skip the stamp. An
        // INTERACTIVE `add section` (no notype flag) keeps the default-stamp so
        // the break renders — that path relies on the explicit <w:type>.
        bool noType = properties.TryGetValue("notype", out var noTypeFlag) && IsTruthy(noTypeFlag);

        // Section break: adds SectionProperties to the last paragraph before the break point
        var breakType = properties.GetValueOrDefault("type", "nextPage").ToLowerInvariant();
        var sectType = breakType switch
        {
            "nextpage" or "next" => SectionMarkValues.NextPage,
            "continuous" => SectionMarkValues.Continuous,
            "evenpage" or "even" => SectionMarkValues.EvenPage,
            "oddpage" or "odd" => SectionMarkValues.OddPage,
            // R7-fuzz-3: nextColumn is a valid OOXML SectionMarkValues
            // member used to start a new column inside multi-column layouts
            // — the whitelist had skipped it, surfacing as a hard reject.
            "nextcolumn" or "column" => SectionMarkValues.NextColumn,
            _ => throw new ArgumentException($"Invalid section break type: '{breakType}'. Valid values: nextPage, continuous, evenPage, oddPage, nextColumn.")
        };

        // Create a paragraph with section properties to mark the break
        var sectPara = new Paragraph();
        var sectPProps = new ParagraphProperties();
        var sectPr = new SectionProperties();
        // CONSISTENCY(sectpr-schema-order): always route sectPr child inserts
        // through InsertSectPrChildInOrder. Mixing raw AppendChild with later
        // schema-aware inserts produced argv-order-dependent invalid OOXML
        // when e.g. `columns` (rank 11) was appended before `lineNumbers`
        // (rank 9).
        // BUG-DUMP-R31-1: skip the SectionType insert for a bare round-tripped
        // section — the source had no <w:type> and fabricating one (the OOXML
        // default is nextPage, so omitting it is equivalent) breaks faithful
        // round-trip. A non-bare section always stamps the type.
        if (!bareSection && !noType)
            InsertSectPrChildInOrder(sectPr, new SectionType { Val = sectType });

        // Ensure body-level sectPr has pgSz/pgMar (fix for docs created by older versions)
        var bodySectPr = body.GetFirstChild<SectionProperties>();
        if (bodySectPr != null && bodySectPr.GetFirstChild<PageSize>() == null)
        {
            bodySectPr.InsertBefore(new PageSize { Width = WordPageDefaults.A4WidthTwips, Height = WordPageDefaults.A4HeightTwips },
                bodySectPr.GetFirstChild<DocGrid>());
        }
        if (bodySectPr != null && bodySectPr.GetFirstChild<PageMargin>() == null)
        {
            bodySectPr.InsertBefore(new PageMargin { Top = 1440, Right = 1800U, Bottom = 1440, Left = 1800U },
                bodySectPr.GetFirstChild<DocGrid>());
        }

        // BUG-DUMP-R31-1: a bare round-tripped section carries no page geometry.
        // The source's childless sectPr had neither pgSz nor pgMar, so stamping
        // the body section's (or A4 default) geometry would inject child elements
        // the source never had — exactly the mid-doc fabrication this fixes.
        // Skip straight to attaching the empty sectPr; the explicit-override
        // blocks below still don't fire (a bare section forwards no override
        // keys), so the rebuilt <w:sectPr/> stays childless.
        if (bareSection)
            return AttachSectionBreakParagraph(body, parentPath, index, sectPara, sectPProps, sectPr, properties);

        // Copy page size/margins from document section, or use A4 defaults.
        // BUG-DUMP-SECT-ORIENT: do NOT inherit the body section's w:orient. The
        // body section is the LAST (often landscape) section; a mid-document
        // section created without an explicit `orientation` prop must derive its
        // orientation from its own width/height, never from the trailing
        // section's stale flag. Inheriting landscape onto a portrait W<H pgSz
        // produced a contradictory `<w:pgSz w:w="11906" w:h="16838"
        // w:orient="landscape"/>` on dump→batch replay. Orient is set below
        // (explicit prop) or derived from final W/H after all overrides land.
        // Copy values, not the UInt32Value references — sharing them would
        // alias the new sectPr's dimensions onto the source body sectPr, so a
        // later orientation swap on the new section mutates the body section
        // simultaneously (and any subsequent section that inherits from body
        // sees the corrupted dimensions).
        var srcPageSize = bodySectPr?.GetFirstChild<PageSize>();
        InsertSectPrChildInOrder(sectPr, new PageSize
        {
            Width = srcPageSize?.Width?.Value ?? WordPageDefaults.A4WidthTwips,
            Height = srcPageSize?.Height?.Value ?? WordPageDefaults.A4HeightTwips
        });
        var srcMargin = bodySectPr?.GetFirstChild<PageMargin>();
        InsertSectPrChildInOrder(sectPr, new PageMargin
        {
            Top = srcMargin?.Top?.Value ?? 1440,
            Bottom = srcMargin?.Bottom?.Value ?? 1440,
            Left = srcMargin?.Left?.Value ?? 1800,
            Right = srcMargin?.Right?.Value ?? 1800
        });

        // Allow per-section overrides
        bool explicitWidth = properties.TryGetValue("pagewidth", out var sw) || properties.TryGetValue("pageWidth", out sw) || properties.TryGetValue("width", out sw);
        if (explicitWidth)
        {
            (EnsureSectPrChild<PageSize>(sectPr)).Width = ParseTwips(sw!);
        }
        bool explicitHeight = properties.TryGetValue("pageheight", out var sh) || properties.TryGetValue("pageHeight", out sh) || properties.TryGetValue("height", out sh);
        if (explicitHeight)
        {
            (EnsureSectPrChild<PageSize>(sectPr)).Height = ParseTwips(sh!);
        }
        if (properties.TryGetValue("orientation", out var orient))
        {
            var ps = EnsureSectPrChild<PageSize>(sectPr);
            ps.Orient = orient.ToLowerInvariant() == "landscape"
                ? PageOrientationValues.Landscape
                : PageOrientationValues.Portrait;
            // BUG-DUMP-SECT-ORIENT-SWAP: only auto-swap W/H to match the
            // orientation flag when the caller did NOT supply both explicit
            // dimensions. `add section --prop orientation=landscape` with
            // inherited portrait dims wants the swap (a convenience); but on
            // dump→batch replay the source's pgSz is authoritative — Word
            // renders w:w/w:h literally and w:orient is only a printer/UI hint.
            // A section legitimately carrying w<h with orient=landscape (e.g. an
            // 11"-wide × 15.75"-tall page) must NOT be rotated, or it becomes
            // 15.75"×11", shrinks the usable height, and reflows onto an extra
            // page on round-trip.
            bool dimsAuthoritative = explicitWidth && explicitHeight;
            if (!dimsAuthoritative)
            {
                if (ps.Orient == PageOrientationValues.Landscape && ps.Width < ps.Height)
                    (ps.Width!.Value, ps.Height!.Value) = (ps.Height.Value, ps.Width.Value);
                if (ps.Orient == PageOrientationValues.Portrait && ps.Width > ps.Height)
                    (ps.Width!.Value, ps.Height!.Value) = (ps.Height.Value, ps.Width.Value);
            }
        }
        else
        {
            // BUG-DUMP-SECT-ORIENT: no explicit orientation prop. Derive the
            // landscape flag from the section's own final W/H so the section
            // never carries a contradictory orient inherited from another
            // section. Width>Height ⇒ landscape; otherwise leave orient unset
            // (portrait is the OOXML default; omitting w:orient matches how
            // Word writes a portrait section, keeping the dump→batch byte shape).
            var ps = sectPr.GetFirstChild<PageSize>();
            if (ps?.Width?.Value is uint dw && ps.Height?.Value is uint dh && dw > dh)
                ps.Orient = PageOrientationValues.Landscape;
        }

        // Columns support: "columns=2" or "columns=2,1cm"
        if (properties.TryGetValue("columns", out var colsVal) || properties.TryGetValue("columns.count", out colsVal))
        {
            var parts = colsVal.Split(',');
            var count = (short)int.Parse(parts[0].Trim());
            // CONSISTENCY(columns-no-equalWidth-default): Set.SectionLayout
            // dropped the EqualWidth auto-stamp (round-trip preservation —
            // sources whose <w:cols> omits w:equalWidth must replay without
            // it). Mirror that on Add so `add section --prop columns=N`
            // doesn't phantom-stamp columns.equalWidth=true.
            var cols = new Columns { ColumnCount = count };
            if (parts.Length > 1)
                cols.Space = ParseTwips(parts[1].Trim()).ToString();
            InsertSectPrChildInOrder(sectPr, cols);
        }
        if (properties.TryGetValue("columns.space", out var colSpace)
            || properties.TryGetValue("columnSpace", out colSpace))
        {
            var cols = EnsureSectPrChild<Columns>(sectPr);
            cols.Space = ParseTwips(colSpace).ToString();
        }
        // columns.equalWidth — Set.SectionLayout exposes this; mirror on Add so
        // batch replay round-trips dumps that include it (Get emits it when the
        // <w:cols> has w:equalWidth set).
        if (properties.TryGetValue("columns.equalWidth", out var eqW)
            || properties.TryGetValue("columns.equalwidth", out eqW))
        {
            var cols = EnsureSectPrChild<Columns>(sectPr);
            cols.EqualWidth = IsTruthy(eqW);
        }
        // CONSISTENCY(add-set-symmetry): mirror TrySetSectionLayout's columns.separator
        // case so AddSection round-trips the vertical column separator. Without an
        // explicit handler the key falls through TypedAttributeFallback against sectPr
        // (separator lives on <w:cols>, not sectPr) and gets dropped silently.
        if (properties.TryGetValue("columns.separator", out var colSep))
        {
            var cols = EnsureSectPrChild<Columns>(sectPr);
            cols.Separator = IsTruthy(colSep);
        }
        // colWidths/colSpaces — non-equal-width column layout. Comma-separated lists
        // of twip values (or unit-qualified like "4cm,3cm"). colWidths sets per-column
        // widths; colSpaces sets per-column spacing AFTER each column (last value
        // typically 0). Both lists must have the same length and equal columns.count.
        // Writes <w:cols><w:col w:w="…" w:space="…"/>…</w:cols> and implicitly flips
        // equalWidth=false (OOXML rule: explicit widths imply non-equal layout).
        ApplySectionColumnWidthsSpaces(properties, sectPr);

        // Per-section margin overrides — mutate the PageMargin child of the
        // new sectPr (not the body sectPr). Margins use Int32Value for Top/
        // Bottom and UInt32Value for Left/Right to match the schema.
        var pm = EnsureSectPrChild<PageMargin>(sectPr);
        if (properties.TryGetValue("marginTop", out var mTop) || properties.TryGetValue("margintop", out mTop))
            pm.Top = (int)ParseTwips(mTop);
        if (properties.TryGetValue("marginBottom", out var mBot) || properties.TryGetValue("marginbottom", out mBot))
            pm.Bottom = (int)ParseTwips(mBot);
        if (properties.TryGetValue("marginLeft", out var mLeft) || properties.TryGetValue("marginleft", out mLeft))
            pm.Left = ParseTwips(mLeft);
        if (properties.TryGetValue("marginRight", out var mRight) || properties.TryGetValue("marginright", out mRight))
            pm.Right = ParseTwips(mRight);
        if (properties.TryGetValue("marginHeader", out var mHdr) || properties.TryGetValue("marginheader", out mHdr))
            pm.Header = ParseTwips(mHdr);
        if (properties.TryGetValue("marginFooter", out var mFtr) || properties.TryGetValue("marginfooter", out mFtr))
            pm.Footer = ParseTwips(mFtr);
        if (properties.TryGetValue("marginGutter", out var mGut) || properties.TryGetValue("margingutter", out mGut))
            pm.Gutter = ParseTwips(mGut);

        // Line numbering — mirrors Set parser (WordHandler.Set.SectionLayout.cs).
        // CONSISTENCY(linenumbers-countby-independent): lineNumberCountBy can
        // be passed alone (without lineNumbers) — default the restart mode to
        // continuous so the countBy isn't silently swallowed when the user
        // omits the companion key.
        bool hasLineNumbers = properties.TryGetValue("lineNumbers", out var lnVal) ||
                              properties.TryGetValue("linenumbers", out lnVal);
        bool hasCountBy = properties.TryGetValue("lineNumberCountBy", out var lnBy) ||
                          properties.TryGetValue("linenumbercountby", out lnBy);
        // BUG-DUMP-SECT-LNDIST: lnNumType/@w:distance (gutter twips) — sibling
        // attr to countBy/start. Triggers LineNumberType creation on its own so a
        // carrier that has only @distance round-trips.
        bool hasLnDist = properties.TryGetValue("lineNumberDistance", out var lnDistVal) ||
                         properties.TryGetValue("linenumberdistance", out lnDistVal);
        // BUG-DUMP-SECT-LNSTART: lnNumType/@w:start — sibling attr to countBy/
        // distance; mirror TrySetSectionLayout's lineNumberStart case so a
        // mid-document carrier round-trips its first-line-number.
        bool hasLnStart = properties.TryGetValue("lineNumberStart", out var lnStartVal) ||
                          properties.TryGetValue("linenumberstart", out lnStartVal);
        if (hasLineNumbers || hasCountBy || hasLnDist || hasLnStart)
        {
            // BUG-DUMP-SECT-LNDIST: only stamp @w:restart when the source
            // actually carried lineNumbers — fabricating restart="continuous" on
            // a countBy/distance-only carrier was the round-trip regression. The
            // SDK omits the attr entirely when Restart is left null.
            var lnType = new LineNumberType();
            if (hasLineNumbers)
                lnType.Restart = lnVal!.ToLowerInvariant() switch
                {
                    "continuous" => LineNumberRestartValues.Continuous,
                    "restartpage" or "page" => LineNumberRestartValues.NewPage,
                    "restartsection" or "section" => LineNumberRestartValues.NewSection,
                    _ => throw new ArgumentException(
                        $"Invalid lineNumbers value: '{lnVal}'. Valid values: continuous, restartPage, restartSection.")
                };
            if (hasCountBy)
            {
                var by = int.Parse(lnBy!);
                if (by > 1) lnType.CountBy = (short)by;
            }
            if (hasLnDist)
            {
                if (!int.TryParse(lnDistVal, out var lnDist) || lnDist < 0)
                    throw new ArgumentException(
                        $"Invalid lineNumberDistance value: '{lnDistVal}'. Must be a non-negative integer (twips).");
                lnType.Distance = lnDist.ToString();
            }
            if (hasLnStart)
            {
                if (!int.TryParse(lnStartVal, out var lnStart) || lnStart < 0)
                    throw new ArgumentException(
                        $"Invalid lineNumberStart value: '{lnStartVal}'. Must be a non-negative integer.");
                lnType.Start = (short)lnStart;
            }
            InsertSectPrChildInOrder(sectPr, lnType);
        }

        // Section-level RTL: <w:bidi/> in sectPr flips page direction.
        // Mirrors Set vocabulary (direction/dir/bidi). Use the schema-aware
        // inserter so the element lands at the canonical CT_SectPrBase
        // position regardless of what other children were appended above.
        if (properties.TryGetValue("direction", out var sectDir)
            || properties.TryGetValue("dir", out sectDir)
            || properties.TryGetValue("bidi", out sectDir))
        {
            if (ParseDirectionRtl(sectDir))
                InsertSectPrChildInOrder(sectPr, new BiDi());
        }

        // Section-level RTL gutter: <w:rtlGutter/> places the binding gutter
        // on the right side. Mirrors Set vocabulary (rtlgutter) and uses the
        // schema-aware inserter for canonical CT_SectPrBase order.
        // CONSISTENCY(add-set-symmetry).
        if (properties.TryGetValue("rtlGutter", out var sectRtlG)
            || properties.TryGetValue("rtlgutter", out sectRtlG))
        {
            if (IsTruthy(sectRtlG))
                InsertSectPrChildInOrder(sectPr, new GutterOnRight());
        }

        // CONSISTENCY(add-set-symmetry): mirror SetSectionLayout's titlePage /
        // pageNumFmt / pageStart handling. Schema declares these add=true so
        // the schema preflight lets them through; without explicit handling
        // here they get silently dropped on add and round-trip via Get fails.
        if (properties.TryGetValue("titlePage", out var tpVal) ||
            properties.TryGetValue("titlepage", out tpVal) ||
            properties.TryGetValue("titlePg", out tpVal) ||
            properties.TryGetValue("titlepg", out tpVal))
        {
            if (IsTruthy(tpVal))
            {
                if (sectPr.GetFirstChild<TitlePage>() == null)
                    InsertSectPrChildInOrder(sectPr, new TitlePage());
            }
        }

        if (properties.TryGetValue("pageNumFmt", out var pnfVal) ||
            properties.TryGetValue("pagenumfmt", out pnfVal) ||
            properties.TryGetValue("pageNumberFormat", out pnfVal) ||
            properties.TryGetValue("pagenumberformat", out pnfVal))
        {
            var pgNum = sectPr.GetFirstChild<PageNumberType>();
            if (pgNum == null)
            {
                pgNum = new PageNumberType();
                InsertSectPrChildInOrder(sectPr, pgNum);
            }
            pgNum.Format = ParseNumberFormat(pnfVal);
        }

        if (properties.TryGetValue("pageStart", out var psVal) ||
            properties.TryGetValue("pagestart", out psVal) ||
            properties.TryGetValue("pageNumberStart", out psVal) ||
            properties.TryGetValue("pagenumberstart", out psVal))
        {
            var startN = ParseHelpers.SafeParseInt(psVal, "pageStart");
            if (startN < 0)
                throw new ArgumentException("pageStart must be a non-negative integer.");
            var pgNum = sectPr.GetFirstChild<PageNumberType>();
            if (pgNum == null)
            {
                pgNum = new PageNumberType();
                InsertSectPrChildInOrder(sectPr, pgNum);
            }
            pgNum.Start = startN;
        }

        // BUG-DUMP-SECT-CHAPNUM: chapter-number page numbering on a mid-document
        // section carrier. CONSISTENCY(add-set-symmetry): mirror TrySetSectionLayout's
        // chapStyle/chapSep cases so `add section` round-trips them (the carrier
        // sectPr readback now forwards sectionBreak.chapStyle/chapSep).
        if (properties.TryGetValue("chapStyle", out var chapStyleVal)
            || properties.TryGetValue("chapstyle", out chapStyleVal))
        {
            if (!byte.TryParse(chapStyleVal, out var lvl) || lvl < 1 || lvl > 9)
                throw new ArgumentException(
                    $"Invalid chapStyle value: '{chapStyleVal}'. Must be 1-9 (heading level).");
            var pgNum = sectPr.GetFirstChild<PageNumberType>();
            if (pgNum == null)
            {
                pgNum = new PageNumberType();
                InsertSectPrChildInOrder(sectPr, pgNum);
            }
            pgNum.ChapterStyle = lvl;
        }
        if (properties.TryGetValue("chapSep", out var chapSepVal)
            || properties.TryGetValue("chapsep", out chapSepVal))
        {
            var pgNum = sectPr.GetFirstChild<PageNumberType>();
            if (pgNum == null)
            {
                pgNum = new PageNumberType();
                InsertSectPrChildInOrder(sectPr, pgNum);
            }
            pgNum.ChapterSeparator = chapSepVal.ToLowerInvariant() switch
            {
                "hyphen" or "-" => ChapterSeparatorValues.Hyphen,
                "period" or "." => ChapterSeparatorValues.Period,
                "colon" or ":" => ChapterSeparatorValues.Colon,
                "emdash" or "—" => ChapterSeparatorValues.EmDash,
                "endash" or "–" => ChapterSeparatorValues.EnDash,
                _ => throw new ArgumentException(
                    $"Invalid chapSep value: '{chapSepVal}'. Valid: hyphen, period, colon, emDash, enDash.")
            };
        }

        // BUG-DUMP-SECT-NOENDNOTE: <w:noEndnote/> suppresses endnote collection for
        // the section. CONSISTENCY(add-set-symmetry): mirror TrySetSectionLayout's
        // noendnote case for a mid-document carrier.
        if ((properties.TryGetValue("noEndnote", out var noEnVal)
             || properties.TryGetValue("noendnote", out noEnVal))
            && IsTruthy(noEnVal))
        {
            sectPr.RemoveAllChildren<NoEndnote>();
            InsertSectPrChildInOrder(sectPr, new NoEndnote());
        }

        // BUG-DUMP-SECT-VALIGN: vertical text alignment on the page
        // (top/center/bottom/both). CONSISTENCY(add-set-symmetry): mirror
        // TrySetSectionLayout's valign case so `add section --prop vAlign=center`
        // round-trips a mid-document section's <w:vAlign>.
        if (properties.TryGetValue("vAlign", out var vaVal) ||
            properties.TryGetValue("valign", out vaVal))
        {
            var lower = vaVal.ToLowerInvariant().Trim();
            if (lower is not ("none" or "off" or "false"))
            {
                var enumVal = lower switch
                {
                    "top" => VerticalJustificationValues.Top,
                    "center" or "centre" or "middle" => VerticalJustificationValues.Center,
                    "bottom" => VerticalJustificationValues.Bottom,
                    "both" => VerticalJustificationValues.Both,
                    _ => throw new ArgumentException(
                        $"Invalid vAlign value: '{vaVal}'. Valid: top, center, bottom, both, none.")
                };
                InsertSectPrChildInOrder(sectPr, new VerticalTextAlignmentOnPage { Val = enumVal });
            }
        }

        // BUG-DUMP-SECT-TEXTDIR: section page text flow (<w:textDirection>),
        // East-Asian vertical layout. CONSISTENCY(add-set-symmetry): mirror
        // TrySetSectionLayout's textDirection case so `add section
        // --prop textDirection=tbRl` round-trips a mid-document section's
        // <w:textDirection>. Distinct from the cell-level textDirection in tcPr.
        if (properties.TryGetValue("textDirection", out var tdVal) ||
            properties.TryGetValue("textdirection", out tdVal))
        {
            var lower = tdVal.ToLowerInvariant().Trim();
            if (lower is not ("none" or "off" or "false"))
                InsertSectPrChildInOrder(sectPr, new TextDirection { Val = ParseSectionTextDirection(tdVal) });
        }

        // BUG-DUMP-SECT-FORMPROT: <w:formProt/> section form-protection flag.
        // CONSISTENCY(add-set-symmetry): mirror TrySetSectionLayout's formprot
        // case so `add section --prop formProt=true` round-trips a mid-document
        // section's <w:formProt>. Bare on/off toggle.
        if (properties.TryGetValue("formProt", out var fpVal) ||
            properties.TryGetValue("formprot", out fpVal))
        {
            if (IsTruthy(fpVal))
                InsertSectPrChildInOrder(sectPr, new FormProtection());
        }

        // BUG-DUMP-SECT-FOOTNOTE: section-level footnote/endnote numbering.
        // CONSISTENCY(add-set-symmetry): route footnotePr.* / endnotePr.* through
        // the shared TrySetFootnoteEndnoteNumProps so `add section` round-trips
        // a mid-document section's <w:footnotePr>/<w:endnotePr>. Tracked below in
        // sectionAlreadyConsumed so the dotted fallback doesn't re-run.
        foreach (var (fk, fv) in properties)
        {
            var fkl = fk.ToLowerInvariant();
            if (fkl.StartsWith("footnotepr.") || fkl.StartsWith("endnotepr."))
            {
                properties.ContainsKey(fk); // ACCOUNTING(handler-as-truth)
                TrySetFootnoteEndnoteNumProps(sectPr, fk, fv);
            }
        }

        // Page border on a mid-document section carrier. CONSISTENCY(add-set-symmetry):
        // mirror TrySetSectionLayout's pgborders.<side> / .offsetFrom / .zOrder /
        // .display cases so `add section --prop pgBorders.top=single;4;auto;24 ...`
        // round-trips a carrier sectPr's <w:pgBorders> (e.g. a cover page boxing
        // only its first page via display=firstPage). The dump folds the per-side
        // sub-keys into the STYLE;SIZE;COLOR;SPACE value ParseBorderValue reads.
        // Keys consumed here are tracked in sectionAlreadyConsumed below so the
        // dotted fallback doesn't re-run TypedAttributeFallback against sectPr
        // (the per-side attrs live on the nested <w:pgBorders>, not sectPr).
        foreach (var (pk, pv) in properties)
        {
            var pkl = pk.ToLowerInvariant();
            if (!pkl.StartsWith("pgborders.")) continue;
            switch (pkl)
            {
                case "pgborders.top" or "pgborders.left"
                    or "pgborders.bottom" or "pgborders.right":
                {
                    var pb = EnsurePageBorders(sectPr);
                    var (style, size, color, space) = ParseBorderValue(pv);
                    var sf = ParseBorderShadowFrame(pv);
                    switch (pkl)
                    {
                        case "pgborders.top":    pb.TopBorder    = MakeBorder<TopBorder>(style, size, color, space, sf.shadow, sf.frame); break;
                        case "pgborders.left":   pb.LeftBorder   = MakeBorder<LeftBorder>(style, size, color, space, sf.shadow, sf.frame); break;
                        case "pgborders.bottom": pb.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space, sf.shadow, sf.frame); break;
                        case "pgborders.right":  pb.RightBorder  = MakeBorder<RightBorder>(style, size, color, space, sf.shadow, sf.frame); break;
                    }
                    break;
                }
                case "pgborders.offsetfrom":
                    EnsurePageBorders(sectPr).OffsetFrom = pv.ToLowerInvariant().Trim() switch
                    {
                        "page" => PageBorderOffsetValues.Page,
                        "text" => PageBorderOffsetValues.Text,
                        _ => throw new ArgumentException(
                            $"Invalid pgBorders.offsetFrom value: '{pv}'. Valid: page, text.")
                    };
                    break;
                case "pgborders.zorder":
                    EnsurePageBorders(sectPr).ZOrder = pv.ToLowerInvariant().Trim() switch
                    {
                        "front" => PageBorderZOrderValues.Front,
                        "back" => PageBorderZOrderValues.Back,
                        _ => throw new ArgumentException(
                            $"Invalid pgBorders.zOrder value: '{pv}'. Valid: front, back.")
                    };
                    break;
                case "pgborders.display":
                    EnsurePageBorders(sectPr).Display = pv.ToLowerInvariant().Trim() switch
                    {
                        "allpages" => PageBorderDisplayValues.AllPages,
                        "firstpage" => PageBorderDisplayValues.FirstPage,
                        "notfirstpage" => PageBorderDisplayValues.NotFirstPage,
                        _ => throw new ArgumentException(
                            $"Invalid pgBorders.display value: '{pv}'. Valid: allPages, firstPage, notFirstPage.")
                    };
                    break;
            }
        }

        // Dotted-key fallback for sectPr-level attrs not modeled by the
        // hand-rolled blocks above (single-attr forms like docGrid.* or
        // future schema additions). CONSISTENCY(add-set-symmetry).
        // Skip the dotted curated keys that AddSection already consumes
        // explicitly to avoid double application.
        var sectionAlreadyConsumed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "columns.count", "columns.space",
            // CONSISTENCY(add-set-symmetry): explicit hand-rolled branches above
            // consumed these — keep the dotted fallback from re-running
            // TypedAttributeFallback against sectPr (it'd fail to find the attr
            // on sectPr because separator/equalWidth/widths/spaces live on the
            // nested <w:cols>, then flag them as unsupported despite being
            // applied correctly).
            "columns.separator", "columns.equalwidth", "columns.equalWidth",
            "colwidths", "colWidths", "colspaces", "colSpaces",
            // pgBorders.* consumed by the dedicated block above; their per-side
            // attrs live on the nested <w:pgBorders>, not sectPr, so the dotted
            // fallback would falsely flag them unsupported.
            "pgBorders.top", "pgBorders.left", "pgBorders.bottom", "pgBorders.right",
            "pgBorders.offsetFrom", "pgBorders.zOrder", "pgBorders.display",
        };
        foreach (var (key, value) in properties)
        {
            if (!key.Contains('.')) continue;
            // ACCOUNTING(handler-as-truth): see AddStyle for rationale.
            // ContainsKey fires the TrackingComparer; without it, dotted
            // props consumed by TypedAttributeFallback leak as false
            // unsupported via TrackingPropertyDictionary.UnusedKeys.
            properties.ContainsKey(key);
            if (sectionAlreadyConsumed.Contains(key)) continue;
            // footnotePr.* / endnotePr.* are consumed by the dedicated loop
            // above; their attrs live on the nested <w:footnotePr>, not sectPr,
            // so TypedAttributeFallback would falsely flag them unsupported.
            if (key.StartsWith("footnotePr.", StringComparison.OrdinalIgnoreCase) ||
                key.StartsWith("endnotePr.", StringComparison.OrdinalIgnoreCase)) continue;
            if (Core.TypedAttributeFallback.TrySet(sectPr, key, value)) continue;
            LastAddUnsupportedProps.Add(key);
        }

        return AttachSectionBreakParagraph(body, parentPath, index, sectPara, sectPProps, sectPr, properties);
    }

    // Attach a section-break carrier paragraph (its sectPr already populated) at
    // the requested index and return its /section[N] path. Factored out so the
    // BUG-DUMP-R31-1 bare-section short-circuit can reuse the exact attach+path
    // logic without re-running the geometry/override blocks.
    private string AttachSectionBreakParagraph(
        Body body, string parentPath, int? index,
        Paragraph sectPara, ParagraphProperties sectPProps, SectionProperties sectPr,
        Dictionary<string, string> properties)
    {
        var parent = body; // section breaks always land in the body
        sectPProps.AppendChild(sectPr);
        sectPara.AppendChild(sectPProps);
        InsertAtIndexOrAppend(parent, sectPara, index);

        // Return the new section's document-order position (1-based) so the
        // path matches the NavigateToElement /section[N] resolver, which
        // walks body paragraphs with SectionProperties in document order.
        // Using the total count would break --before/--after (which insert
        // mid-document): the new section may not be the last one.
        var sectParas = body.Elements<Paragraph>()
            .Where(p => p.ParagraphProperties?.GetFirstChild<SectionProperties>() != null)
            .ToList();
        var secDocOrderIdx = sectParas.FindIndex(p => ReferenceEquals(p, sectPara));
        var resultPath = $"/section[{(secDocOrderIdx >= 0 ? secDocOrderIdx + 1 : sectParas.Count)}]";
        return resultPath;
    }

    // BUG-DUMP-NOTE-DEL: a footnote/endnote/comment whose seed content run carries
    // track-change attribution (revision.type=del|ins on the `add <kind>` op) must
    // have that run wrapped in <w:del>/<w:ins> — otherwise a tracked DELETION
    // resurfaces as LIVE accepted text (a silent meaning change), the same way the
    // body run path wraps via AddRun. Returns the wrapper (or the run unchanged when
    // no revision attribution). For w:del the run's <w:t> become <w:delText>.
    // Nested ins⊃del on a note seed is vanishingly rare and not handled here.
    private OpenXmlElement ApplyNoteSeedRevision(Run run, Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("revision.type", out var rt)
            || (rt != "del" && rt != "ins"))
            return run;
        OpenXmlElement wrapper = rt == "ins" ? new InsertedRun() : new DeletedRun();
        string? author = properties.GetValueOrDefault("revision.author");
        if (!string.IsNullOrEmpty(author))
        {
            if (wrapper is InsertedRun iw) iw.Author = author;
            else if (wrapper is DeletedRun dw) dw.Author = author;
        }
        if (properties.TryGetValue("revision.date", out var dstr) && !string.IsNullOrEmpty(dstr)
            && DateTime.TryParse(dstr, null, System.Globalization.DateTimeStyles.RoundtripKind, out var dt))
        {
            if (wrapper is InsertedRun iw2) iw2.Date = dt;
            else if (wrapper is DeletedRun dw2) dw2.Date = dt;
        }
        var id = properties.GetValueOrDefault("revision.id");
        if (string.IsNullOrEmpty(id)) id = GenerateRevisionId();
        if (wrapper is InsertedRun iw3) iw3.Id = id;
        else if (wrapper is DeletedRun dw3) dw3.Id = id;
        if (rt == "del")
            foreach (var t in run.Elements<Text>().ToList())
                t.Parent?.ReplaceChild(new DeletedText(t.Text ?? "") { Space = t.Space }, t);
        wrapper.AppendChild(run);
        return wrapper;
    }

    private string AddFootnote(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("text", out var fnText))
            throw new ArgumentException("'text' property is required for footnote type");
        OfficeCli.Core.ParseHelpers.ValidateXmlText(fnText, "text", allowSoftBreakChar: true);

        if (parent is not Paragraph fnPara)
            throw new ArgumentException("Footnotes must be added to a paragraph: /body/p[N]");

        // OOXML: CT_HdrFtr has no footnoteReference content model. A footnote
        // reference written into a header/footer part is silently ignored by
        // Word (the note body in the global footnotes part becomes unreachable).
        // Reject upfront — same silent-accept-bad-OOXML family as altChunk and
        // nested-SDT.
        if (parentPath.StartsWith("/header[", StringComparison.OrdinalIgnoreCase)
            || parentPath.StartsWith("/footer[", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException(
                $"Footnotes cannot be added inside a header or footer (path: {parentPath}). " +
                "Add footnotes only to body paragraphs: /body/p[N].");

        var mainPart2 = _doc.MainDocumentPart!;
        var fnPart = mainPart2.FootnotesPart ?? mainPart2.AddNewPart<FootnotesPart>();
        // Word emits the mandatory separator notes carrying the separator
        // glyph marks (<w:separator/> / <w:continuationSeparator/>), not empty
        // text runs. Synthesizing them as the spec-defined marks keeps a
        // dump->batch round-trip of a real document structurally faithful.
        fnPart.Footnotes ??= new Footnotes(
            new Footnote(new Paragraph(new Run(new SeparatorMark()))) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
            new Footnote(new Paragraph(new Run(new ContinuationSeparatorMark()))) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 }
        );

        var fnId = (fnPart.Footnotes.Elements<Footnote>()
            .Where(f => f.Id?.Value > 0)
            .Select(f => f.Id!.Value)
            .DefaultIfEmpty(0).Max() + 1);

        // BUG-DUMP-R42-1: the ref-mark run links to a char style
        // (FootnoteReference) via <w:rStyle> in real Word docs. When the dump
        // carried that source style as `referenceStyle`, restore the rStyle
        // link instead of inlining a hardcoded vertAlign superscript (which
        // severs the style link). With no referenceStyle, keep the legacy
        // inline superscript.
        var fnRefMarkRPr = properties.TryGetValue("referenceMarkRPr", out var fnMarkRPrXml)
                && !string.IsNullOrEmpty(fnMarkRPrXml)
            // Verbatim in-note mark rPr forwarded by the dump (rStyle PLUS any
            // direct rFonts/sz the source put on the mark run).
            ? new RunProperties(fnMarkRPrXml)
            : properties.TryGetValue("referenceStyle", out var fnRefStyle)
                && !string.IsNullOrEmpty(fnRefStyle)
            ? new RunProperties(new RunStyle { Val = fnRefStyle })
            : new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
        var footnote = new Footnote { Id = fnId };
        // BUG-DUMP-NOTEREF-CUSTOMMARK-BODY: a CUSTOM-mark footnote's body does NOT
        // begin with the auto-number <w:footnoteRef> — its leading run is the
        // literal custom glyph (*, †, …) in a run styled FootnoteReference, and
        // the body reference in the document carries w:customMarkFollows="1" so
        // Word suppresses the auto-number. Injecting <w:footnoteRef> here prepended
        // the auto-number to the custom glyph on rebuild ("6*" instead of "*").
        // When the dump signals a custom mark (referenceCustomMark[Follows]), the
        // `text` prop already IS that glyph — emit it as a styled mark run and
        // skip the auto-number FootnoteReferenceMark. Otherwise keep the standard
        // auto-number mark + text run.
        bool fnCustomMark =
            IsTruthy(properties.GetValueOrDefault("referenceCustomMarkFollows"))
            || (properties.TryGetValue("referenceCustomMark", out var fnBodyMark)
                && !string.IsNullOrEmpty(fnBodyMark));
        // The authored text run carries `text` VERBATIM — no synthetic leading
        // space. Real Word footnotes place no space between the reference mark
        // and the first content run (the mark's own glyph spacing handles the
        // gap), so a dump→batch round-trip of a source note must reproduce its
        // first <w:t> byte-for-byte. A previously prepended " " inflated every
        // rebuilt footnote's first run by one U+0020; GetFootnoteText trimmed it
        // back on readback (hiding it from get/view) while the on-disk XML drifted.
        Paragraph fnContentPara;
        if (fnCustomMark)
            // Custom mark: the glyph (text) IS the mark run, styled like the
            // source (fnRefMarkRPr carries the FootnoteReference rStyle when the
            // dump forwarded it). No auto-number <w:footnoteRef>.
            fnContentPara = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "FootnoteText" }),
                new Run(fnRefMarkRPr, new Text(fnText) { Space = SpaceProcessingModeValues.Preserve })
            );
        else
        {
            // BUG-DUMP-NOTE-TAB: build the content run via AppendTextWithBreaks so a
            // tab/newline in the seed text becomes a structural <w:tab/> / <w:br/>
            // (mirrors `add r`), not a literal U+0009/U+000A glyph. A footnote ref
            // mark is commonly followed by a <w:tab/> separating the number from the
            // text; the verbatim Text node dropped that structural tab.
            var fnTextRun = new Run();
            AppendTextWithBreaks(fnTextRun, fnText);
            fnContentPara = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "FootnoteText" }),
                new Run(fnRefMarkRPr, new FootnoteReferenceMark()),
                ApplyNoteSeedRevision(fnTextRun, properties)   // BUG-DUMP-NOTE-DEL
            );
        }
        footnote.AppendChild(fnContentPara);
        // i18n: route remaining keys (direction, font.cs, bold.cs, etc.)
        // through the same paragraph + run helpers SetFootnotePath uses.
        // Mirrors AddHeader's R2-2 fix so RTL footnotes work end-to-end.
        var fnUnsupported = new List<string>();
        ApplyFootnoteEndnoteFormatKeys(footnote, properties, fnUnsupported);
        foreach (var u in fnUnsupported) LastAddUnsupportedProps.Add(u);
        fnPart.Footnotes.AppendChild(footnote);
        fnPart.Footnotes.Save();

        // Insert reference in document body at the requested index, keeping
        // pPr as first child (InsertIntoParagraph clamps forward past pPr).
        // CONSISTENCY(rtl-cascade): if the host paragraph is RTL, stamp
        // <w:rtl/> on the reference run's rPr so the superscript number
        // renders on the correct side of an Arabic / Hebrew paragraph.
        // The body reference run can carry direct formatting beyond the char
        // style (run-level rFonts/sz shrinking the superscript mark). When the
        // dump forwards the verbatim <w:rPr> (referenceRPr), restore it; else
        // keep the bare style link.
        var fnRefRPr = properties.TryGetValue("referenceRPr", out var fnRefRPrXml)
                && !string.IsNullOrEmpty(fnRefRPrXml)
            ? new RunProperties(fnRefRPrXml)
            : new RunProperties(new RunStyle { Val = "FootnoteReference" });
        if (fnPara.ParagraphProperties?.BiDi != null)
            ApplyRunFormatting(fnRefRPr, "rtl", "true");
        // BUG-DUMP-NOTEREF-CUSTOMMARK: a custom-mark reference carries
        // w:customMarkFollows="1" and the literal mark glyph in a sibling <w:t>
        // inside the SAME body run. Restore both so the asterisk/dagger survives.
        var fnRef = new FootnoteReference { Id = fnId };
        if (IsTruthy(properties.GetValueOrDefault("referenceCustomMarkFollows")))
            fnRef.CustomMarkFollows = true;
        var fnRefRun = new Run(fnRefRPr, fnRef);
        properties.TryGetValue("referenceCustomMark", out var fnMark);
        var fnToInsert = ApplyNoteRefMarkAndRevision(fnRefRun, fnMark, properties);
        InsertIntoParagraph(fnPara, fnToInsert, index);

        var resultPath = $"/footnote[@footnoteId={fnId}]";
        return resultPath;
    }

    private string AddEndnote(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("text", out var enText))
            throw new ArgumentException("'text' property is required for endnote type");
        OfficeCli.Core.ParseHelpers.ValidateXmlText(enText, "text", allowSoftBreakChar: true);

        if (parent is not Paragraph enPara)
            throw new ArgumentException("Endnotes must be added to a paragraph: /body/p[N]");

        // OOXML: CT_HdrFtr has no endnoteReference content model. Same
        // silent-accept-bad-OOXML family as footnotes; reject upfront.
        if (parentPath.StartsWith("/header[", StringComparison.OrdinalIgnoreCase)
            || parentPath.StartsWith("/footer[", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException(
                $"Endnotes cannot be added inside a header or footer (path: {parentPath}). " +
                "Add endnotes only to body paragraphs: /body/p[N].");

        var mainPart3 = _doc.MainDocumentPart!;
        var enPart = mainPart3.EndnotesPart ?? mainPart3.AddNewPart<EndnotesPart>();
        // Mirror AddFootnote: separator notes carry the separator glyph marks
        // (<w:separator/> / <w:continuationSeparator/>), not empty text runs.
        enPart.Endnotes ??= new Endnotes(
            new Endnote(new Paragraph(new Run(new SeparatorMark()))) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
            new Endnote(new Paragraph(new Run(new ContinuationSeparatorMark()))) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 }
        );

        var enId = (enPart.Endnotes.Elements<Endnote>()
            .Where(e => e.Id?.Value > 0)
            .Select(e => e.Id!.Value)
            .DefaultIfEmpty(0).Max() + 1);

        // BUG-DUMP-R42-1: mirror AddFootnote — restore the EndnoteReference
        // char-style link from the carried `referenceStyle` prop, else fall
        // back to the inline vertAlign superscript.
        var enRefMarkRPr = properties.TryGetValue("referenceMarkRPr", out var enMarkRPrXml)
                && !string.IsNullOrEmpty(enMarkRPrXml)
            ? new RunProperties(enMarkRPrXml)
            : properties.TryGetValue("referenceStyle", out var enRefStyle)
                && !string.IsNullOrEmpty(enRefStyle)
            ? new RunProperties(new RunStyle { Val = enRefStyle })
            : new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
        var endnote = new Endnote { Id = enId };
        // BUG-DUMP-NOTEREF-CUSTOMMARK-BODY: mirror AddFootnote — a custom-mark
        // endnote's body leads with the literal glyph (styled EndnoteReference),
        // not the auto-number <w:endnoteRef>. Injecting the auto-number prepended
        // it to the glyph ("6*"). Suppress it when the dump signals a custom mark.
        bool enCustomMark =
            IsTruthy(properties.GetValueOrDefault("referenceCustomMarkFollows"))
            || (properties.TryGetValue("referenceCustomMark", out var enBodyMark)
                && !string.IsNullOrEmpty(enBodyMark));
        // Verbatim text run — no synthetic leading space (mirrors AddFootnote).
        Paragraph enContentPara;
        if (enCustomMark)
            enContentPara = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "EndnoteText" }),
                new Run(enRefMarkRPr, new Text(enText) { Space = SpaceProcessingModeValues.Preserve })
            );
        else
        {
            // BUG-DUMP-NOTE-TAB: see AddFootnote — split tab/newline in the seed
            // text into structural <w:tab/> / <w:br/> instead of a literal glyph.
            var enTextRun = new Run();
            AppendTextWithBreaks(enTextRun, enText);
            enContentPara = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "EndnoteText" }),
                new Run(enRefMarkRPr, new EndnoteReferenceMark()),
                ApplyNoteSeedRevision(enTextRun, properties)   // BUG-DUMP-NOTE-DEL
            );
        }
        endnote.AppendChild(enContentPara);
        // i18n: route remaining keys through the same helper as footnote.
        var enUnsupported = new List<string>();
        ApplyFootnoteEndnoteFormatKeys(endnote, properties, enUnsupported);
        foreach (var u in enUnsupported) LastAddUnsupportedProps.Add(u);
        enPart.Endnotes.AppendChild(endnote);
        enPart.Endnotes.Save();

        // Insert reference in document body at the requested index, keeping
        // pPr as first child (InsertIntoParagraph clamps forward past pPr).
        // CONSISTENCY(rtl-cascade): mirror the footnote case — RTL host
        // paragraphs stamp <w:rtl/> on the reference run's rPr.
        // Mirror AddFootnote: restore the body reference run's verbatim rPr
        // when the dump forwards it (referenceRPr).
        var enRefRPr = properties.TryGetValue("referenceRPr", out var enRefRPrXml)
                && !string.IsNullOrEmpty(enRefRPrXml)
            ? new RunProperties(enRefRPrXml)
            : new RunProperties(new RunStyle { Val = "EndnoteReference" });
        if (enPara.ParagraphProperties?.BiDi != null)
            ApplyRunFormatting(enRefRPr, "rtl", "true");
        // BUG-DUMP-NOTEREF-CUSTOMMARK: mirror AddFootnote — restore a custom
        // mark (w:customMarkFollows + sibling <w:t> glyph) in the body ref run.
        var enRef = new EndnoteReference { Id = enId };
        if (IsTruthy(properties.GetValueOrDefault("referenceCustomMarkFollows")))
            enRef.CustomMarkFollows = true;
        var enRefRun = new Run(enRefRPr, enRef);
        properties.TryGetValue("referenceCustomMark", out var enMark);
        var enToInsert = ApplyNoteRefMarkAndRevision(enRefRun, enMark, properties);
        InsertIntoParagraph(enPara, enToInsert, index);

        var resultPath = $"/endnote[@endnoteId={enId}]";
        return resultPath;
    }

    // BUG-DUMP-NOTEREF-CUSTOMMARK-DEL: append a note reference's custom mark glyph
    // and, when the source reference run was a tracked revision (reference.revision.*
    // — most commonly a <w:del> wrapping a deleted reference), re-wrap the rebuilt
    // run so the deletion/insertion/move survives instead of resurfacing as live,
    // accepted content. In a deletion the mark glyph must ride in <w:delText>.
    private OpenXmlElement ApplyNoteRefMarkAndRevision(
        Run refRun, string? customMark, Dictionary<string, string> properties)
    {
        var revType = properties.GetValueOrDefault("reference.revision.type");
        bool isDel = string.Equals(revType, "del", StringComparison.OrdinalIgnoreCase);
        // BUG-DUMP-H97: a SYMBOL-glyph custom mark (<w:sym w:font=… w:char=…/>)
        // round-trips via referenceCustomMarkSym ("font:char"); rebuild the <w:sym>
        // element rather than a <w:t>. Takes precedence over the text mark (a sym
        // mark carries no <w:t>, so customMark is empty for it).
        if (properties.TryGetValue("referenceCustomMarkSym", out var symMark)
            && !string.IsNullOrEmpty(symMark))
        {
            var sp = symMark.Split(':', 2);
            refRun.AppendChild(new SymbolChar
            {
                Font = sp[0],
                Char = sp.Length > 1 ? sp[1] : ""
            });
        }
        else if (!string.IsNullOrEmpty(customMark))
            refRun.AppendChild(isDel
                ? new DeletedText(customMark) { Space = SpaceProcessingModeValues.Preserve }
                : (OpenXmlElement)new Text(customMark) { Space = SpaceProcessingModeValues.Preserve });

        if (string.IsNullOrEmpty(revType))
            return refRun;

        RunTrackChangeType? wrapper = revType.ToLowerInvariant() switch
        {
            "del" => new DeletedRun(),
            "ins" => new InsertedRun(),
            "movefrom" => new MoveFromRun(),
            "moveto" => new MoveToRun(),
            _ => null
        };
        if (wrapper == null)
            return refRun;

        if (properties.TryGetValue("reference.revision.author", out var au) && !string.IsNullOrEmpty(au))
            wrapper.Author = au;
        if (properties.TryGetValue("reference.revision.date", out var dt) && !string.IsNullOrEmpty(dt)
            && DateTime.TryParse(dt, null, System.Globalization.DateTimeStyles.RoundtripKind, out var d))
            wrapper.Date = d;
        wrapper.Id = properties.TryGetValue("reference.revision.id", out var rid) && !string.IsNullOrEmpty(rid)
            ? rid
            : GenerateRevisionId();
        wrapper.AppendChild(refRun);
        return wrapper;
    }

    private string AddToc(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        // TOC fields reference body-level heading styles; adding them in a
        // header/footer part is not meaningful and would yield an unnavigable
        // /toc[0] return path (body TOC count is 0). Reject early with a
        // clean error.
        if (parent is Header || parent is Footer
            || parent.Ancestors<Header>().Any() || parent.Ancestors<Footer>().Any())
        {
            throw new ArgumentException(
                "add --type toc is not supported inside a header or footer part. " +
                "TOC field codes reference body-level headings; insert into /body instead.");
        }

        // Table of Contents field code
        var levels = properties.GetValueOrDefault("levels", "1-3");
        var tocTitle = properties.GetValueOrDefault("title", "");
        var hyperlinks = !properties.TryGetValue("hyperlinks", out var hlVal) || IsTruthy(hlVal);
        var pageNumbers = !properties.TryGetValue("pagenumbers", out var pnVal) || IsTruthy(pnVal);

        // Build field code instruction
        var instrBuilder = new StringBuilder($" TOC \\o \"{levels}\"");
        if (hyperlinks) instrBuilder.Append(" \\h");
        if (!pageNumbers) instrBuilder.Append(" \\z");
        // BUG-R5-03: \t = custom-style→level mapping (Word's "Style; level"
        // syntax, e.g. "MyHeading,1,MySub,2"); \b = bookmark scope (single
        // bookmark name). Both round-trip through dump→add and were
        // silently dropped before, breaking custom TOC layouts.
        if (properties.TryGetValue("customStyles", out var cs) && !string.IsNullOrEmpty(cs))
            instrBuilder.Append($" \\t \"{cs}\"");
        if (properties.TryGetValue("bookmark", out var bm) && !string.IsNullOrEmpty(bm))
            instrBuilder.Append($" \\b \"{bm}\"");
        instrBuilder.Append(" \\u ");

        var tocPara = new Paragraph();

        // Field begin
        tocPara.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
        // Field code
        tocPara.AppendChild(new Run(new FieldCode(instrBuilder.ToString()) { Space = SpaceProcessingModeValues.Preserve }));
        // Field separate
        tocPara.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
        // Placeholder text
        tocPara.AppendChild(new Run(new Text("Update field to see table of contents") { Space = SpaceProcessingModeValues.Preserve }));
        // Field end
        tocPara.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

        // Insert TOC paragraph at the requested position first, then — if a
        // title was requested — insert the title paragraph immediately before
        // it so the title precedes the TOC field in reading order. Previously
        // the title was appended to the parent regardless of --index, ending
        // up after the TOC.
        InsertAtIndexOrAppend(parent, tocPara, index);
        if (!string.IsNullOrEmpty(tocTitle))
        {
            var titlePara = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "TOCHeading" }),
                new Run(new Text(tocTitle))
            );
            tocPara.InsertBeforeSelf(titlePara);
        }

        // Intentionally do NOT set <w:updateFieldsOnOpen w:val="true"/>: it
        // makes Word prompt the user with "update fields?" on every open.
        // The TOC field result stays empty until the user right-clicks ->
        // "Update Field" (or presses F9). Trade-off accepted: empty-by-default
        // beats a dialog every open, since we can't pre-render real page
        // numbers without a layout engine. See chat 2026-05-05.

        // Determine TOC index in document order (not total count)
        var tocParas = body.Elements<Paragraph>()
            .Where(p => p.Descendants<FieldCode>().Any(fc =>
                fc.Text != null && fc.Text.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase)))
            .ToList();
        var tocIdx = tocParas.FindIndex(p => ReferenceEquals(p, tocPara));
        var resultPath = $"/toc[{(tocIdx >= 0 ? tocIdx + 1 : tocParas.Count)}]";
        return resultPath;
    }

    // ==================== CT_Style schema-order insertion ====================
    // BUG-DUMP-STYLE-LATENT: canonical CT_Style child order (subset, in
    // document order):
    //   name, aliases, basedOn, next, link, autoRedefine, hidden,
    //   uiPriority, semiHidden, unhideWhenUsed, qFormat, locked,
    //   personalCompose, personalReply, personal, rPr, pPr, tblPr, trPr, tcPr,
    //   tblStylePr.
    // Word and the OOXML validator reject out-of-order children, so a Set that
    // inserts a latent-style flag onto a style that already has pPr/rPr must
    // place it ahead of those. AddStyle appends in this order already; Set uses
    // this helper to splice into the right slot. Maps a child to its rank.
    private static int StyleChildOrder(OpenXmlElement el) => el switch
    {
        StyleName => 0,
        Aliases => 1,
        BasedOn => 2,
        NextParagraphStyle => 3,
        LinkedStyle => 4,
        AutoRedefine => 5,
        StyleHidden => 6,
        UIPriority => 7,
        SemiHidden => 8,
        UnhideWhenUsed => 9,
        PrimaryStyle => 10,   // <w:qFormat/>
        Locked => 11,
        StyleParagraphProperties => 14,
        StyleRunProperties => 13,
        _ => 99,
    };

    // Insert (or replace) a single-occurrence latent-style flag child at the
    // CT_Style schema-order position. Removes any existing instance first so
    // re-applying via Set is idempotent and never duplicates the element.
    private static void InsertStyleChildInOrder(Style style, OpenXmlElement newChild)
    {
        var newRank = StyleChildOrder(newChild);
        // Drop any existing instance of the same element type.
        var existing = style.ChildElements
            .Where(c => c.GetType() == newChild.GetType()).ToList();
        foreach (var e in existing) e.Remove();

        OpenXmlElement? successor = null;
        foreach (var child in style.ChildElements)
        {
            if (StyleChildOrder(child) > newRank) { successor = child; break; }
        }
        if (successor != null) successor.InsertBeforeSelf(newChild);
        else style.AppendChild(newChild);
    }

    private string AddStyle(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        // Create a new style in the styles part
        var stylesPart = _doc.MainDocumentPart!.StyleDefinitionsPart
            ?? _doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles ??= new Styles();

        // CONSISTENCY(style-dual-key): Get exposes the canonical readback
        // keys `styleId` and `styleName` on every paragraph (Round 2). Add
        // must accept the same alias trio (id / styleId / name / styleName)
        // or the readback writes back as `CustomStyle` — exactly the
        // silent-ignore alias trap that 19b3dd5b banned.
        var explicitId = properties.ContainsKey("id") || properties.ContainsKey("styleId") || properties.ContainsKey("styleid");
        var styleId = properties.GetValueOrDefault("id")
                   ?? properties.GetValueOrDefault("styleId")
                   ?? properties.GetValueOrDefault("styleid")
                   ?? properties.GetValueOrDefault("name")
                   ?? properties.GetValueOrDefault("styleName")
                   ?? properties.GetValueOrDefault("stylename")
                   ?? "CustomStyle";
        // BUG-R7-08: when the caller passes only `id` (no name), AddStyle used
        // to default the name to the id. That mutated the round-trip output
        // for any docx whose original style had an `id` but no `<w:name>`
        // (or empty name) — the next dump showed `name=<id>`. Preserve the
        // "no explicit name" intent by emitting an empty <w:name w:val=""/>
        // (still schema-valid; matches the original).
        var explicitName = properties.ContainsKey("name")
            || properties.ContainsKey("styleName")
            || properties.ContainsKey("stylename");
        var styleName = properties.GetValueOrDefault("name")
                     ?? properties.GetValueOrDefault("styleName")
                     ?? properties.GetValueOrDefault("stylename")
                     ?? (explicitName ? styleId : "");
        var styleType = properties.GetValueOrDefault("type", "paragraph").ToLowerInvariant() switch
        {
            "character" or "char" => StyleValues.Character,
            "table" => StyleValues.Table,
            "numbering" => StyleValues.Numbering,
            "paragraph" or "para" => StyleValues.Paragraph,
            _ => throw new ArgumentException($"Invalid style type: '{properties.GetValueOrDefault("type", "paragraph")}'. Valid values: paragraph, character, table, numbering.")
        };

        // Enforce unique styleId — schema requires unique w:styleId per styles.xml.
        // If the caller specified --prop id explicitly, reject; otherwise auto-suffix
        // to keep the call idempotent-ish for scripts that only pass --prop name.
        bool IdTaken(string candidate) => stylesPart.Styles.Elements<Style>()
            .Any(s => string.Equals(s.StyleId?.Value, candidate, StringComparison.Ordinal));
        // BUG-R6-03: dump→batch on a fresh blank docx fails 42×
        // ("Style Normal already exists") because real documents always
        // carry built-in style definitions (Normal, Heading1-9, Title,
        // ListParagraph, …) and the blank template ships with the same
        // ids reserved. For built-in ids the safe semantics is upsert:
        // remove the existing definition and let the rest of AddStyle
        // re-create it with the caller's full property bag. Mirrors
        // BlankDocCreator's hands-off treatment of built-ins (it only
        // registers the bare style scaffolding).
        var builtInIdsForUpsert = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Normal", "Heading1", "Heading2", "Heading3", "Heading4", "Heading5",
            "Heading6", "Heading7", "Heading8", "Heading9", "Title", "Subtitle",
            "Quote", "IntenseQuote", "ListParagraph", "NoSpacing", "TOCHeading",
            "DefaultParagraphFont", "TableNormal", "NoList",
        };
        // Built-in style display names (parallel to the id set above). The
        // dump→batch input often carries `id="1" name="Normal"` — numeric
        // styleIds from the source doc paired with canonical built-in display
        // names from that doc's styles.xml. The id-based upsert above misses
        // this case because the blank template's "Normal" lives at
        // styleId="Normal", not "1", so IdTaken("1") is false and the strict
        // name check below throws. Track built-in names too so name-side
        // collisions with a built-in entry trigger the same upsert. OOXML
        // built-in display names follow Word's styles.xml convention
        // (heading names are lowercase + space).
        var builtInNamesForUpsert = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Normal",
            "heading 1", "heading 2", "heading 3", "heading 4", "heading 5",
            "heading 6", "heading 7", "heading 8", "heading 9",
            "Title", "Subtitle", "Quote", "Intense Quote",
            "List Paragraph", "No Spacing", "TOC Heading",
            "Default Paragraph Font", "Table Normal", "No List",
        };
        if (IdTaken(styleId))
        {
            if (builtInIdsForUpsert.Contains(styleId))
            {
                // Idempotent re-add: drop the existing definition. We
                // preserve the explicitId path's strictness for non-
                // built-in ids so users authoring custom styles still
                // see a clear "duplicate id" error.
                var existing = stylesPart.Styles.Elements<Style>()
                    .FirstOrDefault(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.Ordinal));
                existing?.Remove();
            }
            else if (explicitId)
                throw new ArgumentException(
                    $"Style '{styleId}' already exists. Pick a unique --prop id or --prop name.");
            else
            {
                var baseId = styleId;
                int suffix = 2;
                while (IdTaken(styleId)) styleId = $"{baseId}{suffix++}";
            }
        }

        // OOXML requires w:name to be unique across styles.xml, same as w:styleId.
        // Reject duplicate display names — silently auto-suffixing the id while
        // leaving name unchanged produced two styles with identical UI labels
        // that users could not tell apart (BUG-R17-02).
        bool NameTaken(string candidate) => stylesPart.Styles.Elements<Style>()
            .Any(s => string.Equals(s.StyleName?.Val?.Value, candidate, StringComparison.Ordinal));
        // BUG-R7-08: empty styleName (id-only style, see styleName fallback
        // above) is allowed to repeat — multiple unnamed styles round-trip
        // from real docx files where the author left out the display name.
        if (!string.IsNullOrEmpty(styleName) && NameTaken(styleName))
        {
            // Name-side upsert: mirror the styleId branch above but key on
            // display name. Covers dump→batch where the caller carries the
            // source doc's numeric/custom styleId alongside a canonical
            // built-in name (e.g. id="1" name="Normal" from gov.cn corpus).
            // Drop the colliding built-in definition and fall through so
            // AddStyle creates the user's style with their full property
            // bag. Non-built-in name collisions (custom display names)
            // still throw — strictness preserved for user-authored
            // duplicates.
            var existingByName = stylesPart.Styles.Elements<Style>()
                .FirstOrDefault(s => string.Equals(s.StyleName?.Val?.Value, styleName, StringComparison.Ordinal));
            var existingIsBuiltIn = existingByName != null
                && ((existingByName.StyleId?.Value is { } eId && builtInIdsForUpsert.Contains(eId))
                    || builtInNamesForUpsert.Contains(styleName));
            if (existingIsBuiltIn)
            {
                existingByName!.Remove();
            }
            else if (explicitId)
            {
                // Two distinct styles (the styleId collision is already caught
                // above, so this explicit styleId is unique) sharing a display
                // name is a real-world artifact — third-party editors / document merges
                // produce e.g. styleId "Subtitle1" + "Subtitle10" both named
                // "Subtitle1". Word keys its Styles pane by name, so the
                // duplicate collides in the UI, but NOTHING in the document
                // references the name (pStyle / basedOn / link / next all use
                // styleId). Auto-suffix the display name to a unique value: both
                // styles survive and become distinguishable in Word, and
                // dump→batch of such a source no longer fails (or drops the
                // second style). Earlier BUG-R17-02 suffixed the *id* and left
                // identical names — suffixing the name is the correct resolution.
                // Gated on explicitId: an interactive `add style name=X` with no
                // id is a likely typo and still gets the hard error below.
                var baseName = styleName!;
                int suffix = 2;
                do { styleName = $"{baseName} {suffix++}"; } while (NameTaken(styleName));
            }
            else
            {
                throw new ArgumentException(
                    $"Style with name '{styleName}' already exists. Pick a unique --prop name.");
            }
        }

        // Built-in styles must not have customStyle=true, or Word won't recognize them
        // (e.g. TOC won't find Heading1 if it's marked as custom).
        // BUG-023 — single source of truth: reuse the upsert set above so that
        // DefaultParagraphFont / TableNormal / NoList (idempotent re-adds on
        // dump→batch) don't get stamped customStyle=true and break Word's
        // run-style fallback chain.
        var isBuiltIn = builtInIdsForUpsert.Contains(styleId);

        var newStyle = new Style
        {
            Type = styleType,
            StyleId = styleId,
        };
        // BUG-DUMP-R30-1: when the caller carries an explicit customStyle prop
        // (dump→batch round-trip — see ReadStyleLatentFlags), honor it verbatim
        // instead of re-deriving built-in-ness from styleId. The styleId
        // discriminator is wrong for docs that rename built-in styles to short
        // ids (Normal→"a", Heading1→"1"): those ids aren't in
        // builtInIdsForUpsert, so the old code stamped customStyle=true on a
        // renamed Normal. Word then stops resolving that style against its
        // built-in Normal table, changing CJK/Latin justification fitting and
        // reflowing every justified line. With the prop present we emit
        // customStyle only when the source had it (true) and leave it absent
        // otherwise — exactly the source shape. Interactive `add style` with no
        // customStyle prop still falls back to the styleId derivation below.
        if (properties.TryGetValue("customStyle", out var sCustomStyle)
            || properties.TryGetValue("customstyle", out sCustomStyle))
        {
            if (IsTruthy(sCustomStyle))
                newStyle.CustomStyle = true;
        }
        else if (!isBuiltIn)
            newStyle.CustomStyle = true;
        // w:default="1" — designate this as the document's default style for its
        // type (paragraph/character/table/numbering). Lets the dump round-trip a
        // source whose default style uses a renamed id (e.g. styleId="Standard"
        // with name "Normal"); without it, replayed body paragraphs inherit a
        // different default and render at the wrong size/alignment.
        if (properties.TryGetValue("default", out var sDefault) && IsTruthy(sDefault))
            newStyle.Default = true;
        // BUG-DUMP-STYLE-EMPTY-NAME: append <w:name> only when the caller actually
        // gave a name. A style added with an id but no name key must round-trip
        // WITHOUT a <w:name> child — auto table sub-styles (Table1/2/3, …) carry no
        // <w:name> in the source, and the recursive style-decomposition emit drops
        // an empty name from the shell props, so explicitName is false here for
        // them. The old unconditional append injected a spurious <w:name w:val=""/>
        // the source never had (validate-clean and render-identical, but it broke
        // the decomposition's exact-element-multiset guarantee). BUG-R7-08
        // (never default the name to the id) is still honored: no name in, no name
        // out. An explicit empty name (name="" passed) still emits <w:name w:val=""/>.
        if (explicitName)
            newStyle.AppendChild(new StyleName { Val = styleName });

        if ((properties.TryGetValue("basedon", out var basedOn) || properties.TryGetValue("basedOn", out basedOn)) && !string.IsNullOrEmpty(basedOn))
            newStyle.AppendChild(new BasedOn { Val = basedOn });
        if (properties.TryGetValue("next", out var nextStyle))
            newStyle.AppendChild(new NextParagraphStyle { Val = nextStyle });
        // <w:link w:val="…"/> — pair a paragraph style with its companion
        // character style (Word's "Linked Styles" feature). Schema order:
        // after `next`, before autoRedefine/hidden/uiPriority. The referenced
        // id must exist in styles.xml; we do not validate here — the dump→batch
        // emit order guarantees both styles land before any consumer.
        if ((properties.TryGetValue("linked", out var sLinked)
                || properties.TryGetValue("link", out sLinked))
            && !string.IsNullOrEmpty(sLinked))
        {
            newStyle.AppendChild(new LinkedStyle { Val = sLinked });
        }
        // BUG-DUMP11-05: top-level Style flags — autoRedefine + hidden.
        // Schema order: after `next`, before pPr/rPr. Toggle elements; only
        // emit when truthy. ParseHelpers.IsTruthy throws on unrecognized
        // values to match the rest of the Word handler's strict-bool intake.
        if (properties.TryGetValue("autoRedefine", out var sAutoRedef)
            || properties.TryGetValue("autoredefine", out sAutoRedef))
        {
            if (IsTruthy(sAutoRedef)) newStyle.AppendChild(new AutoRedefine());
        }
        if (properties.TryGetValue("hidden", out var sHidden))
        {
            if (IsTruthy(sHidden)) newStyle.AppendChild(new StyleHidden());
        }
        // BUG-DUMP-STYLE-LATENT: latent-style flags. CT_Style child order is
        // strict — these come AFTER autoRedefine/hidden and BEFORE pPr/rPr in
        // exactly this sequence: uiPriority, semiHidden, unhideWhenUsed,
        // qFormat, locked. Out-of-order children are rejected by Word/validator.
        // Without these, the default Normal style's <w:qFormat/> (and authored
        // uiPriority/semiHidden) vanished on every dump→batch round-trip.
        if ((properties.TryGetValue("uiPriority", out var sUiPriority)
                || properties.TryGetValue("uipriority", out sUiPriority))
            && !string.IsNullOrEmpty(sUiPriority))
        {
            newStyle.AppendChild(new UIPriority { Val = ParseHelpers.SafeParseInt(sUiPriority, "uiPriority") });
        }
        if (properties.TryGetValue("semiHidden", out var sSemiHidden)
            || properties.TryGetValue("semihidden", out sSemiHidden))
        {
            if (IsTruthy(sSemiHidden)) newStyle.AppendChild(new SemiHidden());
        }
        if (properties.TryGetValue("unhideWhenUsed", out var sUnhide)
            || properties.TryGetValue("unhidewhenused", out sUnhide))
        {
            if (IsTruthy(sUnhide)) newStyle.AppendChild(new UnhideWhenUsed());
        }
        if (properties.TryGetValue("qFormat", out var sQFormat)
            || properties.TryGetValue("qformat", out sQFormat))
        {
            if (IsTruthy(sQFormat)) newStyle.AppendChild(new PrimaryStyle());
        }
        if (properties.TryGetValue("locked", out var sLocked))
        {
            if (IsTruthy(sLocked)) newStyle.AppendChild(new Locked());
        }

        // Style paragraph properties
        var stylePPr = new StyleParagraphProperties();
        bool hasPPr = false;
        if (properties.TryGetValue("align", out var sAlign) || properties.TryGetValue("alignment", out sAlign) || properties.TryGetValue("jc", out sAlign))
        {
            stylePPr.Justification = new Justification { Val = ParseJustification(sAlign) };
            hasPPr = true;
        }
        if (properties.TryGetValue("spacebefore", out var sSBefore) || properties.TryGetValue("spaceBefore", out sSBefore))
        {
            var sp = stylePPr.SpacingBetweenLines ?? (stylePPr.SpacingBetweenLines = new SpacingBetweenLines());
            sp.Before = SpacingConverter.ParseWordSpacing(sSBefore).ToString();
            hasPPr = true;
        }
        if (properties.TryGetValue("spaceafter", out var sSAfter) || properties.TryGetValue("spaceAfter", out sSAfter))
        {
            var sp = stylePPr.SpacingBetweenLines ?? (stylePPr.SpacingBetweenLines = new SpacingBetweenLines());
            sp.After = SpacingConverter.ParseWordSpacing(sSAfter).ToString();
            hasPPr = true;
        }
        // BUG-DUMP-R46-1: style-level auto-spacing toggles (mirror BUG-DUMP-R44-4 paragraph path)
        if (properties.TryGetValue("spacebeforeauto", out var sSBAuto) || properties.TryGetValue("spaceBeforeAuto", out sSBAuto)
            || properties.TryGetValue("beforeautospacing", out sSBAuto))
        {
            var sp = stylePPr.SpacingBetweenLines ?? (stylePPr.SpacingBetweenLines = new SpacingBetweenLines());
            sp.BeforeAutoSpacing = OnOffValue.FromBoolean(IsTruthy(sSBAuto));
            hasPPr = true;
        }
        if (properties.TryGetValue("spaceafterauto", out var sSAAuto) || properties.TryGetValue("spaceAfterAuto", out sSAAuto)
            || properties.TryGetValue("afterautospacing", out sSAAuto))
        {
            var sp = stylePPr.SpacingBetweenLines ?? (stylePPr.SpacingBetweenLines = new SpacingBetweenLines());
            sp.AfterAutoSpacing = OnOffValue.FromBoolean(IsTruthy(sSAAuto));
            hasPPr = true;
        }
        if (properties.TryGetValue("spacebeforelines", out var sSBL) || properties.TryGetValue("spaceBeforeLines", out sSBL))
        {
            var sp = stylePPr.SpacingBetweenLines ?? (stylePPr.SpacingBetweenLines = new SpacingBetweenLines());
            sp.BeforeLines = ParseHelpers.SafeParseInt(sSBL, "spaceBeforeLines");
            hasPPr = true;
        }
        if (properties.TryGetValue("spaceafterlines", out var sSAL) || properties.TryGetValue("spaceAfterLines", out sSAL))
        {
            var sp = stylePPr.SpacingBetweenLines ?? (stylePPr.SpacingBetweenLines = new SpacingBetweenLines());
            sp.AfterLines = ParseHelpers.SafeParseInt(sSAL, "spaceAfterLines");
            hasPPr = true;
        }
        // CONSISTENCY(add-set-symmetry): mirror SetStylePath's lineSpacing case
        // (WordHandler.Set.Dispatch.cs:1403). Without this, `add /styles … --prop
        // lineSpacing=1.5x` was silent-dropped while `set /styles/X --prop
        // lineSpacing=1.5x` worked, breaking dump → batch round-trip on style
        // entries (BUG-R2-08 / BT-8).
        if (properties.TryGetValue("linespacing", out var sLineSpacing) || properties.TryGetValue("lineSpacing", out sLineSpacing))
        {
            var sp = stylePPr.SpacingBetweenLines ?? (stylePPr.SpacingBetweenLines = new SpacingBetweenLines());
            var (twips, isMultiplier) = SpacingConverter.ParseWordLineSpacing(sLineSpacing);
            sp.Line = twips.ToString();
            sp.LineRule = isMultiplier
                ? new DocumentFormat.OpenXml.EnumValue<LineSpacingRuleValues>(LineSpacingRuleValues.Auto)
                : new DocumentFormat.OpenXml.EnumValue<LineSpacingRuleValues>(LineSpacingRuleValues.Exact);
            hasPPr = true;
        }
        // BUG-019: explicit lineRule override (auto/exact/atLeast) — needed
        // because lineSpacing alone serializes AtLeast and Exact identically.
        if (properties.TryGetValue("lineRule", out var sLineRule) || properties.TryGetValue("linerule", out sLineRule))
        {
            var sp = stylePPr.SpacingBetweenLines ?? (stylePPr.SpacingBetweenLines = new SpacingBetweenLines());
            sp.LineRule = ParseLineRule(sLineRule);
            hasPPr = true;
        }
        // Reading direction: <w:bidi/> on style pPr (mirrors AddParagraph).
        // Without this, `add /styles --prop direction=rtl` either fell through
        // to the dotted-key probe (which writes <w:rtl/> on rPr but skips
        // pPr) or surfaced as UNSUPPORTED.
        // R21-fuzz-1: character styles must NOT carry pPr — w:CT_Style for
        // type=character explicitly forbids <w:pPr>. Direction on a character
        // style maps to <w:rtl/> in <w:rPr> (handled in the rPr block below
        // via sStyleRtlFlag), not <w:bidi/> in pPr.
        bool? sStyleRtlFlag = null;
        if (properties.TryGetValue("direction", out var sDirRaw)
            || properties.TryGetValue("dir", out sDirRaw)
            || properties.TryGetValue("bidi", out sDirRaw)
            || properties.TryGetValue("rtl", out sDirRaw))
        {
            var sRtl = ParseDirectionRtl(sDirRaw);
            sStyleRtlFlag = sRtl;
            if (styleType == StyleValues.Character)
            {
                // Defer to the rPr block; nothing to write on pPr.
            }
            else if (sRtl)
            {
                stylePPr.BiDi = new BiDi();
                hasPPr = true;
            }
            else
            {
                // R19-fuzz-1/2: explicit ltr on Add. If the basedOn chain
                // has bidi=true, emit <w:bidi w:val="0"/> to cancel
                // inheritance; otherwise no element (canonical clean state).
                if (properties.TryGetValue("basedOn", out var bOnRaw)
                    || properties.TryGetValue("basedon", out bOnRaw))
                {
                    if (!string.IsNullOrEmpty(bOnRaw) && StyleChainHasBidi(bOnRaw))
                    {
                        stylePPr.BiDi = new BiDi { Val = new DocumentFormat.OpenXml.OnOffValue(false) };
                        hasPPr = true;
                    }
                }
            }
        }
        if (properties.TryGetValue("tabs", out var sTabsVal) || properties.TryGetValue("tabstops", out sTabsVal))
        {
            ApplyTabsShorthand(stylePPr, sTabsVal);
            hasPPr = true;
        }
        if (hasPPr) newStyle.AppendChild(stylePPr);

        // Style run properties
        var styleRPr = new StyleRunProperties();
        bool hasRPr = false;
        // CONSISTENCY(rtl-cascade): paragraph-style direction=rtl is carried
        // ONLY on style pPr (<w:bidi/>). We deliberately do NOT stamp
        // <w:rtl/> on StyleRunProperties for paragraph styles — CT_RPr in
        // styleRPr requires <w:rFonts> as the first child (schema order),
        // and a bare <w:rtl/> there yields a 100-error validator storm in
        // real Office. The effective.direction reduction already walks
        // pPr/bidi via the style chain (see ResolveEffectiveParagraphStyleProperties),
        // so runs in paragraphs that inherit the style still resolve RTL
        // correctly. (Suppresses R7-5 regression: invalid child element 'w:rtl'.)
        //
        // R21-fuzz-1: character styles ARE the rPr-only carrier — they have
        // no pPr surface at all. <w:rtl/> goes here for type=character.
        // Insertion order is handled by sorting the rPr children at the end
        // of this block (see schema-order pass), so emitting <w:rtl/> first
        // is safe; we do not need rFonts to come first.
        if (styleType == StyleValues.Character && sStyleRtlFlag.HasValue)
        {
            // Use InsertRunPropInSchemaOrder so <w:rtl/> lands at its CT_RPr
            // position regardless of insertion order with sibling rPr children.
            InsertRunPropInSchemaOrder(styleRPr, sStyleRtlFlag.Value
                ? new RightToLeftText()
                : new RightToLeftText { Val = DocumentFormat.OpenXml.OnOffValue.FromBoolean(false) });
            hasRPr = true;
        }
        if (properties.TryGetValue("font", out var sFont))
        {
            styleRPr.RunFonts = new RunFonts { Ascii = sFont, HighAnsi = sFont, EastAsia = sFont };
            hasRPr = true;
        }
        // Per-script font split. Each w:rFonts attr is independent — Word falls
        // back through the style chain / docDefaults for any unset attr, so we
        // only write what the caller passed and leave the rest alone. Dotted
        // keys layer on top of the bare `font=` shortcut: `font=Times,
        // font.eastAsia=SimSun` produces ascii/hAnsi=Times, eastAsia=SimSun.
        bool TrySetRFontsAttr(string key, Action<RunFonts, string> assign)
        {
            if (!properties.TryGetValue(key, out var v) || string.IsNullOrEmpty(v)) return false;
            styleRPr.RunFonts ??= new RunFonts();
            assign(styleRPr.RunFonts, v);
            hasRPr = true;
            return true;
        }
        TrySetRFontsAttr("font.ascii",    (rf, v) => rf.Ascii = v);
        TrySetRFontsAttr("font.hAnsi",    (rf, v) => rf.HighAnsi = v);
        // font.ea is the canonical eastAsia key (matches run/paragraph emit and
        // the style Get readback, which BUG-R5A switched from font.eastAsia to
        // font.ea); font.eastAsia stays as a legacy alias. Without font.ea the
        // dumped style font.ea= row hit UNSUPPORTED and the eastAsia face was
        // lost on round-trip.
        TrySetRFontsAttr("font.ea",       (rf, v) => rf.EastAsia = v);
        TrySetRFontsAttr("font.eastAsia", (rf, v) => rf.EastAsia = v);
        TrySetRFontsAttr("font.cs",       (rf, v) => rf.ComplexScript = v);
        // Theme-font slots — bind the style to the theme major/minor font
        // (rFonts/@*Theme) instead of a literal face. Heading1..9 in real Word
        // templates carry these (asciiTheme="majorHAnsi", …); without them the
        // dumped style lost its <w:rFonts> and headings fell back to the body
        // font on rebuild. Mirrors the run-level theme handling in
        // WordHandler.Add.Text.cs (font.asciiTheme / font.hAnsiTheme /
        // font.eaTheme / font.csTheme + lowercase aliases).
        string? sAsciiTheme = null, sHAnsiTheme = null, sEaTheme = null, sCsTheme = null;
        if (properties.TryGetValue("font.asciiTheme", out var sAT) || properties.TryGetValue("font.asciitheme", out sAT))
            sAsciiTheme = sAT;
        if (properties.TryGetValue("font.hAnsiTheme", out var sHAT) || properties.TryGetValue("font.hansitheme", out sHAT))
            sHAnsiTheme = sHAT;
        if (properties.TryGetValue("font.eaTheme", out var sEAT) || properties.TryGetValue("font.eatheme", out sEAT) || properties.TryGetValue("font.eastasiatheme", out sEAT))
            sEaTheme = sEAT;
        if (properties.TryGetValue("font.csTheme", out var sCST) || properties.TryGetValue("font.cstheme", out sCST))
            sCsTheme = sCST;
        if (sAsciiTheme != null || sHAnsiTheme != null || sEaTheme != null || sCsTheme != null)
        {
            styleRPr.RunFonts ??= new RunFonts();
            if (!string.IsNullOrEmpty(sAsciiTheme))
                styleRPr.RunFonts.AsciiTheme = new EnumValue<ThemeFontValues>(new ThemeFontValues(sAsciiTheme));
            if (!string.IsNullOrEmpty(sHAnsiTheme))
                styleRPr.RunFonts.HighAnsiTheme = new EnumValue<ThemeFontValues>(new ThemeFontValues(sHAnsiTheme));
            if (!string.IsNullOrEmpty(sEaTheme))
                styleRPr.RunFonts.EastAsiaTheme = new EnumValue<ThemeFontValues>(new ThemeFontValues(sEaTheme));
            if (!string.IsNullOrEmpty(sCsTheme))
                styleRPr.RunFonts.ComplexScriptTheme = new EnumValue<ThemeFontValues>(new ThemeFontValues(sCsTheme));
            hasRPr = true;
        }
        if (properties.TryGetValue("size", out var sSize))
        {
            styleRPr.FontSize = new FontSize { Val = ((int)Math.Round(ParseFontSize(sSize) * 2, MidpointRounding.AwayFromZero)).ToString() };
            hasRPr = true;
        }
        // CONSISTENCY(toggle-explicit-false): a style whose rPr declares
        // <w:b w:val="0"/> overrides an inherited bold (TOC2 un-bolding
        // TOC1). Truthy-only handling dropped the explicit-off and the
        // inherited bold bled through on replay.
        if (properties.TryGetValue("bold", out var sBold))
        {
            if (IsTruthy(sBold)) { styleRPr.Bold = new Bold(); hasRPr = true; }
            else if (IsExplicitFalseAddOverride(sBold))
            {
                styleRPr.Bold = new Bold { Val = OnOffValue.FromBoolean(false) };
                hasRPr = true;
            }
        }
        if (properties.TryGetValue("italic", out var sItalic))
        {
            if (IsTruthy(sItalic)) { styleRPr.Italic = new Italic(); hasRPr = true; }
            else if (IsExplicitFalseAddOverride(sItalic))
            {
                styleRPr.Italic = new Italic { Val = OnOffValue.FromBoolean(false) };
                hasRPr = true;
            }
        }
        if (properties.TryGetValue("color", out var sColor))
        {
            // BUG-DUMP-R43-3: the style color value may carry a theme linkage as
            // a ';'-tail (themeColor=accent1;themeShade=..;themeTint=..), same
            // convention as shading/border theme tails. Split off the tail,
            // stamp w:val from the positional hex (when present), and re-apply
            // the theme attrs so a heading's color follows the theme after a swap.
            var (colorPositional, colorTheme) = ExtractThemeTail(sColor);
            var styleColor = new Color();
            if (!string.IsNullOrEmpty(colorPositional))
                styleColor.Val = SanitizeHex(colorPositional);
            ApplyColorTheme(styleColor, colorTheme);
            styleRPr.Color = styleColor;
            hasRPr = true;
        }
        // BUG-DUMP-STYLE-W14LIG: OpenType typographic toggles (ligatures /
        // numForm / numSpacing) on a STYLE's rPr were dropped on round-trip —
        // AddStyle's typed setters don't cover the w14 extension elements and the
        // generic per-key fallback (ApplyRunFormatting / GenericXmlQuery) can't
        // write a w14-namespaced element. A body style that disables ligatures
        // (`<w14:ligatures w14:val="none"/>` on Normal) then inherited
        // docDefaults' `standardContextual`, so the whole document re-rendered
        // with ligatures ON — different glyph shaping/line-wrapping that drifted
        // the layout across pages. Mirror the run-level ApplyW14ValEffect: "none"
        // is a real ST_ enum value (explicit disable), never a strip sentinel.
        foreach (var (w14Key, w14Aliases) in new[]
        {
            ("ligatures", new[] { "ligatures" }),
            ("numForm", new[] { "numForm", "numform" }),
            ("numSpacing", new[] { "numSpacing", "numspacing" }),
        })
        {
            string? w14Val = null;
            foreach (var alias in w14Aliases)
                if (properties.TryGetValue(alias, out w14Val)) break;
            if (string.IsNullOrEmpty(w14Val)) continue;
            var w14Child = WordHandler.BuildW14ValElement(w14Key, w14Val);
            if (w14Child != null)
            {
                InsertRunPropInSchemaOrder(styleRPr, w14Child);
                hasRPr = true;
            }
        }
        if (hasRPr) newStyle.AppendChild(styleRPr);

        // Numbering linkage on the style itself (numPr inside StyleParagraphProperties).
        // Lets paragraphs inherit list editing without setting numPr on each paragraph,
        // which is the canonical pattern used by Heading1..9 in real templates.
        // Mirrors WordHandler.Set.cs paragraph-level numId/ilvl handling.
        bool hasStyleNumPr = (properties.TryGetValue("numId", out var sNumIdStr) || properties.TryGetValue("numid", out sNumIdStr))
                          || (properties.TryGetValue("ilvl", out _) || properties.TryGetValue("numLevel", out _) || properties.TryGetValue("numlevel", out _));
        if (hasStyleNumPr)
        {
            var pPrForNum = newStyle.StyleParagraphProperties ?? EnsureStyleParagraphProperties(newStyle);
            var numPr = pPrForNum.NumberingProperties ?? (pPrForNum.NumberingProperties = new NumberingProperties());
            if (!string.IsNullOrEmpty(sNumIdStr))
            {
                var nid = ParseHelpers.SafeParseInt(sNumIdStr, "numId");
                if (nid < 0) throw new ArgumentException($"numId must be >= 0 (got {nid}).");
                // A STYLE's numPr→numId pointing at a num that doesn't exist in
                // /numbering is tolerated by Word (it falls back to a default
                // decimal list) and validates clean — real templates ship a
                // heading style with numId=N and an empty numbering part. Unlike
                // the paragraph-level check (Add.Text.cs), do NOT hard-fail here:
                // rejecting it broke dump→batch of such styles (the style add
                // failed → the style was never created → styles basedOn it
                // dangled → corrupt file), and dropping the numId lost the
                // heading's number on round-trip. Keep the ref; let Word
                // resolve or fall back exactly as it does for the source.
                numPr.NumberingId = new NumberingId { Val = nid };
            }
            string? ilvlRaw = null;
            if (properties.TryGetValue("ilvl", out var iRaw)
                || properties.TryGetValue("numLevel", out iRaw)
                || properties.TryGetValue("numlevel", out iRaw))
                ilvlRaw = iRaw;
            if (!string.IsNullOrEmpty(ilvlRaw))
            {
                var ilvl = ParseHelpers.SafeParseInt(ilvlRaw, "ilvl");
                // BUG-R4B(BUG2): clamp ilvl > 8 (Word tolerates it) rather than
                // reject, mirroring the paragraph add/set + style set paths so a
                // style numPr survives dump→replay.
                if (ilvl < 0 || ilvl > 8)
                {
                    var clampedIlvl = Math.Clamp(ilvl, 0, 8);
                    LastAddWarnings.Add($"ilvl {ilvl} out of OOXML range 0..8 — clamped to {clampedIlvl}");
                    ilvl = clampedIlvl;
                }
                numPr.NumberingLevelReference = new NumberingLevelReference { Val = ilvl };
            }
        }

        // CONSISTENCY(add-set-symmetry): mirror SetStylePath's ApplyRunFormatting
        // + generic OOXML fallback so `add` accepts the same prop surface as
        // `set` for any single-Val style property. Without this sweep, props
        // like underline/strike/highlight/contextualSpacing/kinsoku/snapToGrid
        // would be silently dropped on add (schema preflight lets them
        // through; AddStyle's TryGetValue list only covers ~13 keys).
        var addStyleConsumed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            // CONSISTENCY(style-dual-key): styleId / styleName are the
            // canonical readback keys Get surfaces (Round 2). The id/name
            // alias chain consumed them above; record both spellings here
            // so the per-key 'silent drop' sweep doesn't flag them as
            // unsupported even though they were honored.
            "id", "styleId", "styleid",
            "name", "styleName", "stylename",
            "type", "basedon", "basedOn", "next",
            // <w:link/> paired-style — consumed in the explicit dispatch above.
            // Without listing here, the per-key fallback loop would attempt
            // GenericXmlQuery on `link`, hit the schema's <w:link> element,
            // and either double-stamp or surface a misleading UNSUPPORTED.
            "linked", "link",
            // BUG-DUMP11-05: top-level Style flags consumed in the explicit
            // dispatch above; without listing them here, the per-key fallback
            // loop would route `hidden` to ApplyRunFormatting (vanish alias)
            // and double-stamp it on rPr.
            "autoRedefine", "autoredefine", "hidden",
            // customStyle / customstyle consumed in the explicit dispatch above
            // (sets Style.CustomStyle, BUG-DUMP-R30-1). The dump emits it on
            // every style entry (Normal included), so without listing it here
            // the per-key sweep flags every styled docx's round-trip with a
            // spurious customStyle "(not supported)" warning.
            "customStyle", "customstyle",
            // BUG-DUMP-STYLE-LATENT: latent-style flags consumed in the explicit
            // dispatch above (uiPriority/semiHidden/unhideWhenUsed/qFormat/
            // locked). Without listing them here the per-key fallback loop would
            // route them to GenericXmlQuery / ApplyRunFormatting and either
            // double-stamp or surface a misleading UNSUPPORTED warning.
            "uiPriority", "uipriority",
            "semiHidden", "semihidden",
            "unhideWhenUsed", "unhidewhenused",
            "qFormat", "qformat",
            "locked",
            // w:default="1" — consumed in the explicit dispatch above (sets
            // Style.Default). List here so the per-key sweep doesn't flag it.
            "default",
            "align", "alignment", "spacebefore", "spaceBefore",
            "spaceafter", "spaceAfter", "linespacing", "lineSpacing",
            "spacebeforelines", "spaceBeforeLines", "spaceafterlines", "spaceAfterLines",
            // auto-spacing toggles consumed in the explicit dispatch above; the
            // dump emits them on styles whose pPr carries <w:spacing
            // w:beforeAutospacing/@w:afterAutospacing>, so without listing them
            // the per-key sweep flagged every such style's round-trip.
            "spacebeforeauto", "spaceBeforeAuto", "beforeautospacing",
            "spaceafterauto", "spaceAfterAuto", "afterautospacing",
            "lineRule", "linerule",
            "font", "size", "bold", "italic", "color",
            "direction", "dir", "bidi",
            "font.ascii", "font.hAnsi", "font.eastAsia", "font.cs",
            "numId", "numid", "ilvl", "numLevel", "numlevel",
            "tabs", "tabstops",
            // BUG-DUMP-STYLE-W14LIG: w14 OpenType toggles consumed by the
            // explicit StyleRunProperties block above; without listing them the
            // per-key fallback would route them to GenericXmlQuery (which can't
            // write a w14-namespaced element) and surface a misleading UNSUPPORTED.
            "ligatures", "numForm", "numform", "numSpacing", "numspacing",
        };
        foreach (var (key, value) in properties)
        {
            if (addStyleConsumed.Contains(key)) continue;

            // ACCOUNTING(handler-as-truth): AddStyle's `foreach` over a
            // Dictionary<,>-typed parameter never fires the TrackingComparer
            // (Dictionary<>.Enumerator iterates entries[] directly). Keys that
            // succeed through the fallback probes below — ApplyRunFormatting,
            // TypedAttributeFallback, GenericXmlQuery — would otherwise leak
            // through as `unsupported_property` warnings even when the XML is
            // correctly written. ContainsKey forces a hash lookup that runs
            // the comparer, marking the key as accessed. Typos still surface:
            // anything unrecognized falls through to LastAddUnsupportedProps
            // below, which the CLI layer reports independently of Tracking.
            properties.ContainsKey(key);

            // CONSISTENCY(style-shading-pPr): paragraph/table styles can carry
            // <w:shd> on either pPr (split into val/fill/color sub-keys by
            // Navigation, folded back into "VAL;FILL[;COLOR]" compound form
            // by WordBatchEmitter) or rPr (Query.cs:683 emits compact `shading=<fill>`
            // only). Use the compound form as the signal — values that contain
            // `;` came from pPr in the source and must round-trip to pPr.
            // Compact values stay with the ApplyRunFormatting (rPr) probe below.
            if ((string.Equals(key, "shading", StringComparison.OrdinalIgnoreCase)
                 || string.Equals(key, "shd", StringComparison.OrdinalIgnoreCase))
                && value.Contains(';')
                && (styleType == StyleValues.Paragraph || styleType == StyleValues.Table))
            {
                var pPrShd = newStyle.StyleParagraphProperties ?? EnsureStyleParagraphProperties(newStyle);
                pPrShd.Shading = ParseShadingValue(value);
                continue;
            }
            // CONSISTENCY(style-snaptogrid-pPr): <w:snapToGrid> is dual-valid —
            // legal in BOTH CT_PPr and CT_RPr. On a docGrid document a paragraph
            // style's pPr-level <w:snapToGrid w:val="0"/> turns OFF grid-snapping
            // for the whole paragraph (lines keep their natural height); the dump
            // reader flattens that to the bare `snapToGrid` key. ApplyRunFormatting
            // (step 1) handles snapToGrid as an rPr toggle, so without this branch
            // the property round-tripped into rPr (character-level) and the
            // paragraph re-snapped to the grid — taller lines, cumulative drift
            // (BUG-DUMP-STYLE-SNAPGRID). Route it to pPr for paragraph/table
            // styles; character styles fall through to the rPr path below.
            if (string.Equals(key, "snapToGrid", StringComparison.OrdinalIgnoreCase)
                && (styleType == StyleValues.Paragraph || styleType == StyleValues.Table))
            {
                var pPrSnap = newStyle.StyleParagraphProperties ?? EnsureStyleParagraphProperties(newStyle);
                pPrSnap.SnapToGrid = new SnapToGrid { Val = OnOffValue.FromBoolean(IsTruthy(value)) };
                continue;
            }
            // 1) Run-formatting helper (covers underline/strike/highlight/caps/
            //    smallCaps/dstrike/vanish/shadow/emboss/imprint/noProof/rtl/
            //    superscript/subscript/charSpacing/shading/...).
            var rPrProbeAdd = new StyleRunProperties();
            if (ApplyRunFormatting(rPrProbeAdd, key, value))
            {
                ApplyRunFormatting(
                    newStyle.StyleRunProperties ?? newStyle.AppendChild(new StyleRunProperties()),
                    key, value);
                continue;
            }

            // 1b) Generic dotted "element.attr=value" fallback (e.g.
            //     ind.firstLine=240, shd.fill=FF0000, font.eastAsia=…).
            //     SDK-validated round-trip rejects unknown element/attr
            //     combinations. Runs ahead of the single-val fallback so
            //     dotted keys never accidentally get coerced into a
            //     <w:foo w:val="bar.baz"/> element.
            if (key.Contains('.'))
            {
                var pPrAttrProbe = new StyleParagraphProperties();
                if (Core.TypedAttributeFallback.TrySet(pPrAttrProbe, key, value))
                {
                    var pPrReal = newStyle.StyleParagraphProperties ?? EnsureStyleParagraphProperties(newStyle);
                    Core.TypedAttributeFallback.TrySet(pPrReal, key, value);
                    continue;
                }
                var rPrAttrProbe = new StyleRunProperties();
                if (Core.TypedAttributeFallback.TrySet(rPrAttrProbe, key, value))
                {
                    var rPrReal = newStyle.StyleRunProperties ?? newStyle.AppendChild(new StyleRunProperties());
                    Core.TypedAttributeFallback.TrySet(rPrReal, key, value);
                    continue;
                }
            }

            // 2) Generic OOXML single-Val fallback — pPr first, rPr second,
            //    matching SetStylePath's default branch. Detached probes
            //    avoid leaking empty containers on misses.
            var pPrProbeAdd = new StyleParagraphProperties();
            if (Core.GenericXmlQuery.TryCreateTypedChild(pPrProbeAdd, key, value))
            {
                Core.GenericXmlQuery.TryCreateTypedChild(
                    newStyle.StyleParagraphProperties ?? EnsureStyleParagraphProperties(newStyle),
                    key, value);
                continue;
            }
            var rPrProbeAdd2 = new StyleRunProperties();
            if (Core.GenericXmlQuery.TryCreateTypedChild(rPrProbeAdd2, key, value))
            {
                Core.GenericXmlQuery.TryCreateTypedChild(
                    newStyle.StyleRunProperties ?? newStyle.AppendChild(new StyleRunProperties()),
                    key, value);
                continue;
            }
            // CONSISTENCY(style-indent): list-family styles round-trip with
            // leftIndent / hangingIndent / firstLineIndent / rightIndent on the
            // style definition (BUG BT-5). Mirror SetStylePath's wiring so
            // dump→batch survives without losing list indents.
            switch (key.ToLowerInvariant())
            {
                case "leftindent":
                {
                    var pPrLi = newStyle.StyleParagraphProperties ?? EnsureStyleParagraphProperties(newStyle);
                    var indLi = pPrLi.Indentation ?? (pPrLi.Indentation = new Indentation());
                    // w:ind/@w:left is ST_SignedTwipsMeasure — negative indents
                    // (text extending into the margin) are valid OOXML. Use the
                    // signed parser, matching AddParagraph (BUG-DUMP-NEGIND);
                    // ParseWordSpacing's non-negative guard wrongly rejected real
                    // styles (gov/table fixtures carry negative left/right ind),
                    // failing the style add → dangling basedOn → corrupt file.
                    indLi.Left = SpacingConverter.ParseWordSpacingSigned(value).ToString();
                    continue;
                }
                case "rightindent":
                {
                    var pPrRi = newStyle.StyleParagraphProperties ?? EnsureStyleParagraphProperties(newStyle);
                    var indRi = pPrRi.Indentation ?? (pPrRi.Indentation = new Indentation());
                    indRi.Right = SpacingConverter.ParseWordSpacingSigned(value).ToString();
                    continue;
                }
                case "firstlineindent":
                {
                    var pPrFli = newStyle.StyleParagraphProperties ?? EnsureStyleParagraphProperties(newStyle);
                    var indFli = pPrFli.Indentation ?? (pPrFli.Indentation = new Indentation());
                    // w:ind/@w:firstLine is ST_TwipsMeasure (UNSIGNED) — unlike
                    // left/right (signed), it must be non-negative. Keep the
                    // non-negative parser.
                    indFli.FirstLine = SpacingConverter.ParseWordSpacing(value).ToString();
                    continue;
                }
                case "hangingindent":
                {
                    var pPrHi = newStyle.StyleParagraphProperties ?? EnsureStyleParagraphProperties(newStyle);
                    var indHi = pPrHi.Indentation ?? (pPrHi.Indentation = new Indentation());
                    indHi.Hanging = SpacingConverter.ParseWordSpacing(value).ToString();
                    continue;
                }
                case "firstlinechars":
                {
                    var pPrFlc = newStyle.StyleParagraphProperties ?? EnsureStyleParagraphProperties(newStyle);
                    var indFlc = pPrFlc.Indentation ?? (pPrFlc.Indentation = new Indentation());
                    indFlc.FirstLineChars = ParseHelpers.SafeParseInt(value, "firstLineChars");
                    continue;
                }
                case "leftchars" or "startchars":
                {
                    var pPrLc = newStyle.StyleParagraphProperties ?? EnsureStyleParagraphProperties(newStyle);
                    var indLc = pPrLc.Indentation ?? (pPrLc.Indentation = new Indentation());
                    indLc.LeftChars = ParseHelpers.SafeParseInt(value, "leftChars");
                    continue;
                }
                case "rightchars" or "endchars":
                {
                    var pPrRc = newStyle.StyleParagraphProperties ?? EnsureStyleParagraphProperties(newStyle);
                    var indRc = pPrRc.Indentation ?? (pPrRc.Indentation = new Indentation());
                    indRc.RightChars = ParseHelpers.SafeParseInt(value, "rightChars");
                    continue;
                }
                case "hangingchars":
                {
                    var pPrHc = newStyle.StyleParagraphProperties ?? EnsureStyleParagraphProperties(newStyle);
                    var indHc = pPrHc.Indentation ?? (pPrHc.Indentation = new Indentation());
                    indHc.HangingChars = ParseHelpers.SafeParseInt(value, "hangingChars");
                    continue;
                }
                case "pbdr.top" or "pbdr.bottom" or "pbdr.left" or "pbdr.right" or "pbdr.between" or "pbdr.bar" or "pbdr.all" or "pbdr":
                {
                    var pPrB = newStyle.StyleParagraphProperties ?? EnsureStyleParagraphProperties(newStyle);
                    ApplyStyleParagraphBorders(pPrB, key, value);
                    continue;
                }
            }

            // Anything still unconsumed is a genuine silent drop — composites
            // (font.eastAsia, ind.firstLine, tabs, numId, ...) that the
            // curated AddStyle does not yet model. Record so the CLI layer
            // can surface a WARNING with targeted hints instead of a silent
            // "Added" lie. See StyleUnsupportedHints for the hint catalog.
            LastAddUnsupportedProps.Add(key);
        }

        // BUG-DUMP-STYLE-RPR-ORDER: the rPr is built in two passes — the explicit
        // dispatch (rFonts/sz/color via SDK typed setters + ligatures via
        // InsertRunPropInSchemaOrder) then the per-key sweep (size.cs→szCs,
        // lang.*→lang, … via TypedAttributeFallback/GenericXmlQuery). The sweep
        // paths call SchemaOrder.Place directly, which lacks the w14-extension
        // hoist InsertRunPropInSchemaOrder has: when a <w14:ligatures> already
        // sits in the rPr, the SDK comparator sorts that unknown element to the
        // front, so Place strands a later standard child (szCs/lang) AFTER the
        // w14 block — schema-invalid (CT_RPr requires every standard child to
        // precede the w14 extension) and cascading into a fully scrambled rPr.
        // Re-seat every standard (non-w14) child through InsertRunPropInSchemaOrder
        // once at the end: its Place + w14-hoist converge on the correct CT_RPr
        // order regardless of the pre-existing (scrambled) sibling order.
        if (newStyle.StyleRunProperties is { } finalRPr)
        {
            foreach (var child in finalRPr.ChildElements
                         .Where(c => c.NamespaceUri != W14Ns).ToList())
            {
                child.Remove();
                InsertRunPropInSchemaOrder(finalRPr, child);
            }
        }

        stylesPart.Styles.AppendChild(newStyle);
        stylesPart.Styles.Save();
        InvalidateStyleIndex(); // new StyleId must be visible to FindStyleById

        var resultPath = $"/styles/{styleId}";
        return resultPath;
    }

    /// <summary>
    /// BUG-R4F-02: true when a &lt;w:num w:numId=N&gt; instance is defined in the
    /// numbering part. The dump emitter consults this to avoid emitting an
    /// `add p {numId:N}` for a paragraph whose numId is dangling (no matching
    /// &lt;w:num&gt;) — such an `add` is rejected by the Add-side dangling-numId
    /// guard, which (since `add p` is atomic) would drop the paragraph TEXT too.
    /// </summary>
    internal bool IsNumIdDefined(int numId)
    {
        if (numId <= 0) return false; // 0/-1 are no-list sentinels, not references
        var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        return numbering?.Elements<NumberingInstance>()
            .Any(n => n.NumberID?.Value == numId) ?? false;
    }

    /// <summary>
    /// Add a numbering instance (&lt;w:num&gt;) under /numbering. A num is a thin
    /// pointer that references an existing &lt;w:abstractNum&gt; via abstractNumId.
    ///
    /// Mode B (current): requires --prop abstractNumId=N pointing at an existing
    /// abstractNum. Other modes (auto-create abstractNum, lvlOverride) follow.
    /// </summary>
    private string AddNum(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var mainPart = _doc.MainDocumentPart!;
        var numberingPart = mainPart.NumberingDefinitionsPart
            ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering ??= new Numbering();
        var numbering = numberingPart.Numbering;

        // Three modes:
        //   B/C: --prop abstractNumId=N (reuse existing template; optionally with start overrides)
        //   A:   --prop format=... (no abstractNumId; auto-create a matching abstractNum)
        //   neither: throw with guidance
        bool hasAbsId = properties.TryGetValue("abstractNumId", out var absIdStr) && !string.IsNullOrEmpty(absIdStr);
        bool hasFormat = properties.ContainsKey("format")
                       || properties.ContainsKey("text")
                       || properties.ContainsKey("indent")
                       || properties.ContainsKey("type")
                       || properties.ContainsKey("name")
                       || properties.ContainsKey("styleLink")
                       || properties.ContainsKey("numStyleLink")
                       || properties.Keys.Any(k =>
                            k.StartsWith("level", StringComparison.OrdinalIgnoreCase)
                            && k.Length > 5 && char.IsDigit(k[5]));
        if (hasAbsId && hasFormat)
            throw new ArgumentException(
                "--prop abstractNumId conflicts with --prop format/text/indent/type. " +
                "Either reuse an existing template (abstractNumId) or define a new one (format/text/indent/type), not both.");
        if (!hasAbsId && !hasFormat)
            throw new ArgumentException(
                "--type num requires either --prop abstractNumId=N (reuse existing template) " +
                "or --prop format=decimal|bullet|... (auto-create a matching abstractNum).");

        int abstractNumId;
        if (hasAbsId)
        {
            abstractNumId = ParseHelpers.SafeParseInt(absIdStr!, "abstractNumId");
            // Reject pointers that would dangle — Word silently drops numbering
            // when numId resolves to a missing abstractNum, which is a confusing
            // failure mode to debug. Catch it at write time.
            var abstractExists = numbering.Elements<AbstractNum>()
                .Any(a => a.AbstractNumberId?.Value == abstractNumId);
            if (!abstractExists)
                throw new ArgumentException(
                    $"abstractNumId={abstractNumId} not found in /numbering. " +
                    "Create the abstractNum first, or pick an existing one via 'officecli query <file> abstractNum'.");
        }
        else
        {
            abstractNumId = numbering.Elements<AbstractNum>()
                .Select(a => a.AbstractNumberId?.Value ?? 0).DefaultIfEmpty(-1).Max() + 1;
            BuildAbstractNumElement(numbering, abstractNumId, properties);
        }

        // numId assignment: explicit collides → throw; otherwise max+1.
        // Mirrors AddStyle's IdTaken pattern, but numId is int (not string)
        // so there's no "auto-suffix" — just take next available.
        int numId;
        var explicitId = properties.ContainsKey("id");
        if (explicitId)
        {
            numId = ParseHelpers.SafeParseInt(properties["id"], "id");
            if (numId < 1)
                throw new ArgumentException($"numId must be >= 1 (got {numId}). numId=0 is reserved as 'no numbering'.");
            if (numbering.Elements<NumberingInstance>().Any(n => n.NumberID?.Value == numId))
                throw new ArgumentException(
                    $"numId {numId} already exists. Pick a unique --prop id, or omit --prop id for auto-assignment.");
        }
        else
        {
            numId = numbering.Elements<NumberingInstance>()
                .Select(n => n.NumberID?.Value ?? 0).DefaultIfEmpty(0).Max() + 1;
        }

        // Schema requires AbstractNum elements before NumberingInstance elements.
        // Append the new num at the end of the existing NumberingInstance run.
        var newNum = new NumberingInstance { NumberID = numId };
        newNum.AppendChild(new AbstractNumId { Val = abstractNumId });

        // Mode C: per-level start overrides. `start` is shorthand for
        // `startOverride.0`. `startOverride.N` (0..8) emits a <w:lvlOverride>
        // for that level. Each override is a fresh sibling element — no
        // collision logic needed since we're constructing a brand-new num.
        var startOverrides = new SortedDictionary<int, int>();
        if (properties.TryGetValue("start", out var startStr) && !string.IsNullOrEmpty(startStr))
            startOverrides[0] = ParseHelpers.SafeParseInt(startStr, "start");
        // Read each candidate startOverride.N (levels 0..8) via TryGetValue so
        // the TrackingPropertyDictionary marks consumed keys accessed — a plain
        // `foreach (var kvp in properties)` goes through the Dictionary<,>
        // enumerator and bypasses access tracking (the project conventions handler-as-truth).
        for (int lvl = 0; lvl <= 8; lvl++)
        {
            var key = $"startOverride.{lvl}";
            if (!properties.TryGetValue(key, out var ovrVal)) continue;
            startOverrides[lvl] = ParseHelpers.SafeParseInt(ovrVal, key);
        }
        // Reject out-of-range startOverride.N (N<0 or N>8). Word levels are
        // 0..8 only. Scan property keys for any startOverride.<num> outside
        // that band and throw — but route the detection through TryGetValue on
        // the offending key so the TrackingPropertyDictionary still marks it
        // accessed (handler-as-truth: never bypass via Keys/foreach).
        foreach (var k in properties.Keys.ToList())
        {
            if (!k.StartsWith("startOverride.", StringComparison.Ordinal)) continue;
            if (!int.TryParse(k.AsSpan("startOverride.".Length), out var n)) continue;
            if (n >= 0 && n <= 8) continue;
            // Fire access tracking on the over-range key before failing.
            properties.TryGetValue(k, out _);
            throw new ArgumentException($"startOverride level must be in range 0..8, got {n}");
        }

        // Default-restart: Word's "two num instances on the same abstractNum"
        // behavior is "continue counting" unless the new num carries an
        // explicit <w:lvlOverride><w:startOverride/></w:lvlOverride>. That
        // contradicts what API users expect ("a new num instance = independent
        // counter"), so by default we inject a startOverride on level 0 with
        // the abstractNum's level0 start value (typically 1). Users who want
        // Word's literal continuation behavior pass --prop continue=true.
        bool wantsContinue = properties.TryGetValue("continue", out var contRaw) && IsTruthy(contRaw);
        if (!wantsContinue && !startOverrides.ContainsKey(0))
        {
            var srcAbs = numbering.Elements<AbstractNum>()
                .First(a => a.AbstractNumberId?.Value == abstractNumId);
            var lvl0 = srcAbs.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == 0);
            int defaultStart = lvl0?.StartNumberingValue?.Val?.Value ?? 1;
            startOverrides[0] = defaultStart;
        }

        foreach (var (lvl, startVal) in startOverrides)
        {
            var lvlOverride = new LevelOverride { LevelIndex = lvl };
            lvlOverride.AppendChild(new StartOverrideNumberingValue { Val = startVal });
            newNum.AppendChild(lvlOverride);
        }

        // CT_Numbering order: abstractNum*, num*, numIdMacAtCleanup? — a new
        // <w:num> must precede a trailing <w:numIdMacAtCleanup> (appending after
        // it is schema-invalid "unexpected w:num child"). Insert before it when
        // present; else append. CONSISTENCY(numbering-num-before-maccleanup):
        // mirrors WordHandler.StyleList.cs.
        var newNumMacCleanup = numbering.GetFirstChild<NumberingIdMacAtCleanup>();
        if (newNumMacCleanup != null)
            numbering.InsertBefore(newNum, newNumMacCleanup);
        else
            numbering.AppendChild(newNum);
        numbering.Save();
        return $"/numbering/num[@id={numId}]";
    }

    /// <summary>
    /// Add an AbstractNum (numbering template) under /numbering. This is the
    /// definition layer — what a list "looks like": 9 levels with their
    /// own format, marker text, indent, start, justification, marker font, etc.
    ///
    /// Per-level customization via dotted keys: --prop level0.format=decimal
    /// --prop level0.text=%1. --prop level0.indent=720 ... up through level8.
    /// Bare keys (format/text/indent/start) are aliases for level0.* for
    /// backward compatibility with --type num mode A.
    ///
    /// Levels not explicitly set fall back to a sensible cycle: bullet glyphs
    /// (•/◦/▪) for bullet types, decimal/lowerLetter/lowerRoman cycle for ordered.
    /// </summary>
    private string AddAbstractNum(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var mainPart = _doc.MainDocumentPart!;
        var numberingPart = mainPart.NumberingDefinitionsPart
            ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering ??= new Numbering();
        var numbering = numberingPart.Numbering;

        int abstractNumId;
        if (properties.ContainsKey("id"))
        {
            abstractNumId = ParseHelpers.SafeParseInt(properties["id"], "id");
            if (abstractNumId < 0)
                throw new ArgumentException($"abstractNumId must be >= 0 (got {abstractNumId}).");
            if (numbering.Elements<AbstractNum>().Any(a => a.AbstractNumberId?.Value == abstractNumId))
                throw new ArgumentException(
                    $"abstractNumId {abstractNumId} already exists. Pick a unique --prop id, or omit --prop id for auto-assignment.");
        }
        else
        {
            abstractNumId = numbering.Elements<AbstractNum>()
                .Select(a => a.AbstractNumberId?.Value ?? 0).DefaultIfEmpty(-1).Max() + 1;
        }

        BuildAbstractNumElement(numbering, abstractNumId, properties);
        numbering.Save();
        return $"/numbering/abstractNum[@id={abstractNumId}]";
    }

    /// <summary>
    /// Build a fully-populated AbstractNum and insert it into Numbering in
    /// schema-correct order. Used by both the dedicated AddAbstractNum and
    /// AddNum mode A (auto-create template). Returns nothing — caller already
    /// chose abstractNumId and just needs the side effect.
    /// </summary>
    private static void BuildAbstractNumElement(Numbering numbering, int abstractNumId, Dictionary<string, string> properties)
    {
        // CONSISTENCY(numbering-alias): mirror AddLvl which accepts both
        // curated names (format/text) and OOXML attribute names
        // (numFmt/fmt/lvlText). Without these, `get level[0]` returns
        // canonical `format=…/lvlText=…` and feeding the same keys back
        // into `add abstractNum` reports UNSUPPORTED — round-trip break.
        // Normalize aliases in-place so downstream TryGetValue lookups
        // consume them (and TrackingPropertyDictionary stops flagging
        // them as unused).
        void Alias(string from, string to)
        {
            if (properties.TryGetValue(from, out var v) && !properties.ContainsKey(to))
                properties[to] = v;
        }
        Alias("fmt", "format");
        Alias("numFmt", "format");
        Alias("lvlText", "text");
        for (int aliasLvl = 0; aliasLvl < 9; aliasLvl++)
        {
            Alias($"level{aliasLvl}.fmt", $"level{aliasLvl}.format");
            Alias($"level{aliasLvl}.numFmt", $"level{aliasLvl}.format");
            Alias($"level{aliasLvl}.lvlText", $"level{aliasLvl}.text");
        }

        var abstractNum = new AbstractNum { AbstractNumberId = abstractNumId };

        // Schema order inside abstractNum:
        // nsid → multiLevelType → tmpl → name → styleLink → numStyleLink → lvl[0..8]
        var multiLevelType = properties.GetValueOrDefault("type", "hybridMultilevel").ToLowerInvariant() switch
        {
            "hybridmultilevel" or "hybrid" => MultiLevelValues.HybridMultilevel,
            "multilevel" or "multi" => MultiLevelValues.Multilevel,
            "singlelevel" or "single" => MultiLevelValues.SingleLevel,
            _ => throw new ArgumentException($"Unknown multiLevelType '{properties["type"]}'. Valid: hybridMultilevel, multilevel, singleLevel.")
        };
        abstractNum.AppendChild(new MultiLevelType { Val = multiLevelType });

        if (properties.TryGetValue("name", out var anName) && !string.IsNullOrEmpty(anName))
            abstractNum.AppendChild(new AbstractNumDefinitionName { Val = anName });
        if (properties.TryGetValue("styleLink", out var anSL) && !string.IsNullOrEmpty(anSL))
            abstractNum.AppendChild(new StyleLink { Val = anSL });
        if (properties.TryGetValue("numStyleLink", out var anNSL) && !string.IsNullOrEmpty(anNSL))
            abstractNum.AppendChild(new NumberingStyleLink { Val = anNSL });

        // Top-level format determines level fallback cycle. Bare keys map to level0
        // (backward compat: format=bullet, text=•, indent=720, start=N).
        var topFormatRaw = properties.GetValueOrDefault("format", "decimal").ToLowerInvariant();
        var topIsBullet = topFormatRaw is "bullet" or "unordered" or "ul";
        var bulletChars = new[] { "•", "◦", "▪" };

        for (int lvl = 0; lvl < 9; lvl++)
        {
            var level = new Level { LevelIndex = lvl };
            var prefix = $"level{lvl}.";

            // Per-level format with fallback cycle
            string levelFormatRaw;
            if (lvl == 0 && properties.TryGetValue("format", out var bareFmt))
                levelFormatRaw = bareFmt;
            else if (properties.TryGetValue(prefix + "format", out var perLvlFmt))
                levelFormatRaw = perLvlFmt;
            else if (topIsBullet)
                levelFormatRaw = "bullet";
            else if (properties.ContainsKey("format"))
                // Top-level format explicitly set → propagate to every level instead of
                // cycling through decimal/lowerLetter/lowerRoman, which silently turned
                // `format=decimal` into mixed numbering at level 1+.
                levelFormatRaw = topFormatRaw;
            else
                levelFormatRaw = (lvl % 3) switch { 0 => "decimal", 1 => "lowerLetter", _ => "lowerRoman" };
            var numFmt = ParseNumberFormat(levelFormatRaw);
            var isBulletAtThisLvl = numFmt.Value == NumberFormatValues.Bullet;

            // start (default 1)
            int start = 1;
            if (lvl == 0 && properties.TryGetValue("start", out var bareStart))
                start = ParseHelpers.SafeParseInt(bareStart, "start");
            else if (properties.TryGetValue(prefix + "start", out var perLvlStart))
                start = ParseHelpers.SafeParseInt(perLvlStart, prefix + "start");
            level.AppendChild(new StartNumberingValue { Val = start });
            level.AppendChild(new NumberingFormat { Val = numFmt });

            // suff (tab|space|nothing) — default tab in OOXML, omit unless overridden
            if (properties.TryGetValue(prefix + "suff", out var suffRaw) && !string.IsNullOrEmpty(suffRaw))
            {
                var suffVal = suffRaw.ToLowerInvariant() switch
                {
                    "tab" => LevelSuffixValues.Tab,
                    "space" => LevelSuffixValues.Space,
                    "nothing" or "none" => LevelSuffixValues.Nothing,
                    _ => throw new ArgumentException($"Invalid {prefix}suff '{suffRaw}'. Valid: tab, space, nothing.")
                };
                level.AppendChild(new LevelSuffix { Val = suffVal });
            }

            // lvlText
            string lvlText;
            if (lvl == 0 && properties.TryGetValue("text", out var bareText))
                lvlText = bareText;
            else if (properties.TryGetValue(prefix + "text", out var perLvlText))
                lvlText = perLvlText;
            else if (isBulletAtThisLvl)
                lvlText = bulletChars[lvl % bulletChars.Length];
            else
                lvlText = $"%{lvl + 1}.";
            level.AppendChild(new LevelText { Val = lvlText });

            // lvlJc (justification): left|center|right (default left)
            var jcRaw = properties.GetValueOrDefault(prefix + "justification",
                properties.GetValueOrDefault(prefix + "jc", "left")).ToLowerInvariant();
            var jcVal = jcRaw switch
            {
                "left" or "start" => LevelJustificationValues.Left,
                "center" => LevelJustificationValues.Center,
                "right" or "end" => LevelJustificationValues.Right,
                _ => throw new ArgumentException($"Invalid {prefix}justification '{jcRaw}'. Valid: left, center, right.")
            };
            level.AppendChild(new LevelJustification { Val = jcVal });

            // pPr/ind (indent + hanging)
            int leftIndent;
            if (lvl == 0 && properties.TryGetValue("indent", out var bareIndent))
                leftIndent = ParseHelpers.SafeParseInt(bareIndent, "indent");
            else if (properties.TryGetValue(prefix + "indent", out var perLvlIndent))
                leftIndent = ParseHelpers.SafeParseInt(perLvlIndent, prefix + "indent");
            else
                leftIndent = (lvl + 1) * 720;
            int hanging = properties.TryGetValue(prefix + "hanging", out var hangingRaw)
                ? ParseHelpers.SafeParseInt(hangingRaw, prefix + "hanging")
                : 360;
            level.AppendChild(new PreviousParagraphProperties(
                new Indentation { Left = leftIndent.ToString(), Hanging = hanging.ToString() }
            ));

            // rPr — marker font/size/color/bold/italic. Only emit when caller
            // supplied at least one rPr-relevant prop, otherwise let Word use
            // defaults (don't write a stray empty <w:rPr/>).
            bool hasRpr = properties.ContainsKey(prefix + "font")
                       || properties.ContainsKey(prefix + "size")
                       || properties.ContainsKey(prefix + "color")
                       || properties.ContainsKey(prefix + "bold")
                       || properties.ContainsKey(prefix + "italic");
            if (hasRpr)
            {
                var nspr = new NumberingSymbolRunProperties();
                // CT_RPr schema order: rFonts → b → i → color → sz.
                if (properties.TryGetValue(prefix + "font", out var fontRaw) && !string.IsNullOrEmpty(fontRaw))
                {
                    nspr.AppendChild(new RunFonts { Ascii = fontRaw, HighAnsi = fontRaw, EastAsia = fontRaw });
                }
                if (properties.TryGetValue(prefix + "bold", out var boldRaw) && IsTruthy(boldRaw))
                    nspr.AppendChild(new Bold());
                if (properties.TryGetValue(prefix + "italic", out var italRaw) && IsTruthy(italRaw))
                    nspr.AppendChild(new Italic());
                if (properties.TryGetValue(prefix + "color", out var colorRaw) && !string.IsNullOrEmpty(colorRaw))
                {
                    nspr.AppendChild(new Color { Val = SanitizeHex(colorRaw) });
                }
                if (properties.TryGetValue(prefix + "size", out var sizeRaw) && !string.IsNullOrEmpty(sizeRaw))
                {
                    var halfPt = (int)Math.Round(ParseFontSize(sizeRaw) * 2, MidpointRounding.AwayFromZero);
                    nspr.AppendChild(new FontSize { Val = halfPt.ToString() });
                }
                level.AppendChild(nspr);
            }

            abstractNum.AppendChild(level);
        }

        // Schema requires AbstractNum before NumberingInstance.
        var firstNumInstance = numbering.GetFirstChild<NumberingInstance>();
        if (firstNumInstance != null)
            numbering.InsertBefore(abstractNum, firstNumInstance);
        else
            numbering.AppendChild(abstractNum);
    }

    /// <summary>
    /// Add a single &lt;w:lvl&gt; under an existing &lt;w:abstractNum&gt;. Distinct from
    /// AddDefault → TryCreateTypedElement, which uses schema-aware AddChild and
    /// silently REPLACES any existing lvl in the same parent (data loss when a
    /// caller adds ilvl=0 then ilvl=1 — only ilvl=1 survives). This helper uses
    /// AppendChild so multiple levels coexist, validates ilvl ∈ 0..8 and
    /// start as Int32, and accepts the same per-lvl props (lvlText/format/start/
    /// indent/...) the abstractNum builder accepts via levelN.* prefix.
    /// </summary>
    private string AddLvl(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (parent is not AbstractNum abstractNum)
            throw new ArgumentException(
                $"--type lvl requires parent /numbering/abstractNum[@id=N], got {parentPath}.");

        if (!properties.TryGetValue("ilvl", out var ilvlRaw) || string.IsNullOrEmpty(ilvlRaw))
            throw new ArgumentException("--type lvl requires --prop ilvl=N (0..8).");

        // ilvl: must be integer in 0..8 (OOXML ST_DecimalNumber for lvl is 0..8).
        if (!int.TryParse(ilvlRaw, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var ilvl))
            throw new ArgumentException($"ilvl must be an integer 0..8 (got '{ilvlRaw}').");
        if (ilvl < 0 || ilvl > 8)
            throw new ArgumentException($"ilvl must be in range 0..8 (got {ilvl}).");

        // If a lvl with this ilvl already exists (typically from
        // AddAbstractNum's default lvl[0..8] pre-population), replace it in
        // place. New ilvl values are appended. The schema-aware AddChild path
        // in AddDefault collapsed every lvl onto a single slot; this dedicated
        // helper keeps siblings distinct and only swaps when ilvl matches.
        var existing = abstractNum.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == ilvl);

        // start: integer (no float, no overflow). Default 1.
        int start = 1;
        if (properties.TryGetValue("start", out var startRaw) && !string.IsNullOrEmpty(startRaw))
        {
            if (!int.TryParse(startRaw, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out start))
                throw new ArgumentException(
                    $"start must be a 32-bit integer (got '{startRaw}'). Floats and values outside Int32 range are not accepted.");
        }

        var level = new Level { LevelIndex = ilvl };

        // numFmt: default decimal. Also accept 'numFmt' alias.
        var fmtRaw = properties.GetValueOrDefault("format",
            properties.GetValueOrDefault("numFmt", "decimal"));
        var numFmt = ParseNumberFormat(fmtRaw);

        level.AppendChild(new StartNumberingValue { Val = start });
        level.AppendChild(new NumberingFormat { Val = numFmt });

        // lvlRestart (optional). CT_Lvl schema order places lvlRestart after
        // numFmt, before pStyle/isLgl/suff/lvlText.
        if (properties.TryGetValue("lvlRestart", out var lvlRestartRaw) && !string.IsNullOrEmpty(lvlRestartRaw))
        {
            if (!int.TryParse(lvlRestartRaw, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var lrV))
                throw new ArgumentException($"lvlRestart must be a 32-bit integer (got '{lvlRestartRaw}').");
            level.AppendChild(new LevelRestart { Val = lrV });
        }

        // isLgl (optional). Schema order: after pStyle, before suff/lvlText.
        if (properties.TryGetValue("isLgl", out var isLglRaw) && IsTruthy(isLglRaw))
        {
            level.AppendChild(new IsLegalNumberingStyle());
        }

        // suff (optional)
        if (properties.TryGetValue("suff", out var suffRaw) && !string.IsNullOrEmpty(suffRaw))
        {
            var suffVal = suffRaw.ToLowerInvariant() switch
            {
                "tab" => LevelSuffixValues.Tab,
                "space" => LevelSuffixValues.Space,
                "nothing" or "none" => LevelSuffixValues.Nothing,
                _ => throw new ArgumentException($"Invalid suff '{suffRaw}'. Valid: tab, space, nothing.")
            };
            level.AppendChild(new LevelSuffix { Val = suffVal });
        }

        // lvlText: accept both 'text' and 'lvlText' aliases. Default: %{ilvl+1}. for
        // ordered, • for bullet.
        string lvlText;
        if (properties.TryGetValue("lvlText", out var ltRaw) && !string.IsNullOrEmpty(ltRaw))
            lvlText = ltRaw;
        else if (properties.TryGetValue("text", out var tRaw) && !string.IsNullOrEmpty(tRaw))
            lvlText = tRaw;
        else
            lvlText = numFmt.Equals(NumberFormatValues.Bullet) ? "•" : $"%{ilvl + 1}.";
        level.AppendChild(new LevelText { Val = lvlText });

        // jc (optional)
        if (properties.TryGetValue("justification", out var jcRaw) ||
            properties.TryGetValue("jc", out jcRaw))
        {
            var jcVal = jcRaw.ToLowerInvariant() switch
            {
                "left" or "start" => LevelJustificationValues.Left,
                "center" => LevelJustificationValues.Center,
                "right" or "end" => LevelJustificationValues.Right,
                _ => throw new ArgumentException($"Invalid justification '{jcRaw}'. Valid: left, center, right.")
            };
            level.AppendChild(new LevelJustification { Val = jcVal });
        }

        // pPr/ind (optional)
        int? leftIndent = null;
        if (properties.TryGetValue("indent", out var indRaw))
        {
            if (!int.TryParse(indRaw, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var lv))
                throw new ArgumentException($"indent must be an integer in twips (got '{indRaw}').");
            leftIndent = lv;
        }
        int? hanging = null;
        if (properties.TryGetValue("hanging", out var hangRaw))
        {
            if (!int.TryParse(hangRaw, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var hv))
                throw new ArgumentException($"hanging must be an integer in twips (got '{hangRaw}').");
            hanging = hv;
        }
        // direction/dir/bidi: paragraph-level RTL on the level's pPr.
        // CONSISTENCY(canonical): same vocabulary as paragraph/section direction.
        // Only `rtl` writes <w:bidi/>; `ltr` is the canonical clear (no element)
        // — mirrors WordHandler.Helpers.cs:1220-1222 and section/paragraph add
        // semantics. Lvl pPr has no inheritance source above it (lvl is a leaf),
        // so explicit ltr never needs <w:bidi w:val=0/>.
        bool? lvlBidiOn = null;
        if (properties.TryGetValue("direction", out var dirRaw) ||
            properties.TryGetValue("dir", out dirRaw) ||
            properties.TryGetValue("bidi", out dirRaw))
        {
            lvlBidiOn = (dirRaw ?? string.Empty).ToLowerInvariant() switch
            {
                "rtl" or "righttoleft" or "right-to-left" or "true" or "1" => true,
                "ltr" or "lefttoright" or "left-to-right" or "false" or "0" or "" => false,
                _ => throw new ArgumentException($"Invalid direction value: '{dirRaw}'. Valid values: rtl, ltr.")
            };
        }

        if (leftIndent.HasValue || hanging.HasValue || lvlBidiOn == true)
        {
            var pPr = new PreviousParagraphProperties();
            if (lvlBidiOn == true) pPr.AppendChild(new BiDi());
            if (leftIndent.HasValue || hanging.HasValue)
            {
                var ind = new Indentation();
                if (leftIndent.HasValue) ind.Left = leftIndent.Value.ToString();
                if (hanging.HasValue) ind.Hanging = hanging.Value.ToString();
                pPr.AppendChild(ind);
            }
            level.AppendChild(pPr);
        }

        // BUG-R5-T2: AddLvl previously dropped font/size/color/bold/italic/
        // underline silently — they're documented for SetAbstractNumPath
        // level-scope but Add never consumed them. Mirror the Set branch
        // (NumberingSymbolRunProperties is the lvl-level rPr container).
        NumberingSymbolRunProperties? rPr = null;
        NumberingSymbolRunProperties EnsureRPr() => rPr ??= new NumberingSymbolRunProperties();
        // CT_RPr schema order: rFonts → b → i → color → sz.
        if (properties.TryGetValue("font", out var lvlFontRaw) && !string.IsNullOrEmpty(lvlFontRaw))
        {
            var rp = EnsureRPr();
            var rf = rp.GetFirstChild<RunFonts>() ?? rp.AppendChild(new RunFonts());
            rf.Ascii = lvlFontRaw;
            rf.HighAnsi = lvlFontRaw;
            rf.EastAsia = lvlFontRaw;
        }
        if (properties.TryGetValue("bold", out var lvlBoldRaw) && IsTruthy(lvlBoldRaw))
        {
            EnsureRPr().AppendChild(new Bold());
        }
        if (properties.TryGetValue("italic", out var lvlItalRaw) && IsTruthy(lvlItalRaw))
        {
            EnsureRPr().AppendChild(new Italic());
        }
        if (properties.TryGetValue("color", out var lvlColorRaw) && !string.IsNullOrEmpty(lvlColorRaw))
        {
            var rp = EnsureRPr();
            rp.AppendChild(new Color { Val = SanitizeHex(lvlColorRaw) });
        }
        if (properties.TryGetValue("size", out var lvlSizeRaw) && !string.IsNullOrEmpty(lvlSizeRaw))
        {
            var rp = EnsureRPr();
            var halfPt = (int)Math.Round(ParseFontSize(lvlSizeRaw) * 2, MidpointRounding.AwayFromZero);
            rp.AppendChild(new FontSize { Val = halfPt.ToString() });
        }
        if (properties.TryGetValue("underline", out var lvlUnderRaw) && !string.IsNullOrEmpty(lvlUnderRaw))
        {
            var u = new Underline();
            if (IsTruthy(lvlUnderRaw)) u.Val = UnderlineValues.Single;
            else if (string.Equals(lvlUnderRaw, "double", StringComparison.OrdinalIgnoreCase)) u.Val = UnderlineValues.Double;
            else if (string.Equals(lvlUnderRaw, "none", StringComparison.OrdinalIgnoreCase) || string.Equals(lvlUnderRaw, "false", StringComparison.OrdinalIgnoreCase)) u.Val = UnderlineValues.None;
            else u.Val = UnderlineValues.Single;
            EnsureRPr().AppendChild(u);
        }
        if (rPr != null) level.AppendChild(rPr);

        // CRITICAL: AppendChild — NOT AddChild. Schema-aware AddChild treats
        // <w:lvl> as a single-instance child slot (the SDK's metadata says
        // "lvl[0..8]" but its schema model still flags them all as the same
        // child kind), so it would silently replace whatever lvl already
        // exists. AppendChild keeps every level distinct.
        if (existing != null)
        {
            existing.InsertBeforeSelf(level);
            existing.Remove();
        }
        else
        {
            abstractNum.AppendChild(level);
        }

        var numberingPart = _doc.MainDocumentPart?.NumberingDefinitionsPart;
        numberingPart?.Numbering?.Save();

        var absId = abstractNum.AbstractNumberId?.Value ?? 0;
        return $"/numbering/abstractNum[@id={absId}]/lvl[@ilvl={ilvl}]";
    }

    // Resolve the SectionProperties that a header/footer reference should
    // attach to, based on the parent path. `/section[N]` targets the carrier
    // paragraph's sectPr (mirrors NavigateToElement); `/`, `/body`, or any
    // other path falls back to the body-level (final) sectPr.
    private SectionProperties? ResolveTargetSectPrForHeaderFooter(string parentPath)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return null;
        if (!string.IsNullOrEmpty(parentPath))
        {
            var m = System.Text.RegularExpressions.Regex.Match(
                parentPath, @"^/section\[(\d+)\]/?$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (m.Success && int.TryParse(m.Groups[1].Value, out var n))
            {
                var sectParas = body.Elements<Paragraph>()
                    .Where(p => p.ParagraphProperties?.GetFirstChild<SectionProperties>() != null)
                    .ToList();
                if (n >= 1 && n <= sectParas.Count)
                    return sectParas[n - 1].ParagraphProperties!.GetFirstChild<SectionProperties>();
            }
        }
        return body.Elements<SectionProperties>().LastOrDefault();
    }

    private string AddHeader(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var mainPartH = _doc.MainDocumentPart!;

        // Resolve requested header type first, so we can reject duplicates before
        // creating an orphaned HeaderPart.
        var preHeaderType = HeaderFooterValues.Default;
        if (properties.TryGetValue("type", out var preHTypeStr) ||
            properties.TryGetValue("kind", out preHTypeStr) ||
            properties.TryGetValue("ref", out preHTypeStr))
        {
            preHeaderType = preHTypeStr.ToLowerInvariant() switch
            {
                "first" => HeaderFooterValues.First,
                "even" => HeaderFooterValues.Even,
                "default" => HeaderFooterValues.Default,
                _ => throw new ArgumentException($"Invalid header type: '{preHTypeStr}'. Valid values: default, first, even.")
            };
        }
        var preSectPr = ResolveTargetSectPrForHeaderFooter(parentPath);
        // BUG-R5B(BUG1): an orphaned reference (r:id resolves to no part) is a
        // round-trip leftover, not a genuine duplicate — repoint it instead of
        // throwing so the header CONTENT is created. A reference that resolves
        // to a real part is a true duplicate and still rejected.
        var orphanHeaderRef = FindOrphanedHeaderReference(preSectPr, preHeaderType);
        if (orphanHeaderRef == null
            && preSectPr != null && preSectPr.Elements<HeaderReference>()
                .Any(r => r.Type != null && r.Type.Value == preHeaderType))
        {
            throw new ArgumentException(
                $"Header of type '{HeaderFooterTypeName(preHeaderType)}' already exists in this section. " +
                "Remove the existing one first or use --prop type=<first|even>. " +
                "(If you added a watermark, it creates default/first/even headers automatically — " +
                "use 'set /watermark --prop text=...' to change watermark text, or 'remove /watermark' then re-add.)");
        }

        var headerPart = mainPartH.AddNewPart<HeaderPart>();

        var hPara = new Paragraph();
        AssignParaId(hPara);
        var hPProps = new ParagraphProperties();

        if (properties.TryGetValue("align", out var hAlign) || properties.TryGetValue("alignment", out hAlign) || properties.TryGetValue("jc", out hAlign))
            hPProps.Justification = new Justification { Val = ParseJustification(hAlign) };
        // Reading direction (Arabic / Hebrew). Parsed here, applied at the
        // end of paragraph build via ApplyDirectionCascade (cascades to all
        // runs including text and field runs). See WordHandler.I18n.cs.
        bool? hRtlFlag = null;
        if (properties.TryGetValue("direction", out var hDirRaw)
            || properties.TryGetValue("dir", out hDirRaw)
            || properties.TryGetValue("bidi", out hDirRaw))
        {
            hRtlFlag = ParseDirectionRtl(hDirRaw);
        }
        hPara.AppendChild(hPProps);

        // Build shared run properties for text and field runs
        RunProperties? hSharedRProps = null;
        if (properties.ContainsKey("font") || properties.ContainsKey("size") ||
            properties.ContainsKey("bold") || properties.ContainsKey("italic") || properties.ContainsKey("color"))
        {
            hSharedRProps = new RunProperties();
            if (properties.TryGetValue("font", out var hFont))
                hSharedRProps.AppendChild(new RunFonts { Ascii = hFont, HighAnsi = hFont, EastAsia = hFont });
            if (properties.TryGetValue("size", out var hSize))
                hSharedRProps.AppendChild(new FontSize { Val = ((int)Math.Round(ParseFontSize(hSize) * 2, MidpointRounding.AwayFromZero)).ToString() });
            if (properties.TryGetValue("bold", out var hBold) && IsTruthy(hBold))
                hSharedRProps.Bold = new Bold();
            if (properties.TryGetValue("italic", out var hItalic) && IsTruthy(hItalic))
                hSharedRProps.Italic = new Italic();
            if (properties.TryGetValue("color", out var hColor))
                hSharedRProps.Color = new Color { Val = SanitizeHex(hColor) };
        }

        if (properties.TryGetValue("text", out var hText))
        {
            OfficeCli.Core.ParseHelpers.ValidateXmlText(hText, "text");
            var hRun = new Run();
            if (hSharedRProps != null) hRun.AppendChild((RunProperties)hSharedRProps.CloneNode(true));
            hRun.AppendChild(new Text(hText) { Space = SpaceProcessingModeValues.Preserve });
            hPara.AppendChild(hRun);
        }

        // Support field=page|numpages|date etc. — generates fldChar complex field
        if (properties.TryGetValue("field", out var hFieldType))
        {
            var hFieldInstr = hFieldType.ToLowerInvariant() switch
            {
                "page" or "pagenum" or "pagenumber" => " PAGE ",
                "numpages" => " NUMPAGES ",
                "date" => " DATE \\@ \"yyyy-MM-dd\" ",
                "author" => " AUTHOR ",
                "title" => " TITLE ",
                "time" => " TIME ",
                "filename" => " FILENAME ",
                _ => $" {hFieldType.ToUpperInvariant()} "
            };
            var hBeginRun = new Run(new FieldChar { FieldCharType = FieldCharValues.Begin });
            var hInstrRun = new Run(new FieldCode(hFieldInstr) { Space = SpaceProcessingModeValues.Preserve });
            var hSepRun = new Run(new FieldChar { FieldCharType = FieldCharValues.Separate });
            var hResultRun = new Run(new Text("1") { Space = SpaceProcessingModeValues.Preserve });
            var hEndRun = new Run(new FieldChar { FieldCharType = FieldCharValues.End });
            if (hSharedRProps != null)
            {
                hBeginRun.PrependChild((RunProperties)hSharedRProps.CloneNode(true));
                hInstrRun.PrependChild((RunProperties)hSharedRProps.CloneNode(true));
                hSepRun.PrependChild((RunProperties)hSharedRProps.CloneNode(true));
                hResultRun.PrependChild((RunProperties)hSharedRProps.CloneNode(true));
                hEndRun.PrependChild((RunProperties)hSharedRProps.CloneNode(true));
            }
            hPara.AppendChild(hBeginRun);
            hPara.AppendChild(hInstrRun);
            hPara.AppendChild(hSepRun);
            hPara.AppendChild(hResultRun);
            hPara.AppendChild(hEndRun);
        }

        // CONSISTENCY(rtl-cascade): apply after all runs (text + field) are
        // appended so every run gets <w:rtl/>. Previously field runs were
        // missed by the inline stamp. See WordHandler.I18n.cs.
        if (hRtlFlag.HasValue)
            ApplyDirectionCascade(hPara, hRtlFlag.Value);

        // AssignParaId stamps w14:paraId / w14:textId on each w:p. Those
        // attributes are MS-2010 extensions and OpenXmlValidator rejects
        // them with Sch_UndeclaredAttribute unless the part declares the
        // w14 namespace and lists it in mc:Ignorable. The body part
        // (document.xml) does this at the document root; header/footer
        // parts need the same so paragraphs validated independently
        // accept the extension attrs.
        var hRoot = new Header(hPara);
        hRoot.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        hRoot.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
        hRoot.SetAttribute(new OpenXmlAttribute("Ignorable", "http://schemas.openxmlformats.org/markup-compatibility/2006", "w14"));
        headerPart.Header = hRoot;
        headerPart.Header.Save();

        var hBody = mainPartH.Document!.Body!;
        var hSectPr = ResolveTargetSectPrForHeaderFooter(parentPath)
            ?? hBody.AppendChild(new SectionProperties());

        var headerType = preHeaderType;

        // BUG-R5B(BUG1): repoint an orphaned reference to the newly-created part
        // rather than prepending a second reference of the same type.
        var orphanHeaderRefToRewire = FindOrphanedHeaderReference(hSectPr, headerType);
        if (orphanHeaderRefToRewire != null)
        {
            orphanHeaderRefToRewire.Id = mainPartH.GetIdOfPart(headerPart);
        }
        else
        {
            var headerRef = new HeaderReference
            {
                Id = mainPartH.GetIdOfPart(headerPart),
                Type = headerType
            };
            hSectPr.PrependChild(headerRef);
        }

        if (headerType == HeaderFooterValues.First)
        {
            // UX convenience: stamp <w:titlePg/> so the first-page header is
            // actually rendered (Word ignores type="first" headerRef when
            // titlePg is absent). Round-trip emit short-circuits this with
            // `noTitlePg=true` so a source whose sectPr lacks <w:titlePg/>
            // round-trips with the same shape — emitting `titlePage=true`
            // there would surface a phantom key the source never had.
            bool skipTitlePg = properties.TryGetValue("notitlepg", out var ntp)
                            || properties.TryGetValue("noTitlePg", out ntp);
            if (!(skipTitlePg && IsTruthy(ntp))
                && hSectPr.GetFirstChild<TitlePage>() == null)
                hSectPr.AddChild(new TitlePage(), throwOnError: false);
        }
        // CONSISTENCY(headerfooter-effective-toggle): mirror the type=first
        // → titlePg auto-write pattern. Without /settings/evenAndOddHeaders,
        // Word silently ignores the even header reference at render time.
        if (headerType == HeaderFooterValues.Even)
        {
            // CONSISTENCY(headerfooter-noEvenAndOdd-opt-out): dump→batch emits
            // `noEvenAndOddHeaders=true` when the source's settings.xml lacks
            // <w:evenAndOddHeaders/> so the auto-stamp doesn't phantom-write a
            // toggle the source never had.
            bool skipHEvenAndOdd = properties.TryGetValue("noevenandoddheaders", out var hNeo)
                                || properties.TryGetValue("noEvenAndOddHeaders", out hNeo);
            if (!(skipHEvenAndOdd && IsTruthy(hNeo)))
            {
                var hSettingsPart = mainPartH.DocumentSettingsPart
                    ?? mainPartH.AddNewPart<DocumentSettingsPart>();
                hSettingsPart.Settings ??= new Settings();
                if (hSettingsPart.Settings.GetFirstChild<EvenAndOddHeaders>() == null)
                    hSettingsPart.Settings.AddChild(new EvenAndOddHeaders(), throwOnError: false);
                hSettingsPart.Settings.Save();
            }
        }

        var hIdx = mainPartH.HeaderParts.ToList().IndexOf(headerPart);
        return $"/header[{hIdx + 1}]";
    }

    private string AddFooter(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var mainPartF = _doc.MainDocumentPart!;

        // Resolve requested footer type first, so we can reject duplicates before
        // creating an orphaned FooterPart.
        var preFooterType = HeaderFooterValues.Default;
        if (properties.TryGetValue("type", out var preFTypeStr) ||
            properties.TryGetValue("kind", out preFTypeStr) ||
            properties.TryGetValue("ref", out preFTypeStr))
        {
            preFooterType = preFTypeStr.ToLowerInvariant() switch
            {
                "first" => HeaderFooterValues.First,
                "even" => HeaderFooterValues.Even,
                "default" => HeaderFooterValues.Default,
                _ => throw new ArgumentException($"Invalid footer type: '{preFTypeStr}'. Valid values: default, first, even.")
            };
        }
        var preFSectPr = ResolveTargetSectPrForHeaderFooter(parentPath);
        // BUG-R5B(BUG1): orphaned reference (r:id resolves to no part) ⇒ repoint
        // instead of throwing so the footer CONTENT survives round-trip; a
        // reference resolving to a real part is still a genuine duplicate.
        var orphanFooterRef = FindOrphanedFooterReference(preFSectPr, preFooterType);
        if (orphanFooterRef == null
            && preFSectPr != null && preFSectPr.Elements<FooterReference>()
                .Any(r => r.Type != null && r.Type.Value == preFooterType))
        {
            throw new ArgumentException(
                $"Footer of type '{HeaderFooterTypeName(preFooterType)}' already exists in this section. " +
                "Remove the existing one first or use --prop type=<first|even>.");
        }

        var footerPart = mainPartF.AddNewPart<FooterPart>();

        var fPara = new Paragraph();
        AssignParaId(fPara);
        var fPProps = new ParagraphProperties();

        if (properties.TryGetValue("align", out var fAlign) || properties.TryGetValue("alignment", out fAlign) || properties.TryGetValue("jc", out fAlign))
            fPProps.Justification = new Justification { Val = ParseJustification(fAlign) };
        // Reading direction (Arabic / Hebrew) — mirrors AddHeader. Applied
        // at end of paragraph build via ApplyDirectionCascade.
        bool? fRtlFlag = null;
        if (properties.TryGetValue("direction", out var fDirRaw)
            || properties.TryGetValue("dir", out fDirRaw)
            || properties.TryGetValue("bidi", out fDirRaw))
        {
            fRtlFlag = ParseDirectionRtl(fDirRaw);
        }
        fPara.AppendChild(fPProps);

        // Build shared run properties for text and field runs
        RunProperties? sharedRProps = null;
        if (properties.ContainsKey("font") || properties.ContainsKey("size") ||
            properties.ContainsKey("bold") || properties.ContainsKey("italic") || properties.ContainsKey("color"))
        {
            sharedRProps = new RunProperties();
            if (properties.TryGetValue("font", out var fFont))
                sharedRProps.AppendChild(new RunFonts { Ascii = fFont, HighAnsi = fFont, EastAsia = fFont });
            if (properties.TryGetValue("size", out var fSize))
                sharedRProps.AppendChild(new FontSize { Val = ((int)Math.Round(ParseFontSize(fSize) * 2, MidpointRounding.AwayFromZero)).ToString() });
            if (properties.TryGetValue("bold", out var fBold) && IsTruthy(fBold))
                sharedRProps.Bold = new Bold();
            if (properties.TryGetValue("italic", out var fItalic) && IsTruthy(fItalic))
                sharedRProps.Italic = new Italic();
            if (properties.TryGetValue("color", out var fColor))
                sharedRProps.Color = new Color { Val = SanitizeHex(fColor) };
        }

        if (properties.TryGetValue("text", out var fText))
        {
            OfficeCli.Core.ParseHelpers.ValidateXmlText(fText, "text");
            var fRun = new Run();
            if (sharedRProps != null) fRun.AppendChild((RunProperties)sharedRProps.CloneNode(true));
            fRun.AppendChild(new Text(fText) { Space = SpaceProcessingModeValues.Preserve });
            fPara.AppendChild(fRun);
        }

        // Support field=page|numpages|date etc. — generates fldChar complex field
        if (properties.TryGetValue("field", out var fieldType))
        {
            var fieldInstr = fieldType.ToLowerInvariant() switch
            {
                "page" or "pagenum" or "pagenumber" => " PAGE ",
                "numpages" => " NUMPAGES ",
                "date" => " DATE \\@ \"yyyy-MM-dd\" ",
                "author" => " AUTHOR ",
                "title" => " TITLE ",
                "time" => " TIME ",
                "filename" => " FILENAME ",
                _ => $" {fieldType.ToUpperInvariant()} "
            };
            var beginRun = new Run(new FieldChar { FieldCharType = FieldCharValues.Begin });
            var instrRun = new Run(new FieldCode(fieldInstr) { Space = SpaceProcessingModeValues.Preserve });
            var sepRun = new Run(new FieldChar { FieldCharType = FieldCharValues.Separate });
            var resultRun = new Run(new Text("1") { Space = SpaceProcessingModeValues.Preserve });
            var endRun = new Run(new FieldChar { FieldCharType = FieldCharValues.End });
            if (sharedRProps != null)
            {
                beginRun.PrependChild((RunProperties)sharedRProps.CloneNode(true));
                instrRun.PrependChild((RunProperties)sharedRProps.CloneNode(true));
                sepRun.PrependChild((RunProperties)sharedRProps.CloneNode(true));
                resultRun.PrependChild((RunProperties)sharedRProps.CloneNode(true));
                endRun.PrependChild((RunProperties)sharedRProps.CloneNode(true));
            }
            fPara.AppendChild(beginRun);
            fPara.AppendChild(instrRun);
            fPara.AppendChild(sepRun);
            fPara.AppendChild(resultRun);
            fPara.AppendChild(endRun);
        }

        // CONSISTENCY(rtl-cascade): mirror AddHeader — apply after all runs.
        if (fRtlFlag.HasValue)
            ApplyDirectionCascade(fPara, fRtlFlag.Value);

        // Same w14 / mc:Ignorable declaration as AddHeader: paragraphs
        // here also carry w14:paraId / w14:textId from AssignParaId, and
        // OpenXmlValidator rejects them as undeclared without this.
        var fRoot = new Footer(fPara);
        fRoot.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        fRoot.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
        fRoot.SetAttribute(new OpenXmlAttribute("Ignorable", "http://schemas.openxmlformats.org/markup-compatibility/2006", "w14"));
        footerPart.Footer = fRoot;
        footerPart.Footer.Save();

        var fBody = mainPartF.Document!.Body!;
        var fSectPr = ResolveTargetSectPrForHeaderFooter(parentPath)
            ?? fBody.AppendChild(new SectionProperties());

        var footerType = preFooterType;

        // BUG-R5B(BUG1): repoint an orphaned reference to the newly-created part
        // rather than inserting a second reference of the same type.
        var orphanFooterRefToRewire = FindOrphanedFooterReference(fSectPr, footerType);
        if (orphanFooterRefToRewire != null)
        {
            orphanFooterRefToRewire.Id = mainPartF.GetIdOfPart(footerPart);
        }
        else
        {
            var footerRef = new FooterReference
            {
                Id = mainPartF.GetIdOfPart(footerPart),
                Type = footerType
            };
            // Insert footerReference after the last headerReference to maintain schema order
            var lastHeaderRef = fSectPr.Elements<HeaderReference>().LastOrDefault();
            if (lastHeaderRef != null)
                fSectPr.InsertAfter(footerRef, lastHeaderRef);
            else
                fSectPr.PrependChild(footerRef);
        }

        if (footerType == HeaderFooterValues.First)
        {
            // CONSISTENCY(headerfooter-noTitlePg-opt-out): mirror AddHeader.
            // Round-trip emit passes `noTitlePg=true` when the source's sectPr
            // had no <w:titlePg/>, so the footer add does not phantom-stamp it.
            bool skipFooterTitlePg = properties.TryGetValue("notitlepg", out var fNtp)
                                  || properties.TryGetValue("noTitlePg", out fNtp);
            if (!(skipFooterTitlePg && IsTruthy(fNtp))
                && fSectPr.GetFirstChild<TitlePage>() == null)
                fSectPr.AddChild(new TitlePage(), throwOnError: false);
        }
        // CONSISTENCY(headerfooter-effective-toggle): even-footer also needs
        // settings.xml/w:evenAndOddHeaders to render.
        if (footerType == HeaderFooterValues.Even)
        {
            bool skipFEvenAndOdd = properties.TryGetValue("noevenandoddheaders", out var fNeo)
                                || properties.TryGetValue("noEvenAndOddHeaders", out fNeo);
            if (!(skipFEvenAndOdd && IsTruthy(fNeo)))
            {
                var fSettingsPart = mainPartF.DocumentSettingsPart
                    ?? mainPartF.AddNewPart<DocumentSettingsPart>();
                fSettingsPart.Settings ??= new Settings();
                if (fSettingsPart.Settings.GetFirstChild<EvenAndOddHeaders>() == null)
                    fSettingsPart.Settings.AddChild(new EvenAndOddHeaders(), throwOnError: false);
                fSettingsPart.Settings.Save();
            }
        }

        var fIdx = mainPartF.FooterParts.ToList().IndexOf(footerPart);
        return $"/footer[{fIdx + 1}]";
    }

    /// <summary>
    /// Apply colWidths / colSpaces non-equal-width column layout to a sectPr.
    /// Mirrors the Get-side emit form in WordHandler.Query.cs (separate widths
    /// and spaces lists). Without this, Get emits `colWidths="4000,6000"` +
    /// `colSpaces="720,0"` but Add/Set silently dropped both — round-trip
    /// loses the non-equal layout. Accepts either key alone (the missing side
    /// fills with the existing col children's value or zero).
    /// </summary>
    private static void ApplySectionColumnWidthsSpaces(
        Dictionary<string, string> properties, SectionProperties sectPr)
    {
        bool hasW = properties.TryGetValue("colWidths", out var widthsVal)
                 || properties.TryGetValue("colwidths", out widthsVal);
        bool hasS = properties.TryGetValue("colSpaces", out var spacesVal)
                 || properties.TryGetValue("colspaces", out spacesVal);
        if (!hasW && !hasS) return;

        var cols = EnsureSectPrChild<Columns>(sectPr);
        var widths = hasW
            ? widthsVal!.Split(',').Select(s => ParseTwips(s.Trim()).ToString()).ToList()
            : null;
        var spaces = hasS
            ? spacesVal!.Split(',').Select(s => ParseTwips(s.Trim()).ToString()).ToList()
            : null;

        // Preserve the orthogonal axis from any existing <w:col> children so
        // per-prop Set dispatch (one --prop colWidths followed by one --prop
        // colSpaces, called separately) doesn't wipe the first axis on the
        // second call.
        var existing = cols.Elements<Column>().ToList();

        int n = widths?.Count ?? spaces!.Count;
        if (widths != null && spaces != null && widths.Count != spaces.Count)
            throw new ArgumentException(
                $"colWidths and colSpaces must have the same length (got {widths.Count} vs {spaces.Count}).");

        cols.RemoveAllChildren<Column>();
        for (int i = 0; i < n; i++)
        {
            var col = new Column();
            if (widths != null)
                col.Width = widths[i];
            else if (i < existing.Count && existing[i].Width?.Value is string w)
                col.Width = w;
            if (spaces != null)
                col.Space = spaces[i];
            else if (i < existing.Count && existing[i].Space?.Value is string sp)
                col.Space = sp;
            cols.AppendChild(col);
        }
        cols.ColumnCount = (short)n;
        // Explicit per-column widths imply non-equal layout per OOXML.
        cols.EqualWidth = false;
    }
}
