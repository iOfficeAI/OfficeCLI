// Copyright 2025 OfficeCli (officecli.ai)
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
    // BUG-TESTER fuzz-2: bound regex match time on user-supplied find patterns to
    // prevent catastrophic-backtracking DoS (e.g. "(a+)+b" against long inputs).
    private static readonly TimeSpan FindRegexMatchTimeout = TimeSpan.FromSeconds(5);

    /// <summary>
    /// Resolve the OpenXmlPart that owns a given element. Returns the
    /// HeaderPart/FooterPart/FootnotesPart/EndnotesPart/CommentsPart when the
    /// element lives inside one of those parts, falling back to MainDocumentPart.
    /// Used for part-local relationships like hyperlinks that must be added to
    /// the host part's rels file (e.g. word/_rels/header1.xml.rels) rather than
    /// the document rels.
    /// </summary>
    private OpenXmlPart ResolveHostPart(OpenXmlElement element)
    {
        var main = _doc.MainDocumentPart!;
        // Walk to the part-root element (Header/Footer/Footnotes/Endnotes/Comments/Document)
        var hdr = element.Ancestors<Header>().FirstOrDefault();
        if (hdr != null)
        {
            var hp = main.HeaderParts.FirstOrDefault(p => ReferenceEquals(p.Header, hdr));
            if (hp != null) return hp;
        }
        var ftr = element.Ancestors<Footer>().FirstOrDefault();
        if (ftr != null)
        {
            var fp = main.FooterParts.FirstOrDefault(p => ReferenceEquals(p.Footer, ftr));
            if (fp != null) return fp;
        }
        // Footnote/Endnote: parts live on MainDocumentPart.FootnotesPart / EndnotesPart
        if (element.Ancestors<Footnote>().Any() && main.FootnotesPart != null)
            return main.FootnotesPart;
        if (element.Ancestors<Endnote>().Any() && main.EndnotesPart != null)
            return main.EndnotesPart;
        return main;
    }

    /// <summary>
    /// Resolve a hyperlink relationship by id, searching the element's host
    /// part first, then falling back to MainDocumentPart and other host parts.
    /// </summary>
    private HyperlinkRelationship? ResolveHyperlinkRelationship(OpenXmlElement element, string relId)
    {
        var host = ResolveHostPart(element);
        var rel = host.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
        if (rel != null) return rel;
        // Fallback: scan MainDocumentPart and all header/footer parts (handles
        // documents authored with rels in unexpected places).
        var main = _doc.MainDocumentPart!;
        if (!ReferenceEquals(host, main))
        {
            rel = main.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
            if (rel != null) return rel;
        }
        foreach (var hp in main.HeaderParts)
        {
            if (ReferenceEquals(hp, host)) continue;
            rel = hp.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
            if (rel != null) return rel;
        }
        foreach (var fp in main.FooterParts)
        {
            if (ReferenceEquals(fp, host)) continue;
            rel = fp.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
            if (rel != null) return rel;
        }
        return null;
    }

    // ==================== Private Helpers ====================

    /// <summary>
    /// Format twips as a human-readable cm string (e.g., "21cm").
    /// 1 inch = 1440 twips, 1 inch = 2.54 cm.
    /// </summary>
    private static string FormatTwipsToCm(uint twips)
    {
        var cm = twips * 2.54 / 1440.0;
        return $"{cm:0.##}cm";
    }

    private static bool IsTruthy(string? value) =>
        ParseHelpers.IsTruthy(value);

    /// <summary>
    /// Normalize a user-provided underline token to a valid Word OOXML UnderlineValues enum string.
    /// Accepts common aliases (wavy → wave, dashdot → dotDash, etc.) plus truthy/none.
    /// </summary>
    internal static string NormalizeUnderlineValue(string value)
    {
        var v = (value ?? "").Trim();
        var mapped = v.ToLowerInvariant() switch
        {
            "true" or "single" or "1" => "single",
            "false" or "none" or "0" or "" => "none",
            "double" => "double",
            "thick" => "thick",
            "dotted" => "dotted",
            "dottedheavy" or "dotted-heavy" or "dotted_heavy" => "dottedHeavy",
            "dash" or "dashed" => "dash",
            "dashedheavy" or "dashheavy" => "dashedHeavy",
            "dashlong" or "longdash" => "dashLong",
            "dashlongheavy" or "longdashheavy" => "dashLongHeavy",
            // Word uses "dotDash" and "dashDotHeavy" (note asymmetric casing in OOXML spec).
            "dotdash" or "dashdot" => "dotDash",
            "dotdashheavy" or "dashdotheavy" => "dashDotHeavy",
            "dotdotdash" or "dashdotdot" => "dotDotDash",
            "dotdotdashheavy" or "dashdotdotheavy" => "dashDotDotHeavy",
            "wave" or "wavy" => "wave",
            "waveheavy" or "wavyheavy" => "wavyHeavy",
            "wavedouble" or "wavydouble" or "doublewave" => "wavyDouble",
            "words" or "word" => "words",
            _ => v  // pass-through for already-valid OOXML tokens
        };
        // CONSISTENCY(allowlist): mirror tab val/leader allowlist (R1 a1554d59) and
        // ParseJustification — validate before handing off to OpenXML SDK to avoid
        // leaking "specified value is not valid according to the specified enum type".
        if (!ValidUnderlineValues.Contains(mapped))
            throw new ArgumentException(
                $"Invalid underline value: '{value}'. Valid values: single, double, thick, dotted, " +
                "dottedHeavy, dash, dashedHeavy, dashLong, dashLongHeavy, dotDash, dashDotHeavy, " +
                "dotDotDash, dashDotDotHeavy, wave, wavyHeavy, wavyDouble, words, none.");
        return mapped;
    }

    private static readonly HashSet<string> ValidUnderlineValues = new(StringComparer.Ordinal)
    {
        "single", "double", "thick", "dotted", "dottedHeavy",
        "dash", "dashedHeavy", "dashLong", "dashLongHeavy",
        "dotDash", "dashDotHeavy", "dotDotDash", "dashDotDotHeavy",
        "wave", "wavyHeavy", "wavyDouble", "words", "none"
    };

    private static JustificationValues ParseJustification(string value) =>
        value.ToLowerInvariant() switch
        {
            "left" => JustificationValues.Left,
            "center" => JustificationValues.Center,
            "right" => JustificationValues.Right,
            "justify" or "both" => JustificationValues.Both,
            _ => throw new ArgumentException($"Invalid alignment value: '{value}'. Valid values: left, center, right, justify.")
        };

    /// <summary>
    /// Sanitize a hex color for Word OOXML (ST_HexColorRGB = exactly 6-char RGB).
    /// Strips # prefix, uppercases, and handles 8-char AARRGGBB by extracting RGB portion.
    /// </summary>
    private static string SanitizeHex(string value) =>
        ParseHelpers.SanitizeColorForOoxml(value).Rgb;

    /// <summary>
    /// Parse a highlight color name, throwing ArgumentException with valid options on failure.
    /// </summary>
    private static readonly HashSet<string> ValidHighlightColors = new(StringComparer.OrdinalIgnoreCase)
    {
        "yellow", "green", "cyan", "magenta", "blue", "red",
        "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow",
        "darkGray", "lightGray", "black", "white", "none"
    };

    private static HighlightColorValues ParseHighlightColor(string value)
    {
        if (!ValidHighlightColors.Contains(value))
            throw new ArgumentException(
                $"Invalid 'highlight' value '{value}'. Valid values: yellow, green, cyan, magenta, blue, red, " +
                $"darkBlue, darkCyan, darkGreen, darkMagenta, darkRed, darkYellow, darkGray, lightGray, black, white, none.");
        return new HighlightColorValues(value);
    }

    /// <summary>
    /// Warn if a value that should be a shading pattern name looks like a hex color instead.
    /// </summary>
    private static void WarnIfShadingOrderWrong(string patternSegment)
    {
        var trimmed = patternSegment.TrimStart('#');
        if (trimmed.Length >= 6 && trimmed.All(char.IsAsciiHexDigit))
            Console.Error.WriteLine($"Warning: '{patternSegment}' looks like a color, but is in the pattern position. "
                + "Shading format: FILL (single value) or PATTERN;FILL[;COLOR] e.g. clear;FF0000");
    }

    /// <summary>
    /// Extract the root path segment (e.g. "/body", "/header[1]", "/footer[2]",
    /// "/styles") from a full parent path. Used by Add helpers that need to
    /// return a path rooted at the actual OOXML part — header/footer parents
    /// must not claim a /body-rooted path since that path won't resolve.
    /// Defaults to "/body" when the input is empty or doesn't start with a
    /// recognized root.
    /// </summary>
    private static string ExtractRootSegment(string? parentPath)
    {
        if (string.IsNullOrEmpty(parentPath)) return "/body";
        var trimmed = parentPath.TrimEnd('/');
        if (trimmed.Length == 0 || trimmed == "/") return "/body";
        // Take the first segment (between leading '/' and the next '/').
        var start = trimmed.StartsWith("/") ? 1 : 0;
        var nextSlash = trimmed.IndexOf('/', start);
        var firstSeg = nextSlash < 0 ? trimmed[start..] : trimmed[start..nextSlash];
        return "/" + firstSeg;
    }

    /// <summary>
    /// Append a child element to parent, but if parent is Body, insert before
    /// the final SectionProperties to maintain valid OOXML structure.
    /// </summary>
    private static void AppendToParent(OpenXmlElement parent, OpenXmlElement child)
    {
        if (parent is Body body)
        {
            var lastSectPr = body.GetFirstChild<SectionProperties>();
            if (lastSectPr != null)
            {
                body.InsertBefore(child, lastSectPr);
                return;
            }
        }
        parent.AppendChild(child);
    }

    /// <summary>
    /// Insert <paramref name="child"/> into <paramref name="parent"/> at the
    /// ChildElements index specified by <paramref name="index"/>. If the
    /// index is null or out of range, falls back to <see cref="AppendToParent"/>
    /// (which respects Body's trailing sectPr).
    /// </summary>
    private static void InsertAtIndexOrAppend(OpenXmlElement parent, OpenXmlElement child, int? index)
    {
        if (index.HasValue && index.Value >= 0 && index.Value < parent.ChildElements.Count)
        {
            parent.InsertBefore(child, parent.ChildElements[index.Value]);
            return;
        }
        AppendToParent(parent, child);
    }

    /// <summary>
    /// Insert <paramref name="newElem"/> into <paramref name="para"/> at the
    /// ChildElements index specified by <paramref name="index"/>, clamping
    /// forward past any leading ParagraphProperties so pPr stays first child.
    /// Null/out-of-range index appends.
    /// </summary>
    private static void InsertIntoParagraph(Paragraph para, OpenXmlElement newElem, int? index)
    {
        var children = para.ChildElements.ToList();
        if (index.HasValue && index.Value >= 0 && index.Value < children.Count)
        {
            var refElem = children[index.Value];
            if (refElem is ParagraphProperties)
            {
                if (index.Value + 1 < children.Count)
                    para.InsertBefore(newElem, children[index.Value + 1]);
                else
                    para.AppendChild(newElem);
            }
            else
            {
                para.InsertBefore(newElem, refElem);
            }
            return;
        }
        para.AppendChild(newElem);
    }

    /// <summary>
    /// Insert multiple elements consecutively into a paragraph, starting at
    /// the ChildElements index (clamped forward past pPr). Later elements go
    /// after earlier ones in order.
    /// </summary>
    private static void InsertIntoParagraph(Paragraph para, IList<OpenXmlElement> newElems, int? index)
    {
        if (newElems == null || newElems.Count == 0) return;
        InsertIntoParagraph(para, newElems[0], index);
        for (int i = 1; i < newElems.Count; i++)
        {
            para.InsertAfter(newElems[i], newElems[i - 1]);
        }
    }

    private static double ParseFontSize(string value) =>
        ParseHelpers.ParseFontSize(value);

    /// <summary>
    /// Get footnote/endnote text, skipping the reference mark run and its trailing space.
    /// </summary>
    private static string GetFootnoteText(OpenXmlElement fnOrEn)
    {
        return string.Join("", fnOrEn.Descendants<Run>()
            .Where(r => r.GetFirstChild<FootnoteReferenceMark>() == null
                     && r.GetFirstChild<EndnoteReferenceMark>() == null)
            .SelectMany(r => r.Elements<Text>())
            .Select(t => t.Text)).TrimStart();
    }

    private static string GetParagraphText(Paragraph para)
    {
        var sb = new StringBuilder();
        foreach (var child in para.ChildElements)
        {
            if (child is Run run)
                sb.Append(string.Concat(run.Elements<Text>().Select(t => t.Text)));
            else if (child is Hyperlink hyperlink)
                sb.Append(string.Concat(hyperlink.Descendants<Text>().Select(t => t.Text)));
        }
        return sb.ToString();
    }

    /// <summary>
    /// Get paragraph text including inline math rendered as readable Unicode.
    /// </summary>
    private static string GetParagraphTextWithMath(Paragraph para)
    {
        var sb = new StringBuilder();
        foreach (var child in para.ChildElements)
        {
            if (child is Run run)
                sb.Append(GetRunText(run));
            else if (child.LocalName == "oMath" || child is M.OfficeMath)
                sb.Append(FormulaParser.ToReadableText(child));
            else if (child is Hyperlink hyperlink)
                sb.Append(string.Concat(hyperlink.Descendants<Text>().Select(t => t.Text)));
        }
        return sb.ToString();
    }

    /// <summary>
    /// Find math elements in a paragraph using both type and localName matching.
    /// </summary>
    private static List<OpenXmlElement> FindMathElements(Paragraph para)
    {
        return para.ChildElements
            .Where(e => e.LocalName == "oMath" || e is M.OfficeMath)
            .ToList();
    }

    /// <summary>
    /// Get all body-level elements, flattening SdtContent containers.
    /// This ensures paragraphs and tables inside w:sdt are not missed.
    /// </summary>
    private static IEnumerable<OpenXmlElement> GetBodyElements(Body body)
    {
        foreach (var element in FlattenWrappers(body.ChildElements))
            yield return element;
    }

    // Descend into SDT (structured document tag) and customXml transparent
    // wrappers so their wrapped paragraphs/tables participate in the body
    // element axis. Without this, docs emitted by e.g. Pages/Google Docs
    // that wrap entire sections in <w:customXml> produce an empty preview.
    private static IEnumerable<OpenXmlElement> FlattenWrappers(IEnumerable<OpenXmlElement> elements)
    {
        foreach (var element in elements)
        {
            if (element is SdtBlock sdt)
            {
                var content = sdt.SdtContentBlock;
                if (content != null)
                    foreach (var child in FlattenWrappers(content.ChildElements))
                        yield return child;
            }
            else if (element.LocalName == "customXml"
                && element.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            {
                foreach (var child in FlattenWrappers(element.ChildElements))
                    yield return child;
            }
            else
            {
                yield return element;
            }
        }
    }

    /// <summary>
    /// Checks if an element is a structural document element worth displaying
    /// (not inline markers like bookmarkStart, bookmarkEnd, proofErr, etc.)
    /// </summary>
    private static bool IsStructuralElement(OpenXmlElement element)
    {
        var name = element.LocalName;
        return name == "sectPr" || name == "altChunk" || name == "customXml";
    }

    /// <summary>
    /// Get all Run elements in a paragraph, including those nested inside
    /// Hyperlink and SdtContent containers.
    /// </summary>
    private static List<Run> GetAllRuns(Paragraph para)
    {
        return para.Descendants<Run>()
            .Where(r => r.GetFirstChild<CommentReference>() == null)
            .ToList();
    }

    /// <summary>
    /// Find the paragraph path where a CommentRangeStart with the given ID is anchored.
    /// Returns "/body/p[N]" or null if not found.
    /// </summary>
    private string? FindCommentAnchorPath(string commentId)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return null;

        var paragraphs = body.Elements<Paragraph>().ToList();
        for (int i = 0; i < paragraphs.Count; i++)
        {
            var hasRange = paragraphs[i].Descendants<CommentRangeStart>()
                .Any(rs => rs.Id?.Value == commentId);
            if (hasRange) return $"/body/{BuildParaPathSegment(paragraphs[i], i + 1)}";
        }
        return null;
    }

    private static string GetRunText(Run run)
    {
        return string.Concat(run.Elements<Text>().Select(t => t.Text));
    }

    // CONSISTENCY(field-cache-stale): walk back from a run carrying an
    // <w:instrText> to the nearest preceding <w:fldChar fldCharType="begin">
    // in the same paragraph and set its dirty="true" attribute, so Word
    // recomputes the field on next open. Used by Set when the instruction
    // text is rewritten — without dirty, the cached result run keeps the
    // old display value (e.g. "PAGE → DATE" still shows the old page
    // number) until the user manually presses F9.
    // CONSISTENCY(field-cache-stale): true when <paramref name="run"/> sits
    // between an owning field's <w:fldChar w:fldCharType="separate"/> and
    // <w:fldChar w:fldCharType="end"/> — i.e. it is the cached result run
    // that Word will overwrite when it recomputes the field. Used by the
    // Set "text=" path to decide whether the caller needs the field marked
    // dirty so their manual edit is preserved on next Word open.
    private static bool IsFieldCachedRun(Run run)
    {
        // Walk backward; the most recent field-char we hit must be a
        // `separate` (with no closing `end` between us and it). Track depth
        // to ignore fully-closed nested fields.
        int closedDepth = 0;
        OpenXmlElement? sibling = run.PreviousSibling();
        while (sibling != null)
        {
            if (sibling is Run sibRun)
            {
                var fld = sibRun.GetFirstChild<FieldChar>();
                if (fld?.FieldCharType?.HasValue == true)
                {
                    var t = fld.FieldCharType.InnerText;
                    if (t == "end")
                        closedDepth++;
                    else if (t == "begin")
                    {
                        if (closedDepth == 0) return false; // begin without separate → not cached
                        closedDepth--;
                    }
                    else if (t == "separate" && closedDepth == 0)
                        return true;
                }
            }
            sibling = sibling.PreviousSibling();
        }
        return false;
    }

    private static void MarkOwningFieldDirty(Run run)
    {
        var para = run.Parent;
        if (para == null) return;
        // Walk siblings backward from this run looking for the OWNING
        // field's <w:fldChar w:fldCharType="begin">. Track depth so that
        // a fully-closed inner field does not get its begin mistaken for
        // the owner of an outer instr. Each `end` we pass while walking
        // means we entered a closed nested field (going backwards), so
        // its `begin` is below us — skip past it. Only the begin at
        // depth 0 is the owner. Use InnerText (not enum equality) since
        // SDK v3 enum equality on FieldCharValues is unreliable (same
        // trap as LineSpacingRuleValues — see WordHandler CLAUDE.md).
        int closedDepth = 0;
        OpenXmlElement? sibling = run.PreviousSibling();
        while (sibling != null)
        {
            if (sibling is Run sibRun)
            {
                var fld = sibRun.GetFirstChild<FieldChar>();
                if (fld?.FieldCharType?.HasValue == true)
                {
                    var t = fld.FieldCharType.InnerText;
                    if (t == "end")
                    {
                        closedDepth++;
                    }
                    else if (t == "begin")
                    {
                        if (closedDepth == 0)
                        {
                            fld.Dirty = true;
                            return;
                        }
                        closedDepth--;
                    }
                }
            }
            sibling = sibling.PreviousSibling();
        }
    }

    // CONSISTENCY(run-special-content): true when <paramref name="key"/>
    // names a typography property that has no glyph to apply on a ptab /
    // fieldChar / instrText / tab / break run. Used by SetElementRun to
    // reject cosmetic writes on these runs, mirroring the readback strip.
    private static bool IsTypographyOnlyKey(string key)
    {
        var k = key.ToLowerInvariant();
        return k is "font" or "font.ascii" or "font.eastasia" or "font.hansi" or "font.cs"
            or "size"
            or "bold" or "italic"
            or "color"
            or "underline" or "underline.color"
            or "strike" or "dstrike" or "highlight"
            or "caps" or "smallcaps" or "vanish"
            or "outline" or "shadow" or "emboss" or "imprint"
            or "noproof" or "rtl"
            or "superscript" or "subscript"
            or "charspacing" or "shading";
    }

    // CONSISTENCY(run-special-content): typography-only Format keys that
    // get scrubbed from runs whose Type was upgraded to ptab / fieldChar /
    // instrText / tab / break. These properties are valid in the underlying
    // <w:rPr> but have no glyph to apply to on these specialized runs, so
    // surfacing them is noise that primes audit tools to misread cosmetic
    // styling on a structural marker as meaningful.
    private static readonly string[] TypographyOnlyKeys =
    {
        "font.ascii", "font.eastAsia", "font.hAnsi", "font.cs",
        "size", "bold", "italic", "color",
        "underline", "underline.color",
        "strike", "dstrike", "highlight",
        "caps", "smallcaps", "vanish",
        "outline", "shadow", "emboss", "imprint",
        "noproof", "rtl", "superscript", "subscript",
        "charSpacing", "shading",
        "effective.size", "effective.size.src",
        "effective.font.ascii", "effective.font.ascii.src",
        "effective.font.eastAsia", "effective.font.eastAsia.src",
        "effective.font.hAnsi", "effective.font.hAnsi.src",
        "effective.font.cs", "effective.font.cs.src",
        "effective.bold", "effective.bold.src",
        "effective.italic", "effective.italic.src",
        "effective.color", "effective.color.src",
        "effective.underline", "effective.underline.src",
    };

    // CONSISTENCY(run-special-content): canonical parsers for the run-internal
    // structural types (ptab / fldChar / break) shared by Add and Set.
    // Lowercase XML attribute values are the canonical input; legacy
    // synonyms (`line`→TextWrapping) are accepted for ergonomics.
    private static EnumValue<AbsolutePositionTabAlignmentValues> ParsePtabAlignment(string s)
    {
        return (s ?? "").Trim().ToLowerInvariant() switch
        {
            "left" => AbsolutePositionTabAlignmentValues.Left,
            "center" => AbsolutePositionTabAlignmentValues.Center,
            "right" => AbsolutePositionTabAlignmentValues.Right,
            _ => throw new ArgumentException(
                $"Invalid ptab alignment '{s}'. Valid: left, center, right.")
        };
    }

    private static EnumValue<AbsolutePositionTabPositioningBaseValues> ParsePtabRelativeTo(string s)
    {
        return (s ?? "").Trim().ToLowerInvariant() switch
        {
            "margin" => AbsolutePositionTabPositioningBaseValues.Margin,
            "indent" => AbsolutePositionTabPositioningBaseValues.Indent,
            _ => throw new ArgumentException(
                $"Invalid ptab relativeTo '{s}'. Valid: margin, indent.")
        };
    }

    private static EnumValue<AbsolutePositionTabLeaderCharValues> ParsePtabLeader(string s)
    {
        return (s ?? "").Trim().ToLowerInvariant() switch
        {
            "none" => AbsolutePositionTabLeaderCharValues.None,
            "dot" => AbsolutePositionTabLeaderCharValues.Dot,
            "hyphen" => AbsolutePositionTabLeaderCharValues.Hyphen,
            "middledot" => AbsolutePositionTabLeaderCharValues.MiddleDot,
            "underscore" => AbsolutePositionTabLeaderCharValues.Underscore,
            _ => throw new ArgumentException(
                $"Invalid ptab leader '{s}'. Valid: none, dot, hyphen, middleDot, underscore.")
        };
    }

    private static EnumValue<FieldCharValues> ParseFieldCharType(string s)
    {
        return (s ?? "").Trim().ToLowerInvariant() switch
        {
            "begin" => FieldCharValues.Begin,
            "separate" => FieldCharValues.Separate,
            "end" => FieldCharValues.End,
            _ => throw new ArgumentException(
                $"Invalid fieldCharType '{s}'. Valid: begin, separate, end.")
        };
    }

    // CONSISTENCY(para-path-canonical): replace the last `/p[...]` segment
    // in <paramref name="path"/> with paraId-form (`/p[@paraId=X]`) when the
    // paragraph carries a w14:paraId. Used by Add helpers whose `parentPath`
    // already targets the paragraph itself (so re-appending /p[N] would
    // double the segment) — the result mirrors what Get later surfaces, so
    // the returned path round-trips through subsequent Get/Set calls
    // without rewriting.
    private static string ReplaceTrailingParaSegment(string path, Paragraph para)
    {
        if (para.ParagraphId?.Value == null) return path;
        var idx = path.LastIndexOf("/p[", StringComparison.Ordinal);
        if (idx < 0) return path;
        var endIdx = path.IndexOf(']', idx);
        if (endIdx < 0) return path;
        return path[..idx] + $"/p[@paraId={para.ParagraphId.Value}]" + path[(endIdx + 1)..];
    }

    private static EnumValue<BreakValues> ParseBreakType(string s)
    {
        return (s ?? "").Trim().ToLowerInvariant() switch
        {
            "page" => BreakValues.Page,
            "column" => BreakValues.Column,
            "textwrapping" or "line" => BreakValues.TextWrapping,
            _ => throw new ArgumentException(
                $"Invalid break type '{s}'. Valid: page, column, line, textwrapping.")
        };
    }

    private string GetStyleName(Paragraph para)
    {
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return "Normal";

        // Try to resolve display name from styles part
        var stylesPart = _doc.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles != null)
        {
            var style = stylesPart.Styles.Elements<Style>()
                .FirstOrDefault(s => s.StyleId?.Value == styleId);
            if (style?.StyleName?.Val?.Value != null)
                return style.StyleName.Val.Value;
        }

        return styleId;
    }

    private static string? GetRunFont(Run run)
    {
        var fonts = run.RunProperties?.RunFonts;
        return fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value ?? fonts?.EastAsia?.Value;
    }

    private static string? GetRunFontSize(Run run)
    {
        var size = run.RunProperties?.FontSize?.Val?.Value;
        if (size == null) return null;
        return $"{int.Parse(size) / 2.0:0.##}pt"; // stored as half-points
    }

    private string GetRunFormatDescription(Run run, Paragraph? para = null)
    {
        var parts = new List<string>();

        RunProperties? rProps;
        if (para != null)
        {
            rProps = ResolveEffectiveRunProperties(run, para);
        }
        else
        {
            rProps = run.RunProperties;
        }
        if (rProps == null) return "(default)";

        var font = GetFontFromProperties(rProps);
        if (font != null) parts.Add(font);

        var size = GetSizeFromProperties(rProps);
        if (size != null) parts.Add(size);

        if (rProps.Bold != null) parts.Add("bold");
        if (rProps.Italic != null) parts.Add("italic");
        if (rProps.Underline != null) parts.Add("underline");
        if (rProps.Strike != null) parts.Add("strikethrough");

        return parts.Count > 0 ? string.Join(" ", parts) : "(default)";
    }

    private static int GetHeadingLevel(string styleName)
    {
        // Heading 1, Heading 2, heading1, 标题 1, etc.
        foreach (var ch in styleName)
        {
            if (char.IsDigit(ch))
                return ch - '0';
        }
        if (styleName == "Title") return 0;
        if (styleName == "Subtitle") return 1;
        return 1;
    }

    private static bool IsNormalStyle(string styleName)
    {
        return styleName is "Normal" or "正文" or "Body Text" or "Body" or "a"
            || styleName.StartsWith("Normal");
    }

    private string? FindWatermark()
    {
        var headerParts = _doc.MainDocumentPart?.HeaderParts;
        if (headerParts == null) return null;

        foreach (var headerPart in headerParts)
        {
            var header = headerPart.Header;
            if (header == null) continue;

            // Search for VML shapes with watermark
            foreach (var pict in header.Descendants<DocumentFormat.OpenXml.Wordprocessing.Picture>())
            {
                foreach (var shape in pict.Descendants<Vml.Shape>())
                {
                    var id = shape.GetAttribute("id", "");
                    if (id.Value?.Contains("WaterMark", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        var textPath = shape.Descendants<Vml.TextPath>().FirstOrDefault();
                        return textPath?.String?.Value ?? "(image watermark)";
                    }
                }
            }

            // Also check for DrawingML watermarks
            foreach (var drawing in header.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>())
            {
                // Simple detection: check if it looks like a watermark by inline/anchor properties
                var docProps = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().FirstOrDefault();
                if (docProps?.Name?.Value?.Contains("WaterMark", StringComparison.OrdinalIgnoreCase) == true)
                {
                    return "(image watermark)";
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Remove all header parts that contain watermark SDT elements.
    /// </summary>
    private void RemoveWatermarkHeaders()
    {
        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return;

        var toRemove = new List<HeaderPart>();
        foreach (var hp in mainPart.HeaderParts)
        {
            if (hp.Header == null) continue;
            // Check for watermark: SDT with docPartGallery="Watermarks" or VML shape with "WaterMark" in id
            var hasSdt = hp.Header.Descendants<SdtProperties>()
                .Any(sp => sp.Descendants<DocPartGallery>().Any(g =>
                    g.Val?.Value?.Equals("Watermarks", StringComparison.OrdinalIgnoreCase) == true));
            if (hasSdt)
            {
                toRemove.Add(hp);
                continue;
            }
            foreach (var pict in hp.Header.Descendants<DocumentFormat.OpenXml.Wordprocessing.Picture>())
            {
                var hasWm = pict.InnerXml.Contains("WaterMark", StringComparison.OrdinalIgnoreCase);
                if (hasWm) { toRemove.Add(hp); break; }
            }
        }

        foreach (var hp in toRemove)
        {
            // Remove header references from section properties
            var relId = mainPart.GetIdOfPart(hp);
            foreach (var sectPr in mainPart.Document?.Body?.Descendants<SectionProperties>() ?? Enumerable.Empty<SectionProperties>())
            {
                var refs = sectPr.Elements<HeaderReference>().Where(r => r.Id?.Value == relId).ToList();
                foreach (var r in refs) r.Remove();
            }
            mainPart.DeletePart(hp);
        }
    }

    private List<string> GetHeaderTexts()
    {
        var results = new List<string>();
        var headerParts = _doc.MainDocumentPart?.HeaderParts;
        if (headerParts == null) return results;

        foreach (var headerPart in headerParts)
        {
            var header = headerPart.Header;
            if (header == null) continue;
            var text = string.Concat(header.Descendants<Text>().Select(t => t.Text)).Trim();
            if (!string.IsNullOrEmpty(text))
                results.Add(text);
        }

        return results;
    }

    private List<string> GetFooterTexts()
    {
        var results = new List<string>();
        var footerParts = _doc.MainDocumentPart?.FooterParts;
        if (footerParts == null) return results;

        foreach (var footerPart in footerParts)
        {
            var footer = footerPart.Footer;
            if (footer == null) continue;

            // Build footer text by processing paragraphs, resolving field codes
            var footerLines = new List<string>();
            foreach (var para in footer.Descendants<Paragraph>())
            {
                var sb = new System.Text.StringBuilder();
                bool inField = false;
                bool pastSeparator = false;

                foreach (var run in para.Elements<Run>())
                {
                    var fldChar = run.GetFirstChild<FieldChar>();
                    if (fldChar != null)
                    {
                        if (fldChar.FieldCharType! == FieldCharValues.Begin)
                        {
                            inField = true;
                            pastSeparator = false;
                        }
                        else if (fldChar.FieldCharType! == FieldCharValues.Separate)
                        {
                            pastSeparator = true;
                        }
                        else if (fldChar.FieldCharType! == FieldCharValues.End)
                        {
                            inField = false;
                            pastSeparator = false;
                        }
                        continue;
                    }

                    var fieldCode = run.GetFirstChild<FieldCode>();
                    if (fieldCode != null)
                    {
                        // Extract field type from instruction (e.g., " PAGE " -> "PAGE")
                        var instr = fieldCode.Text?.Trim() ?? "";
                        var fieldType = instr.Split(' ', System.StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? instr;
                        sb.Append($"[{fieldType.ToUpperInvariant()}]");
                        continue;
                    }

                    // Skip result runs inside a field (they contain stale/literal values)
                    if (inField && pastSeparator)
                        continue;

                    var text = run.GetFirstChild<Text>();
                    if (text != null)
                        sb.Append(text.Text);
                }

                var line = sb.ToString().Trim();
                if (!string.IsNullOrEmpty(line))
                    footerLines.Add(line);
            }

            if (footerLines.Count > 0)
                results.Add(string.Join(" ", footerLines));
        }

        return results;
    }

    private static bool HasMixedPunctuation(string text)
    {
        var chinesePunct = "\uff0c\u3002\uff01\uff1f\u3001\uff1b\uff1a\u201c\u201d\u2018\u2019\uff08\uff09\u3010\u3011";
        bool hasChinese = text.Any(c => chinesePunct.Contains(c));
        bool hasEnglish = text.Any(c => ",.!?;:\"'()[]".Contains(c));
        bool hasChineseChars = text.Any(c => c >= 0x4E00 && c <= 0x9FFF);
        return hasChinese && hasEnglish && hasChineseChars;
    }

    private static RunProperties EnsureRunProperties(Run run)
    {
        return run.RunProperties ?? run.PrependChild(new RunProperties());
    }

    /// <summary>
    /// Parse a w:shd value string ("fill", "val;fill", "val;fill;color") into a Shading element.
    /// Shared by paragraph-level, run-level, and pmrp shading handlers.
    /// </summary>
    private static Shading ParseShadingValue(string value)
    {
        var shdParts = value.Split(';');
        var shd = new Shading();
        if (shdParts.Length == 1)
        {
            shd.Val = ShadingPatternValues.Clear;
            shd.Fill = SanitizeHex(shdParts[0]);
        }
        else if (shdParts.Length >= 2)
        {
            var firstAsHex = shdParts[0].TrimStart('#');
            if (firstAsHex.Length >= 6 && firstAsHex.All(char.IsAsciiHexDigit))
            {
                shd.Val = ShadingPatternValues.Clear;
                shd.Fill = SanitizeHex(shdParts[0]);
            }
            else
            {
                WarnIfShadingOrderWrong(shdParts[0]);
                shd.Val = new ShadingPatternValues(shdParts[0]);
                shd.Fill = SanitizeHex(shdParts[1]);
                if (shdParts.Length >= 3) shd.Color = SanitizeHex(shdParts[2]);
            }
        }
        return shd;
    }

    /// <summary>
    /// Apply a run-level (rPr-style) property to any container that holds rPr children:
    /// <c>RunProperties</c>, <c>ParagraphMarkRunProperties</c>, or <c>StyleRunProperties</c>.
    /// Uses <see cref="OpenXmlCompositeElement"/> + RemoveAllChildren+InsertRunPropInSchemaOrder
    /// so the same logic works across all three despite their different typed property surfaces.
    /// Returns true if the key was handled, false if caller should fall through.
    /// </summary>
    private static bool ApplyRunFormatting(OpenXmlCompositeElement props, string key, string? value)
    {
        if (value is null) return false;
        switch (key.ToLowerInvariant())
        {
            case "size":
            case "font.size":
                var existingFs = props.GetFirstChild<FontSize>();
                if (existingFs != null) existingFs.Val = ((int)Math.Round(ParseFontSize(value) * 2, MidpointRounding.AwayFromZero)).ToString();
                else InsertRunPropInSchemaOrder(props, new FontSize { Val = ((int)Math.Round(ParseFontSize(value) * 2, MidpointRounding.AwayFromZero)).ToString() });
                return true;
            case "font":
            case "font.name":
                var existingRf = props.GetFirstChild<RunFonts>();
                if (existingRf != null) { existingRf.Ascii = value; existingRf.HighAnsi = value; existingRf.EastAsia = value; }
                else InsertRunPropInSchemaOrder(props, new RunFonts { Ascii = value, HighAnsi = value, EastAsia = value });
                return true;
            case "bold":
            case "font.bold":
                props.RemoveAllChildren<Bold>();
                if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new Bold());
                return true;
            case "italic":
            case "font.italic":
                props.RemoveAllChildren<Italic>();
                if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new Italic());
                return true;
            case "color":
            case "font.color":
                props.RemoveAllChildren<Color>();
                // Scheme colors (e.g. accent1, dark2, hyperlink) write to the
                // ThemeColor attribute instead of Val; Val is left at "auto"
                // per ECMA-376 §17.3.2.6 (Excel rejects Val=accent1).
                {
                    var schemeName = OfficeCli.Core.ParseHelpers.NormalizeSchemeColorName(value);
                    Color colorEl;
                    if (schemeName != null)
                    {
                        colorEl = new Color { Val = "auto", ThemeColor = new EnumValue<ThemeColorValues>(new ThemeColorValues(schemeName)) };
                    }
                    else
                    {
                        colorEl = new Color { Val = SanitizeHex(value) };
                    }
                    InsertRunPropInSchemaOrder(props, colorEl);
                }
                return true;
            case "highlight":
                props.RemoveAllChildren<Highlight>();
                InsertRunPropInSchemaOrder(props, new Highlight { Val = ParseHighlightColor(value) });
                return true;
            case "underline":
            case "font.underline":
                props.RemoveAllChildren<Underline>();
                var ulMapped = NormalizeUnderlineValue(value);
                InsertRunPropInSchemaOrder(props, new Underline { Val = new UnderlineValues(ulMapped) });
                return true;
            case "strike" or "strikethrough" or "font.strike" or "font.strikethrough":
                props.RemoveAllChildren<Strike>();
                if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new Strike());
                return true;
            case "dstrike":
                props.RemoveAllChildren<DoubleStrike>();
                if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new DoubleStrike());
                return true;
            case "outline":
                props.RemoveAllChildren<Outline>();
                if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new Outline());
                return true;
            case "shadow":
                props.RemoveAllChildren<Shadow>();
                if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new Shadow());
                return true;
            case "emboss":
                props.RemoveAllChildren<Emboss>();
                if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new Emboss());
                return true;
            case "imprint":
                props.RemoveAllChildren<Imprint>();
                if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new Imprint());
                return true;
            case "noproof":
                props.RemoveAllChildren<NoProof>();
                if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new NoProof());
                return true;
            case "rtl":
                props.RemoveAllChildren<RightToLeftText>();
                if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new RightToLeftText());
                return true;
            case "charspacing" or "letterspacing" or "spacing":
                var csPt = value.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
                    ? ParseHelpers.SafeParseDouble(value[..^2], "charspacing")
                    : ParseHelpers.SafeParseDouble(value, "charspacing");
                props.RemoveAllChildren<Spacing>();
                InsertRunPropInSchemaOrder(props, new Spacing { Val = (int)Math.Round(csPt * 20, MidpointRounding.AwayFromZero) });
                return true;
            case "shading" or "shd":
                props.RemoveAllChildren<Shading>();
                InsertRunPropInSchemaOrder(props, ParseShadingValue(value));
                return true;
            case "superscript":
                props.RemoveAllChildren<VerticalTextAlignment>();
                if (IsTruthy(value))
                    InsertRunPropInSchemaOrder(props, new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
                return true;
            case "subscript":
                props.RemoveAllChildren<VerticalTextAlignment>();
                if (IsTruthy(value))
                    InsertRunPropInSchemaOrder(props, new VerticalTextAlignment { Val = VerticalPositionValues.Subscript });
                return true;
            case "caps":
            case "allcaps":
                props.RemoveAllChildren<Caps>();
                if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new Caps());
                return true;
            case "smallcaps":
                props.RemoveAllChildren<SmallCaps>();
                if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new SmallCaps());
                return true;
            case "vanish":
                props.RemoveAllChildren<Vanish>();
                if (IsTruthy(value)) InsertRunPropInSchemaOrder(props, new Vanish());
                return true;
            default:
                return false;
        }
    }

    /// <summary>
    /// Insert a run property element in the correct CT_RPr schema position.
    /// CT_RPr order: rFonts, b, bCs, i, iCs, caps, smallCaps, strike, dstrike, outline, shadow,
    /// emboss, imprint, noProof, snapToGrid, vanish, webHidden, color, spacing, w, kern, position,
    /// sz, szCs, highlight, u, effect, ...
    /// </summary>
    private static void InsertRunPropInSchemaOrder(OpenXmlCompositeElement props, OpenXmlElement elem)
    {
        // Map element types to their position in the CT_RPr schema sequence.
        // Only the types we actually use are listed; unlisted types get a high index (appended at end).
        static int SchemaIndex(OpenXmlElement e) => e switch
        {
            RunFonts => 0,
            Bold => 1,
            BoldComplexScript => 2,
            Italic => 3,
            ItalicComplexScript => 4,
            Caps => 5,
            SmallCaps => 6,
            Strike => 7,
            // dstrike, outline, shadow, emboss, imprint, noProof, snapToGrid
            Vanish => 14,
            // webHidden = 15
            Color => 16,
            Spacing => 17,
            // w = 18, kern = 19, position = 20
            FontSize => 21,
            FontSizeComplexScript => 22,
            Highlight => 23,
            Underline => 24,
            // effect = 25, bdr = 26
            Shading => 27,
            // fitText = 28
            VerticalTextAlignment => 29,
            // rtl, cs, em, lang, ...
            _ => 100,
        };

        int targetIdx = SchemaIndex(elem);

        // Find the first existing child whose schema position is after the element we're inserting
        foreach (var child in props.ChildElements)
        {
            if (SchemaIndex(child) > targetIdx)
            {
                child.InsertBeforeSelf(elem);
                return;
            }
        }
        // No later element found — append at end
        props.AppendChild(elem);
    }

    private static string GetBookmarkText(BookmarkStart bkStart)
    {
        var bkId = bkStart.Id?.Value;
        if (bkId == null) return "";

        var sb = new System.Text.StringBuilder();
        var sibling = bkStart.NextSibling();
        while (sibling != null)
        {
            if (sibling is BookmarkEnd bkEnd && bkEnd.Id?.Value == bkId)
                break;
            if (sibling is Run run)
                sb.Append(string.Concat(run.Descendants<Text>().Select(t => t.Text)));
            sibling = sibling.NextSibling();
        }
        return sb.ToString();
    }

    // ==================== Find / Format / Replace ====================

    /// <summary>
    /// Build a flat list of (Run, Text, charStart, charEnd) spans for a paragraph.
    /// Uses Descendants to include runs inside hyperlinks, w:ins, w:del, etc.
    /// Shared by ProcessFindInParagraph, SplitRunsAtRange, etc.
    /// </summary>
    private static List<(Run Run, Text TextElement, int Start, int End)> BuildRunTexts(Paragraph para)
    {
        var runTexts = new List<(Run Run, Text TextElement, int Start, int End)>();
        int pos = 0;
        foreach (var run in para.Descendants<Run>())
        {
            foreach (var text in run.Elements<Text>())
            {
                var len = text.Text?.Length ?? 0;
                if (len > 0)
                    runTexts.Add((run, text, pos, pos + len));
                pos += len;
            }
        }
        return runTexts;
    }

    /// <summary>
    /// Parse a find pattern: plain text or regex (r"..." prefix).
    /// Returns (pattern, isRegex).
    /// </summary>
    private static (string Pattern, bool IsRegex) ParseFindPattern(string value)
    {
        // r"..." or r'...' → regex
        if (value.Length >= 3 && value[0] == 'r' && (value[1] == '"' || value[1] == '\''))
        {
            var quote = value[1];
            var endIdx = value.LastIndexOf(quote);
            if (endIdx > 1)
                return (value[2..endIdx], true);
        }
        return (value, false);
    }

    /// <summary>
    /// Find all match ranges in fullText using either plain text or regex.
    /// Returns list of (start, length) pairs, sorted by start ascending.
    /// </summary>
    private static List<(int Start, int Length)> FindMatchRanges(string fullText, string pattern, bool isRegex)
    {
        var ranges = new List<(int Start, int Length)>();
        if (isRegex)
        {
            try
            {
                // BUG-TESTER fuzz-2: bound matching with a hard timeout so
                // catastrophic-backtracking patterns (e.g. "(a+)+b") fail fast
                // instead of hanging the CLI / process.
                foreach (System.Text.RegularExpressions.Match m in
                    System.Text.RegularExpressions.Regex.Matches(
                        fullText,
                        pattern,
                        System.Text.RegularExpressions.RegexOptions.None,
                        FindRegexMatchTimeout))
                {
                    if (m.Length > 0) // skip zero-length matches
                        ranges.Add((m.Index, m.Length));
                }
            }
            catch (System.Text.RegularExpressions.RegexParseException ex)
            {
                throw new ArgumentException($"Invalid regex pattern '{pattern}': {ex.Message}", ex);
            }
            catch (System.Text.RegularExpressions.RegexMatchTimeoutException ex)
            {
                throw new ArgumentException(
                    $"Regex pattern '{pattern}' exceeded {FindRegexMatchTimeout.TotalSeconds}s match timeout (catastrophic backtracking?)",
                    ex);
            }
        }
        else
        {
            int idx = 0;
            while ((idx = fullText.IndexOf(pattern, idx, StringComparison.Ordinal)) >= 0)
            {
                ranges.Add((idx, pattern.Length));
                idx += pattern.Length;
            }
        }
        return ranges;
    }

    /// <summary>
    /// Split a run at a character offset within its text content.
    /// Returns the new right-side run (inserted after the original).
    /// The original run keeps text [0..charOffset), new run gets [charOffset..).
    /// RunProperties are deep-cloned. rsidR is cleared on the new run.
    /// </summary>
    /// <summary>
    /// Split a paragraph at the given character offset, producing a head
    /// paragraph (the original <paramref name="para"/>, now holding
    /// runs/content up to <paramref name="charOffset"/>) followed by a tail
    /// paragraph inserted as its immediate next sibling (holding content
    /// from <paramref name="charOffset"/> onward). The tail inherits a
    /// clone of the head's paragraph properties so style/numbering/heading
    /// is preserved on both halves — matching Word's own Enter-key split.
    /// Preconditions: 0 &lt; charOffset &lt; fullText length (boundary cases
    /// should be handled by the caller without splitting).
    /// </summary>
    private static Paragraph SplitParagraphAtOffset(Paragraph para, int charOffset)
    {
        var runTexts = BuildRunTexts(para);

        // Split the run that straddles charOffset so a clean run boundary
        // exists at the split point. After this call, runTexts is stale.
        foreach (var rt in runTexts)
        {
            if (charOffset > rt.Start && charOffset < rt.End)
            {
                var localOffset = charOffset - rt.Start;
                SplitRunAtOffset(rt.Run, localOffset);
                break;
            }
        }

        // Recompute run positions and partition runs into head (< charOffset)
        // and tail (>= charOffset). Inline children other than Run
        // (hyperlink/bookmark/field/sdt/…) are routed by their document
        // order relative to the cumulative text length: anything whose
        // text footprint falls entirely on the tail side moves with the
        // tail paragraph. Runs with zero-length text at the boundary stay
        // with the head (matches Enter-key behavior in Word).
        var tail = new Paragraph();
        if (para.ParagraphProperties != null)
            tail.PrependChild((ParagraphProperties)para.ParagraphProperties.CloneNode(true));

        // Walk children in document order. For Run, compute its text range
        // and decide; for non-Run inline children, treat their text contribution
        // as zero-length at the current cumulative offset (consistent with how
        // BuildRunTexts ignores them).
        int cumulative = 0;
        var toMove = new List<OpenXmlElement>();
        foreach (var child in para.ChildElements.ToList())
        {
            if (child is ParagraphProperties) continue;
            if (child is Run run)
            {
                var runLen = run.Elements<Text>().Sum(t => t.Text?.Length ?? 0);
                if (cumulative >= charOffset)
                {
                    toMove.Add(child);
                }
                cumulative += runLen;
            }
            else
            {
                // Non-run inline content: keep on head side if we're still
                // before the split point, move to tail if we've crossed it.
                if (cumulative >= charOffset)
                    toMove.Add(child);
            }
        }

        foreach (var el in toMove)
        {
            el.Remove();
            tail.AppendChild(el);
        }

        para.InsertAfterSelf(tail);
        return tail;
    }

    private static Run SplitRunAtOffset(Run run, int charOffset)
    {
        // Find the Text element containing the split point
        int pos = 0;
        foreach (var text in run.Elements<Text>().ToList())
        {
            var len = text.Text?.Length ?? 0;
            if (pos + len > charOffset && charOffset > pos)
            {
                var localOffset = charOffset - pos;
                var leftText = text.Text![..localOffset];
                var rightText = text.Text![localOffset..];

                // Clone the run for the right side
                var rightRun = (Run)run.CloneNode(true);
                // Clear rsidR on cloned run
                rightRun.RsidRunProperties = null;
                rightRun.RsidRunAddition = null;

                // Set left run text
                text.Text = leftText;
                text.Space = SpaceProcessingModeValues.Preserve;

                // Set right run text — find corresponding Text in clone
                var rightTexts = rightRun.Elements<Text>().ToList();
                // The cloned run has same structure; find the matching Text node
                int textIdx = run.Elements<Text>().ToList().IndexOf(text);
                if (textIdx >= 0 && textIdx < rightTexts.Count)
                {
                    rightTexts[textIdx].Text = rightText;
                    rightTexts[textIdx].Space = SpaceProcessingModeValues.Preserve;
                    // Remove any Text elements before the split Text in right run
                    for (int i = 0; i < textIdx; i++)
                        rightTexts[i].Text = "";
                }

                // Insert right run after original
                run.InsertAfterSelf(rightRun);
                return rightRun;
            }
            pos += len;
        }
        // charOffset is at boundary — shouldn't normally be called, return run itself
        return run;
    }

    /// <summary>
    /// Split runs in a paragraph so that the character range [charStart, charEnd)
    /// is covered by dedicated runs. Returns the list of runs covering that range.
    /// </summary>
    private static List<Run> SplitRunsAtRange(Paragraph para, int charStart, int charEnd)
    {
        // Split at charEnd first (so charStart offsets remain valid)
        var runTexts = BuildRunTexts(para);
        foreach (var rt in runTexts)
        {
            if (charEnd > rt.Start && charEnd < rt.End)
            {
                var localOffset = charEnd - rt.Start;
                SplitRunAtOffset(rt.Run, localOffset);
                break;
            }
        }

        // Rebuild after split, then split at charStart
        runTexts = BuildRunTexts(para);
        foreach (var rt in runTexts)
        {
            if (charStart > rt.Start && charStart < rt.End)
            {
                var localOffset = charStart - rt.Start;
                SplitRunAtOffset(rt.Run, localOffset);
                break;
            }
        }

        // Rebuild and collect runs covering [charStart, charEnd)
        runTexts = BuildRunTexts(para);
        var result = new List<Run>();
        foreach (var rt in runTexts)
        {
            if (rt.Start >= charStart && rt.End <= charEnd)
                result.Add(rt.Run);
        }
        return result;
    }

    /// <summary>
    /// Unified find operation on a paragraph: replace text and/or apply formatting.
    /// Returns the number of matches processed.
    /// </summary>
    private static int ProcessFindInParagraph(
        Paragraph para,
        string pattern,
        bool isRegex,
        string? replace,
        Dictionary<string, string>? formatProps)
    {
        var runTexts = BuildRunTexts(para);
        if (runTexts.Count == 0) return 0;

        var fullText = string.Concat(runTexts.Select(rt => rt.TextElement.Text));
        // CONSISTENCY(regex-backref-expand): collect Match objects in regex mode so we can
        // call Match.Result(replace) — which expands backreferences against the original
        // match captures, and unlike re-running Regex.Replace on the substring, correctly
        // handles lookaround anchors (e.g. r"foo(?=bar)") whose context is lost in isolation.
        // BUG-TESTER+FUZZER R31: wrap with try/catch so RegexMatchTimeoutException is
        // converted to ArgumentException (consistent with FindMatchRanges), and avoid
        // a second Regex.Matches call by deriving ranges from the same Match list.
        List<System.Text.RegularExpressions.Match>? matchObjs = null;
        List<(int Start, int Length)> matches;
        if (isRegex)
        {
            try
            {
                matchObjs = System.Text.RegularExpressions.Regex.Matches(
                        fullText,
                        pattern,
                        System.Text.RegularExpressions.RegexOptions.None,
                        FindRegexMatchTimeout)
                    .Cast<System.Text.RegularExpressions.Match>()
                    .Where(m => m.Length > 0)
                    .ToList();
            }
            catch (System.Text.RegularExpressions.RegexParseException ex)
            {
                throw new ArgumentException($"Invalid regex pattern '{pattern}': {ex.Message}", ex);
            }
            catch (System.Text.RegularExpressions.RegexMatchTimeoutException ex)
            {
                throw new ArgumentException(
                    $"Regex pattern '{pattern}' exceeded {FindRegexMatchTimeout.TotalSeconds}s match timeout (catastrophic backtracking?)",
                    ex);
            }
            matches = matchObjs.Select(m => (m.Index, m.Length)).ToList();
        }
        else
        {
            matches = FindMatchRanges(fullText, pattern, isRegex);
        }
        if (matches.Count == 0) return 0;

        // Process from end to start to preserve character offsets
        for (int i = matches.Count - 1; i >= 0; i--)
        {
            var (matchStart, matchLen) = matches[i];
            var matchEnd = matchStart + matchLen;

            if (replace != null)
            {
                // For regex replace, expand backreferences ($1, ${name}, etc.) via
                // Match.Result so lookaround context is preserved.
                string effectiveReplace = replace;
                if (isRegex && matchObjs != null && i < matchObjs.Count)
                {
                    effectiveReplace = matchObjs[i].Result(replace);
                }

                // BUG-BT-2: detect cross-hyperlink-boundary replacement. If the
                // match spans runs whose Hyperlink ancestors differ (e.g. one
                // run inside a Hyperlink, another in plain paragraph body),
                // a naive cross-run text edit destroys the hyperlink structure
                // (URL + blue/underline formatting are lost). Reject up-front
                // with a clear error rather than silently corrupting the doc.
                {
                    var affected = BuildRunTexts(para)
                        .Where(rt => rt.End > matchStart && rt.Start < matchEnd)
                        .Select(rt => rt.Run.Ancestors<Hyperlink>().FirstOrDefault())
                        .Distinct()
                        .ToList();
                    if (affected.Count > 1)
                    {
                        throw new ArgumentException(
                            $"find/replace cannot span a hyperlink boundary (match at offset {matchStart}, length {matchLen}): " +
                            $"the match crosses into or out of a <w:hyperlink>, which would destroy its URL and formatting. " +
                            $"Narrow the pattern to stay inside or outside the hyperlink, or edit the hyperlink text directly.");
                    }
                }

                // Step 1: Replace text in affected runs (same logic as old ReplaceInParagraph)
                var currentRunTexts = BuildRunTexts(para);
                bool first = true;
                foreach (var rt in currentRunTexts)
                {
                    if (rt.End <= matchStart || rt.Start >= matchEnd)
                        continue;

                    var textStr = rt.TextElement.Text ?? "";
                    var localStart = Math.Max(0, matchStart - rt.Start);
                    var localEnd = Math.Min(textStr.Length, matchEnd - rt.Start);

                    if (first)
                    {
                        rt.TextElement.Text = textStr[..localStart] + effectiveReplace + textStr[localEnd..];
                        rt.TextElement.Space = SpaceProcessingModeValues.Preserve;
                        first = false;
                    }
                    else
                    {
                        rt.TextElement.Text = textStr[..Math.Max(0, matchStart - rt.Start)] + textStr[localEnd..];
                        rt.TextElement.Space = SpaceProcessingModeValues.Preserve;
                    }
                }

                // BUG-TESTER fuzz-1: cross-run replace consumes intermediate runs leaving
                // them with empty <w:t/> — drop those orphan runs so persisted XML stays clean.
                // Only remove runs whose Text element is now empty AND have no other
                // semantic children (Break, TabChar, Drawing, FieldChar, Picture, etc.).
                // RunProperties (rPr) alone is not semantic content.
                var emptyRunsToRemove = new List<Run>();
                foreach (var run in para.Descendants<Run>())
                {
                    bool hasContent = false;
                    bool hasEmptyText = false;
                    foreach (var child in run.ChildElements)
                    {
                        if (child is RunProperties)
                            continue;
                        if (child is Text t)
                        {
                            if (string.IsNullOrEmpty(t.Text))
                                hasEmptyText = true;
                            else
                                hasContent = true;
                        }
                        else
                        {
                            hasContent = true;
                        }
                    }
                    if (hasEmptyText && !hasContent)
                        emptyRunsToRemove.Add(run);
                }
                foreach (var run in emptyRunsToRemove)
                    run.Remove();

                // Step 2: If format props, split at the replaced text position and apply
                if (formatProps != null && formatProps.Count > 0)
                {
                    // The replaced text now starts at matchStart with length = effectiveReplace.Length
                    var replacedEnd = matchStart + effectiveReplace.Length;
                    if (effectiveReplace.Length > 0)
                    {
                        var targetRuns = SplitRunsAtRange(para, matchStart, replacedEnd);
                        foreach (var run in targetRuns)
                        {
                            var rPr = EnsureRunProperties(run);
                            foreach (var (key, value) in formatProps)
                                ApplyRunFormatting(rPr, key, value);
                        }
                    }
                }
            }
            else if (formatProps != null && formatProps.Count > 0)
            {
                // No replace, just split and format
                var targetRuns = SplitRunsAtRange(para, matchStart, matchEnd);
                foreach (var run in targetRuns)
                {
                    var rPr = EnsureRunProperties(run);
                    foreach (var (key, value) in formatProps)
                        ApplyRunFormatting(rPr, key, value);
                }
            }
        }

        return matches.Count;
    }

    /// <summary>
    /// Unified find operation: process find/replace/format across paragraphs resolved from a path.
    /// Called from Set when 'find' key is present.
    /// Returns (matchCount, unsupportedKeys).
    /// </summary>
    private int ProcessFind(
        string path,
        string findValue,
        string? replace,
        Dictionary<string, string> formatProps)
    {
        var (pattern, isRegex) = ParseFindPattern(findValue);
        if (string.IsNullOrEmpty(pattern) && !isRegex) return 0;

        // Resolve paragraphs from path
        var paragraphs = ResolveParagraphsForFind(path);

        int totalCount = 0;
        foreach (var para in paragraphs)
        {
            var count = ProcessFindInParagraph(para, pattern, isRegex, replace, formatProps.Count > 0 ? formatProps : null);
            if (count > 0)
                para.TextId = GenerateParaId();
            totalCount += count;
        }

        return totalCount;
    }

    /// <summary>
    /// Resolve paragraphs for a find operation based on path.
    /// "/" or "/body" → body paragraphs; "/header[N]" → header N; "/footer[N]" → footer N;
    /// "/paragraph[N]" → specific paragraph; selector → query results.
    ///
    /// BUG-TESTER+FUZZER R33: out-of-bound indices and unrecognized Word
    /// roots (e.g. /slide[1]) must throw ArgumentException instead of
    /// silently returning an empty paragraph list. Mirrors the PPTX
    /// ResolvePptParagraphsForFind contract — see commit 898f9284.
    /// CONSISTENCY(find-strict-path): Word + PPTX share this strict-path
    /// behaviour; if the contract is relaxed, update both sites in one pass.
    /// </summary>
    private List<Paragraph> ResolveParagraphsForFind(string path)
    {
        var paragraphs = new List<Paragraph>();
        var mainPart = _doc.MainDocumentPart;

        if (path is "/" or "" or "/body")
        {
            if (mainPart?.Document?.Body != null)
                paragraphs.AddRange(mainPart.Document.Body.Descendants<Paragraph>());
            return paragraphs;
        }

        if (path.StartsWith("/header[", StringComparison.OrdinalIgnoreCase))
        {
            var idx = ParseHelpers.SafeParseInt(path.Split('[', ']')[1], "header index") - 1;
            var headers = mainPart?.HeaderParts.ToList() ?? new List<HeaderPart>();
            if (idx < 0 || idx >= headers.Count)
                throw new ArgumentException($"Header index out of range: {idx + 1} (have {headers.Count} header(s)).");
            var headerPart = headers[idx];
            if (headerPart.Header != null)
                paragraphs.AddRange(headerPart.Header.Descendants<Paragraph>());
            return paragraphs;
        }

        if (path.StartsWith("/footer[", StringComparison.OrdinalIgnoreCase))
        {
            var idx = ParseHelpers.SafeParseInt(path.Split('[', ']')[1], "footer index") - 1;
            var footers = mainPart?.FooterParts.ToList() ?? new List<FooterPart>();
            if (idx < 0 || idx >= footers.Count)
                throw new ArgumentException($"Footer index out of range: {idx + 1} (have {footers.Count} footer(s)).");
            var footerPart = footers[idx];
            if (footerPart.Footer != null)
                paragraphs.AddRange(footerPart.Footer.Descendants<Paragraph>());
            return paragraphs;
        }

        if (path.StartsWith("/"))
        {
            // Specific element path — navigate to it. NavigateToElement returns
            // null for both unknown roots (e.g. /slide[1]) and out-of-bound
            // indices (e.g. /body/p[999]); both must throw, never silently
            // resolve to zero paragraphs.
            var element = NavigateToElement(ParsePath(path));
            if (element == null)
                throw new ArgumentException(
                    $"Cannot resolve find scope path: '{path}'. "
                    + "Expected /, /body, /body/p[N], /body/p[N]/r[K], /body/tbl[N], "
                    + "/body/tbl[N]/tr[R]/tc[C], /header[N], or /footer[N]; "
                    + "or a CSS-style selector (e.g. paragraph, run).");

            if (element is Paragraph p)
            {
                paragraphs.Add(p);
                return paragraphs;
            }

            // BUG-BT-1: when path resolves to an inline element (e.g. a Run
            // under /body/p[N]/r[K], or a Hyperlink), Descendants<Paragraph>()
            // is empty — the find would silently match nothing. Walk up to
            // the containing paragraph instead so /run paths still work,
            // and also harvest any paragraphs nested inside (e.g. tables).
            var nestedParas = element.Descendants<Paragraph>().ToList();
            if (nestedParas.Count > 0)
            {
                paragraphs.AddRange(nestedParas);
            }
            else
            {
                var ancestorPara = element.Ancestors<Paragraph>().FirstOrDefault();
                if (ancestorPara != null)
                    paragraphs.Add(ancestorPara);
            }
            return paragraphs;
        }

        // Selector — query and resolve each result's paragraphs
        var targets = Query(path);
        foreach (var target in targets)
        {
            var elem = NavigateToElement(ParsePath(target.Path));
            if (elem is Paragraph tp)
                paragraphs.Add(tp);
            else if (elem != null)
                paragraphs.AddRange(elem.Descendants<Paragraph>());
        }

        return paragraphs;
    }

    // ==================== Add at find position ====================

    private static readonly HashSet<string> InlineTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "run", "r", "picture", "image", "img", "hyperlink", "link",
        "field", "pagenum", "pagenumber", "page", "numpages", "sectionpages", "section",
        "date", "createdate", "savedate", "printdate", "edittime", "time",
        "author", "lastsavedby", "title", "subject", "filename",
        "numwords", "numchars", "revnum", "template", "comments", "doccomments", "keywords",
        "mergefield", "ref", "pageref", "noteref", "seq", "styleref", "docproperty", "if",
        "pagebreak", "columnbreak", "break", "footnote", "endnote",
        "equation", "formula", "math", "bookmark", "formfield",
        "comment", "sdt", "contentcontrol", "chart"
    };

    /// <summary>
    /// Add an element at a text-find position within a paragraph.
    /// For inline types: split the run at the find position and insert inline.
    /// For block types: split the paragraph at the find position and insert the block element between.
    /// </summary>
    private string AddAtFindPosition(
        OpenXmlElement parent,
        string parentPath,
        string type,
        string findValue,
        bool isAfter, // true = after-find, false = before-find
        InsertPosition? position,
        Dictionary<string, string> properties)
    {
        // Support regex=true prop as alternative to r"..." prefix
        // CONSISTENCY(find-regex): mirror of WordHandler.Set.cs:60-61. grep
        // "CONSISTENCY(find-regex)" for every project-wide call site.
        if (properties.TryGetValue("regex", out var regexFlag) && ParseHelpers.IsTruthySafe(regexFlag) && !findValue.StartsWith("r\"") && !findValue.StartsWith("r'"))
            findValue = $"r\"{findValue}\"";

        var (pattern, isRegex) = ParseFindPattern(findValue);

        // Guard: empty find pattern would produce unbounded matches and blow
        // up downstream regex/plain-text scans. Surface a clean error instead
        // of leaking the raw .NET exception.
        if (string.IsNullOrEmpty(pattern))
            throw new ArgumentException("find: pattern must not be empty. Example: --after \"find:hello\".");

        // Resolve to a paragraph — either the parent itself, or the first
        // descendant paragraph of a container (body/cell/sdt) whose text
        // matches the pattern.
        Paragraph para;
        string paraPath;
        if (parent is Paragraph p)
        {
            para = p;
            paraPath = parentPath;
        }
        else
        {
            var hit = FindParagraphContainingText(parent, parentPath, pattern, isRegex)
                ?? throw new ArgumentException(
                    $"Text '{findValue}' not found in any paragraph under {parentPath}.");
            para = hit.Para;
            paraPath = hit.Path;
        }

        var runTexts = BuildRunTexts(para);
        if (runTexts.Count == 0)
            throw new ArgumentException("Paragraph has no text content to search.");

        var fullText = string.Concat(runTexts.Select(rt => rt.TextElement.Text));
        var matches = FindMatchRanges(fullText, pattern, isRegex);
        if (matches.Count == 0)
            throw new ArgumentException($"Text '{findValue}' not found in paragraph.");

        // Use first match
        var (matchStart, matchLen) = matches[0];
        var splitPoint = isAfter ? matchStart + matchLen : matchStart;

        bool isInline = InlineTypes.Contains(type);

        if (isInline)
        {
            return AddInlineAtSplitPoint(para, paraPath, splitPoint, type, position, properties);
        }
        else
        {
            // Block types (paragraph/table/section/toc/…) under a `find:`
            // anchor: honor the literal position. When the anchor lands at
            // a paragraph boundary (splitPoint == 0 or == full length),
            // insert as a sibling before/after the matched paragraph
            // (no split needed). When the anchor lands mid-paragraph,
            // split the paragraph at that offset and insert the new block
            // between the two halves as body-level siblings.
            //
            // This mirrors Word's native "cursor mid-sentence → Insert →
            // Table" behavior: the user asked for position X, they get
            // the block at position X, even if that requires splitting
            // the containing paragraph.
            var container = para.Parent
                ?? throw new InvalidOperationException("Matched paragraph has no parent container.");
            var containerPath = paraPath.Contains('/')
                ? paraPath[..paraPath.LastIndexOf('/')]
                : "/body";
            var siblings = container.Elements<OpenXmlElement>().ToList();
            var paraIdx = siblings.IndexOf(para);
            if (paraIdx < 0)
                throw new InvalidOperationException("Matched paragraph not found among its parent's children.");

            var totalLen = fullText.Length;
            bool atBoundary = splitPoint == 0 || splitPoint == totalLen;

            if (atBoundary)
            {
                var insertIdx = (splitPoint == totalLen) ? paraIdx + 1 : paraIdx;
                return Add(containerPath, type, InsertPosition.AtIndex(insertIdx), properties);
            }

            // Mid-paragraph: split the paragraph, inherit pPr on the tail,
            // then insert the new block between the head and tail paragraphs.
            SplitParagraphAtOffset(para, splitPoint);
            // Head paragraph is now `para`; tail paragraph is its immediate
            // following sibling. Insert the new block between them.
            var insertIdxMid = paraIdx + 1;
            return Add(containerPath, type, InsertPosition.AtIndex(insertIdxMid), properties);
        }
    }

    /// <summary>
    /// Walk the child paragraphs of a container and return the first paragraph
    /// (plus its constructed path) whose text matches the given pattern.
    /// Used to let body-level find: anchors resolve without requiring the
    /// caller to spell out a specific paragraph path.
    /// </summary>
    private (Paragraph Para, string Path)? FindParagraphContainingText(
        OpenXmlElement container, string containerPath, string pattern, bool isRegex)
    {
        var paragraphs = container.Elements<Paragraph>().ToList();
        for (int i = 0; i < paragraphs.Count; i++)
        {
            var candidate = paragraphs[i];
            var runTexts = BuildRunTexts(candidate);
            if (runTexts.Count == 0) continue;

            var fullText = string.Concat(runTexts.Select(rt => rt.TextElement.Text));
            if (FindMatchRanges(fullText, pattern, isRegex).Count > 0)
            {
                var paraPath = $"{containerPath}/{BuildParaPathSegment(candidate, i + 1)}";
                return (candidate, paraPath);
            }
        }
        return null;
    }

    /// <summary>
    /// Insert an inline element at a character split point within a paragraph.
    /// Splits the run at the position and inserts the element.
    /// </summary>
    private string AddInlineAtSplitPoint(
        Paragraph para,
        string parentPath,
        int splitPoint,
        string type,
        InsertPosition? position,
        Dictionary<string, string> properties)
    {
        // Split runs at the point
        var runTexts = BuildRunTexts(para);
        Run? insertAfterRun = null;

        foreach (var rt in runTexts)
        {
            if (splitPoint >= rt.Start && splitPoint <= rt.End)
            {
                if (splitPoint == rt.Start)
                {
                    // Insert before this run — find previous run
                    insertAfterRun = rt.Run.PreviousSibling<Run>();
                }
                else if (splitPoint == rt.End)
                {
                    // Insert after this run
                    insertAfterRun = rt.Run;
                }
                else
                {
                    // Split the run at the offset
                    var localOffset = splitPoint - rt.Start;
                    SplitRunAtOffset(rt.Run, localOffset);
                    insertAfterRun = rt.Run; // insert after the left portion
                }
                break;
            }
        }

        // Calculate run-based index for insertion
        var runs = para.Elements<Run>().ToList();
        int runIndex;
        if (insertAfterRun != null)
        {
            var idx = runs.IndexOf(insertAfterRun);
            runIndex = idx >= 0 ? idx + 1 : runs.Count;
        }
        else
        {
            runIndex = 0; // insert before all runs
        }

        // Convert run-count index → ChildElements-index so downstream handlers
        // (which read parent.ChildElements[index]) land at the right slot. When
        // the paragraph has a ParagraphProperties child, the ChildElements
        // index is shifted by one; when inserting before all runs, point at
        // the first run's ChildElements index rather than 0 (which is pPr).
        var childElems = para.ChildElements.ToList();
        int childIndex;
        if (runIndex >= runs.Count)
        {
            childIndex = childElems.Count;
        }
        else
        {
            var targetRun = runs[runIndex];
            childIndex = childElems.IndexOf(targetRun);
            if (childIndex < 0) childIndex = childElems.Count;
        }

        return Add(parentPath, type, InsertPosition.AtIndex(childIndex), properties);
    }

    /// <summary>
    /// Insert a block element at a character split point within a paragraph.
    /// Splits the paragraph into two and inserts the block element between them.
    /// </summary>
    private string AddBlockAtSplitPoint(
        Paragraph para,
        string parentPath,
        int splitPoint,
        string type,
        InsertPosition? position,
        Dictionary<string, string> properties)
    {
        var runTexts = BuildRunTexts(para);
        var fullText = string.Concat(runTexts.Select(rt => rt.TextElement.Text));

        // If split point is at the very end, just insert after the paragraph
        if (splitPoint >= fullText.Length)
        {
            var bodyPath = parentPath.Contains('/') ? parentPath[..parentPath.LastIndexOf('/')] : "/body";
            return Add(bodyPath, type, InsertPosition.AfterElement(parentPath.Split('/').Last()), properties);
        }

        // If split point is at the very beginning, just insert before the paragraph
        if (splitPoint <= 0)
        {
            var bodyPath = parentPath.Contains('/') ? parentPath[..parentPath.LastIndexOf('/')] : "/body";
            return Add(bodyPath, type, InsertPosition.BeforeElement(parentPath.Split('/').Last()), properties);
        }

        // Split runs at the point
        foreach (var rt in runTexts)
        {
            if (splitPoint > rt.Start && splitPoint < rt.End)
            {
                var localOffset = splitPoint - rt.Start;
                SplitRunAtOffset(rt.Run, localOffset);
                break;
            }
        }

        // Rebuild run list after split
        runTexts = BuildRunTexts(para);
        fullText = string.Concat(runTexts.Select(rt => rt.TextElement.Text));

        // Find the first run that starts at or after splitPoint
        Run? firstRightRun = null;
        foreach (var rt in runTexts)
        {
            if (rt.Start >= splitPoint)
            {
                firstRightRun = rt.Run;
                break;
            }
        }

        if (firstRightRun == null)
        {
            // All text before split — insert after paragraph
            var bodyPath = parentPath.Contains('/') ? parentPath[..parentPath.LastIndexOf('/')] : "/body";
            return Add(bodyPath, type, InsertPosition.AfterElement(parentPath.Split('/').Last()), properties);
        }

        // Create a new paragraph for the right portion, inheriting paragraph properties
        var rightPara = new Paragraph();
        if (para.ParagraphProperties != null)
            rightPara.ParagraphProperties = (ParagraphProperties)para.ParagraphProperties.CloneNode(true);
        AssignParaId(rightPara);

        // Move runs from firstRightRun onwards to the new paragraph
        var runsToMove = new List<OpenXmlElement>();
        OpenXmlElement? current = firstRightRun;
        while (current != null)
        {
            runsToMove.Add(current);
            current = current.NextSibling();
            // Stop if we hit another paragraph-level structure (shouldn't happen normally)
        }
        // Filter: only move runs and inline elements, not ParagraphProperties
        foreach (var elem in runsToMove)
        {
            if (elem is ParagraphProperties) continue;
            elem.Remove();
            rightPara.AppendChild(elem);
        }

        // Collect existing children before Add, so we can find the newly added element
        var parentOfPara = para.Parent!;
        var childrenBefore = new HashSet<OpenXmlElement>(parentOfPara.ChildElements);

        // Insert rightPara after the original paragraph
        para.InsertAfterSelf(rightPara);

        // Add the block element via normal Add (appends before sectPr)
        var bodyParentPath = parentPath.Contains('/') ? parentPath[..parentPath.LastIndexOf('/')] : "/body";
        var result = Add(bodyParentPath, type, null, properties);

        // Find the newly added element (the one not in childrenBefore and not rightPara)
        OpenXmlElement? addedElement = null;
        foreach (var child in parentOfPara.ChildElements)
        {
            if (!childrenBefore.Contains(child) && child != rightPara)
            {
                addedElement = child;
                break;
            }
        }

        // Move it between para and rightPara
        if (addedElement != null)
        {
            addedElement.Remove();
            parentOfPara.InsertAfter(addedElement, para);
        }

        _doc.MainDocumentPart?.Document?.Save();
        return result;
    }

    /// <summary>
    /// Ensure Columns exists in SectionProperties in correct schema order.
    /// Schema order: ..., PageMargin, ..., Columns, ...
    /// </summary>
    private static Columns EnsureColumns(SectionProperties sectPr)
    {
        var existing = sectPr.GetFirstChild<Columns>();
        if (existing != null) return existing;

        var cols = new Columns();
        var pm = sectPr.GetFirstChild<PageMargin>();
        if (pm != null)
            pm.InsertAfterSelf(cols);
        else
        {
            var pgSz = sectPr.GetFirstChild<PageSize>();
            if (pgSz != null)
                pgSz.InsertAfterSelf(cols);
            else
            {
                // Insert after SectionType, or after last headerReference/footerReference
                var sectionType = sectPr.GetFirstChild<SectionType>();
                if (sectionType != null)
                    sectionType.InsertAfterSelf(cols);
                else
                {
                    OpenXmlElement? lastRef = null;
                    foreach (var child in sectPr.ChildElements)
                    {
                        if (child is HeaderReference || child is FooterReference)
                            lastRef = child;
                    }
                    if (lastRef != null)
                        lastRef.InsertAfterSelf(cols);
                    else
                        sectPr.PrependChild(cols);
                }
            }
        }
        return cols;
    }

    /// <summary>
    /// Ensure PageSize exists in SectionProperties in correct schema order.
    /// Schema order: SectionType, PageSize, PageMargin, ...
    /// </summary>
    private static PageSize EnsureSectPrPageSize(SectionProperties sectPr)
    {
        var existing = sectPr.GetFirstChild<PageSize>();
        if (existing != null) return existing;

        var ps = new PageSize();
        // Insert after SectionType if present, then after FooterReference/HeaderReference,
        // otherwise prepend. OOXML schema order: headerReference*, footerReference*, ..., sectType, pgSz, pgMar
        var sectionType = sectPr.GetFirstChild<SectionType>();
        if (sectionType != null)
        {
            sectionType.InsertAfterSelf(ps);
        }
        else
        {
            // Find the last HeaderReference or FooterReference to insert after
            OpenXmlElement? lastRef = null;
            foreach (var child in sectPr.ChildElements)
            {
                if (child is HeaderReference || child is FooterReference)
                    lastRef = child;
            }
            if (lastRef != null)
                lastRef.InsertAfterSelf(ps);
            else
                sectPr.PrependChild(ps);
        }
        return ps;
    }

    /// <summary>
    /// Ensure PageMargin exists in SectionProperties in correct schema order.
    /// Schema order: SectionType, PageSize, PageMargin, ...
    /// </summary>
    private static PageMargin EnsureSectPrPageMargin(SectionProperties sectPr)
    {
        var existing = sectPr.GetFirstChild<PageMargin>();
        if (existing != null) return existing;

        var pm = new PageMargin();
        // Insert after PageSize if present, after SectionType, after last headerRef/footerRef, or prepend
        var pageSize = sectPr.GetFirstChild<PageSize>();
        if (pageSize != null)
            pageSize.InsertAfterSelf(pm);
        else
        {
            var sectionType = sectPr.GetFirstChild<SectionType>();
            if (sectionType != null)
                sectionType.InsertAfterSelf(pm);
            else
            {
                OpenXmlElement? lastRef = null;
                foreach (var child in sectPr.ChildElements)
                {
                    if (child is HeaderReference || child is FooterReference)
                        lastRef = child;
                }
                if (lastRef != null)
                    lastRef.InsertAfterSelf(pm);
                else
                    sectPr.PrependChild(pm);
            }
        }
        return pm;
    }

    // ==================== w14 Text Effects ====================

    private const string W14Ns = "http://schemas.microsoft.com/office/word/2010/wordml";

    /// <summary>
    /// Remove an existing w14 element from RunProperties by local name.
    /// </summary>
    private static void RemoveW14Element(RunProperties rPr, string localName)
    {
        var existing = rPr.ChildElements
            .Where(e => e.LocalName == localName && e.NamespaceUri == W14Ns)
            .ToList();
        foreach (var e in existing) e.Remove();
    }

    /// <summary>
    /// Split a w14 effect value string by ';' (preferred) or '-' (legacy fallback).
    /// ';' is unambiguous; '-' is only used as fallback when no ';' is present.
    /// </summary>
    private static string[] SplitEffectValue(string value) =>
        value.Contains(';') ? value.Split(';') : value.Split('-');

    /// <summary>
    /// Build w14:textOutline XML.
    /// Format: "WIDTH;COLOR" (e.g. "0.5pt;FF0000"), "WIDTH" (defaults to black), or "none"
    /// Width in pt, internally stored in EMU (1pt = 12700 EMU).
    /// Legacy: "WIDTH-COLOR" also accepted.
    /// </summary>
    internal static string BuildW14TextOutline(string value)
    {
        var parts = SplitEffectValue(value);
        var widthPt = ParseHelpers.SafeParseDouble(parts[0].Replace("pt", ""), "textOutline width");
        var widthEmu = (long)(widthPt * 12700);
        var color = parts.Length > 1 ? ParseHelpers.SanitizeColorForOoxml(parts[1]).Rgb : "000000";

        return $@"<w14:textOutline xmlns:w14=""{W14Ns}"" w14:w=""{widthEmu}"" w14:cap=""flat"" w14:cmpd=""sng"" w14:algn=""ctr""><w14:solidFill><w14:srgbClr w14:val=""{color}""/></w14:solidFill><w14:prstDash w14:val=""solid""/></w14:textOutline>";
    }

    /// <summary>
    /// Build w14:textFill XML.
    /// Format: "C1;C2[;ANGLE]" for linear gradient, "radial:C1;C2" for radial, or single color for solid.
    /// Legacy: '-' separator also accepted.
    /// </summary>
    internal static string BuildW14TextFill(string value)
    {
        if (value.StartsWith("radial:", StringComparison.OrdinalIgnoreCase))
        {
            var radParts = SplitEffectValue(value[7..]);
            var (c1, _) = ParseHelpers.SanitizeColorForOoxml(radParts[0]);
            var c2 = radParts.Length > 1 ? ParseHelpers.SanitizeColorForOoxml(radParts[1]).Rgb : c1;
            return $@"<w14:textFill xmlns:w14=""{W14Ns}""><w14:gradFill><w14:gsLst><w14:gs w14:pos=""0""><w14:srgbClr w14:val=""{c1}""/></w14:gs><w14:gs w14:pos=""100000""><w14:srgbClr w14:val=""{c2}""/></w14:gs></w14:gsLst><w14:path w14:path=""circle""><w14:fillToRect w14:l=""50000"" w14:t=""50000"" w14:r=""50000"" w14:b=""50000""/></w14:path></w14:gradFill></w14:textFill>";
        }

        var parts = SplitEffectValue(value);
        if (parts.Length == 1)
        {
            // Solid fill
            var (rgb, _) = ParseHelpers.SanitizeColorForOoxml(parts[0]);
            return $@"<w14:textFill xmlns:w14=""{W14Ns}""><w14:solidFill><w14:srgbClr w14:val=""{rgb}""/></w14:solidFill></w14:textFill>";
        }

        // Linear gradient: C1;C2[;angle]
        var (gc1, _a1) = ParseHelpers.SanitizeColorForOoxml(parts[0]);
        var (gc2, _a2) = ParseHelpers.SanitizeColorForOoxml(parts[1]);
        var angle = parts.Length > 2 ? ParseHelpers.SafeParseInt(parts[2], "textFill angle") * 60000 : 0;
        return $@"<w14:textFill xmlns:w14=""{W14Ns}""><w14:gradFill><w14:gsLst><w14:gs w14:pos=""0""><w14:srgbClr w14:val=""{gc1}""/></w14:gs><w14:gs w14:pos=""100000""><w14:srgbClr w14:val=""{gc2}""/></w14:gs></w14:gsLst><w14:lin w14:ang=""{angle}"" w14:scaled=""1""/></w14:gradFill></w14:textFill>";
    }

    /// <summary>
    /// Build w14:shadow XML.
    /// Format: "COLOR[;BLUR[;ANGLE[;DIST[;OPACITY]]]]"
    /// Defaults: blur=4pt, angle=45°, dist=3pt, opacity=40%
    /// Legacy: '-' separator also accepted.
    /// </summary>
    internal static string BuildW14Shadow(string value)
    {
        var parts = SplitEffectValue(value);
        var (color, _) = ParseHelpers.SanitizeColorForOoxml(parts[0]);
        var blurPt = parts.Length > 1 ? ParseHelpers.SafeParseDouble(parts[1], "shadow blur") : 4.0;
        var angleDeg = parts.Length > 2 ? ParseHelpers.SafeParseDouble(parts[2], "shadow angle") : 45.0;
        var distPt = parts.Length > 3 ? ParseHelpers.SafeParseDouble(parts[3], "shadow distance") : 3.0;
        var opacity = parts.Length > 4 ? ParseHelpers.SafeParseDouble(parts[4], "shadow opacity") : 40.0;

        var blurEmu = (long)(blurPt * 12700);
        var distEmu = (long)(distPt * 12700);
        var angleOoxml = (int)(angleDeg * 60000);
        var alphaVal = (int)(opacity * 1000);

        return $@"<w14:shadow xmlns:w14=""{W14Ns}"" w14:blurRad=""{blurEmu}"" w14:dist=""{distEmu}"" w14:dir=""{angleOoxml}"" w14:sx=""100000"" w14:sy=""100000"" w14:kx=""0"" w14:ky=""0"" w14:algn=""tl""><w14:srgbClr w14:val=""{color}""><w14:alpha w14:val=""{alphaVal}""/></w14:srgbClr></w14:shadow>";
    }

    /// <summary>
    /// Build w14:glow XML.
    /// Format: "COLOR[;RADIUS[;OPACITY]]"
    /// Defaults: radius=8pt, opacity=75%
    /// Legacy: '-' separator also accepted.
    /// </summary>
    internal static string BuildW14Glow(string value)
    {
        var parts = SplitEffectValue(value);
        var (color, _) = ParseHelpers.SanitizeColorForOoxml(parts[0]);
        var radiusPt = parts.Length > 1 ? ParseHelpers.SafeParseDouble(parts[1], "glow radius") : 8.0;
        var opacity = parts.Length > 2 ? ParseHelpers.SafeParseDouble(parts[2], "glow opacity") : 75.0;

        var radiusEmu = (long)(radiusPt * 12700);
        var alphaVal = (int)(opacity * 1000);

        return $@"<w14:glow xmlns:w14=""{W14Ns}"" w14:rad=""{radiusEmu}""><w14:srgbClr w14:val=""{color}""><w14:alpha w14:val=""{alphaVal}""/></w14:srgbClr></w14:glow>";
    }

    /// <summary>
    /// Build w14:reflection XML.
    /// Values: "tight"/"small", "half"/"true", "full"
    /// </summary>
    internal static string BuildW14Reflection(string value)
    {
        var endPos = value.ToLowerInvariant() switch
        {
            "tight" or "small" => 55000,
            "true" or "half" => 90000,
            "full" => 100000,
            _ => int.TryParse(value, out var pct) ? (int)Math.Min((long)pct * 1000, 100000) : 90000
        };

        return $@"<w14:reflection xmlns:w14=""{W14Ns}"" w14:blurRad=""6350"" w14:stA=""52000"" w14:stPos=""0"" w14:endA=""300"" w14:endPos=""{endPos}"" w14:dist=""0"" w14:dir=""5400000"" w14:fadeDir=""5400000"" w14:sx=""100000"" w14:sy=""-100000"" w14:kx=""0"" w14:ky=""0"" w14:algn=""bl""/>";
    }

    /// <summary>
    /// Apply a w14 text effect to a run's RunProperties.
    /// Handles set and remove logic.
    /// </summary>
    internal static void ApplyW14TextEffect(Run run, string effectName, string value, Func<string, string> builder)
    {
        var rPr = EnsureRunProperties(run);
        RemoveW14Element(rPr, effectName);

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("false", StringComparison.OrdinalIgnoreCase))
            return;

        var xml = builder(value);
        var element = new OpenXmlUnknownElement("w14", "tmp", W14Ns);
        element.InnerXml = xml;
        var child = element.FirstChild;
        if (child != null)
        {
            child.Remove();
            rPr.AppendChild(child);
        }
    }

    /// <summary>
    /// Read w14 text effect values from RunProperties.
    /// Returns a dictionary of effect names to their parsed values.
    /// </summary>
    internal static void ReadW14TextEffects(RunProperties? rPr, DocumentNode node)
    {
        if (rPr == null) return;

        foreach (var child in rPr.ChildElements)
        {
            if (child.NamespaceUri != W14Ns) continue;

            switch (child.LocalName)
            {
                case "textOutline":
                {
                    var wAttr = child.GetAttributes().FirstOrDefault(a => a.LocalName == "w");
                    var widthEmu = long.TryParse(wAttr.Value, out var w) ? w : 0;
                    var widthPt = widthEmu / 12700.0;
                    var colorMatch = System.Text.RegularExpressions.Regex.Match(
                        child.InnerXml, @"val=""([0-9A-Fa-f]{6})""");
                    var color = colorMatch.Success ? ParseHelpers.FormatHexColor(colorMatch.Groups[1].Value) : "";
                    node.Format["textOutline"] = string.IsNullOrEmpty(color) ? $"{widthPt:0.##}pt" : $"{widthPt:0.##}pt;{color}";
                    break;
                }
                case "textFill":
                {
                    var innerXml = child.InnerXml;
                    if (innerXml.Contains("gradFill"))
                    {
                        var colors = new List<string>();
                        foreach (System.Text.RegularExpressions.Match m in
                            System.Text.RegularExpressions.Regex.Matches(innerXml, @"val=""([0-9A-Fa-f]{6})"""))
                            colors.Add(m.Groups[1].Value);

                        // Add # prefix to gradient colors
                        for (int ci = 0; ci < colors.Count; ci++)
                            colors[ci] = ParseHelpers.FormatHexColor(colors[ci]);

                        var isRadial = innerXml.Contains("<w14:path");
                        if (isRadial && colors.Count >= 2)
                            node.Format["textFill"] = $"radial:{colors[0]};{colors[1]}";
                        else if (colors.Count >= 2)
                        {
                            var angleMatch = System.Text.RegularExpressions.Regex.Match(innerXml, @"ang=""(\d+)""");
                            var angle = angleMatch.Success ? int.Parse(angleMatch.Groups[1].Value) / 60000.0 : 0.0;
                            var angleStr = angle % 1 == 0 ? $"{(int)angle}" : $"{angle:0.##}";
                            node.Format["textFill"] = $"{colors[0]};{colors[1]};{angleStr}";
                        }
                        else if (colors.Count == 1)
                            node.Format["textFill"] = colors[0];
                    }
                    else if (innerXml.Contains("solidFill"))
                    {
                        var colorMatch = System.Text.RegularExpressions.Regex.Match(
                            innerXml, @"val=""([0-9A-Fa-f]{6})""");
                        if (colorMatch.Success)
                            node.Format["textFill"] = ParseHelpers.FormatHexColor(colorMatch.Groups[1].Value);
                    }
                    break;
                }
                case "shadow":
                {
                    var attrs = child.GetAttributes().ToDictionary(a => a.LocalName, a => a.Value);
                    var colorMatch = System.Text.RegularExpressions.Regex.Match(
                        child.InnerXml, @"val=""([0-9A-Fa-f]{6})""");
                    var color = colorMatch.Success ? ParseHelpers.FormatHexColor(colorMatch.Groups[1].Value) : "#000000";
                    var blurEmu = attrs.TryGetValue("blurRad", out var br) && long.TryParse(br, out var blurVal) ? blurVal : 0;
                    var blurPt = blurEmu / 12700.0;
                    var dirVal = attrs.TryGetValue("dir", out var dir) && long.TryParse(dir, out var dirLong) ? dirLong : 0;
                    var angleDeg = dirVal / 60000.0;
                    var distEmu = attrs.TryGetValue("dist", out var dist) && long.TryParse(dist, out var distLong) ? distLong : 0;
                    var distPt = distEmu / 12700.0;
                    // Read alpha (opacity) from inner srgbClr child
                    var alphaMatch = System.Text.RegularExpressions.Regex.Match(
                        child.InnerXml, @"alpha[^>]*val=""(\d+)""");
                    var opacity = alphaMatch.Success && double.TryParse(alphaMatch.Groups[1].Value, out var alphaVal) ? alphaVal / 1000.0 : 100.0;
                    node.Format["w14shadow"] = $"{color};{blurPt:0.##};{angleDeg:0.##};{distPt:0.##};{opacity:0.##}";
                    break;
                }
                case "glow":
                {
                    var radAttr = child.GetAttributes().FirstOrDefault(a => a.LocalName == "rad");
                    var radiusEmu = long.TryParse(radAttr.Value, out var r) ? r : 0;
                    var radiusPt = radiusEmu / 12700.0;
                    var colorMatch = System.Text.RegularExpressions.Regex.Match(
                        child.InnerXml, @"val=""([0-9A-Fa-f]{6})""");
                    var color = colorMatch.Success ? ParseHelpers.FormatHexColor(colorMatch.Groups[1].Value) : "#000000";
                    // Read alpha (opacity) from inner srgbClr child
                    var alphaMatch = System.Text.RegularExpressions.Regex.Match(
                        child.InnerXml, @"alpha[^>]*val=""(\d+)""");
                    var opacity = alphaMatch.Success && double.TryParse(alphaMatch.Groups[1].Value, out var av) ? av / 1000.0 : 100.0;
                    node.Format["w14glow"] = $"{color};{radiusPt:0.##};{opacity:0.##}";
                    break;
                }
                case "reflection":
                {
                    var endPosAttr = child.GetAttributes().FirstOrDefault(a => a.LocalName == "endPos");
                    var endPos = int.TryParse(endPosAttr.Value, out var ep) ? ep : 90000;
                    node.Format["w14reflection"] = endPos switch
                    {
                        <= 55000 => "tight",
                        <= 90000 => "half",
                        _ => "full"
                    };
                    break;
                }
            }
        }
    }

    // ==================== Extended Chart Helpers ====================

    private const string WordChartExUri = "http://schemas.microsoft.com/office/drawing/2014/chartex";
    private const string WordChartUri = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    /// <summary>
    /// Count all charts (both standard ChartPart and ExtendedChartPart) in the document.
    /// </summary>
    private static int CountWordCharts(MainDocumentPart mainPart)
    {
        return mainPart.ChartParts.Count() + mainPart.ExtendedChartParts.Count();
    }

    /// <summary>
    /// Represents a chart part in Word that could be either a standard ChartPart or an ExtendedChartPart.
    /// </summary>
    private class WordChartInfo
    {
        public ChartPart? StandardPart { get; set; }
        public ExtendedChartPart? ExtendedPart { get; set; }
        public DW.DocProperties? DocProperties { get; set; }
        /// <summary>
        /// The <c>wp:inline</c> element that hosts this chart — needed by
        /// chart position Set to mutate the <c>wp:extent</c> child.
        /// </summary>
        public DW.Inline? Inline { get; set; }
        public bool IsExtended => ExtendedPart != null;
    }

    /// <summary>
    /// Get all chart parts (standard + extended) in document order by walking Drawing/Inline elements.
    /// </summary>
    private List<WordChartInfo> GetAllWordCharts()
    {
        var result = new List<WordChartInfo>();
        var mainPart = _doc.MainDocumentPart;
        if (mainPart?.Document?.Body == null) return result;

        foreach (var inline in mainPart.Document.Body.Descendants<DW.Inline>())
        {
            var graphicData = inline.Descendants<A.GraphicData>().FirstOrDefault();
            if (graphicData == null) continue;

            var docProps = inline.Descendants<DW.DocProperties>().FirstOrDefault();

            if (graphicData.Uri == WordChartUri)
            {
                // Standard chart
                var chartRef = graphicData.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().FirstOrDefault();
                if (chartRef?.Id?.Value == null) continue;
                try
                {
                    var chartPart = (ChartPart)mainPart.GetPartById(chartRef.Id.Value);
                    result.Add(new WordChartInfo { StandardPart = chartPart, DocProperties = docProps, Inline = inline });
                }
                catch { /* skip invalid references */ }
            }
            else if (graphicData.Uri == WordChartExUri)
            {
                // Extended chart (funnel, treemap, etc.)
                var relId = GetWordExtendedChartRelId(inline);
                if (relId == null) continue;
                try
                {
                    var extPart = (ExtendedChartPart)mainPart.GetPartById(relId);
                    result.Add(new WordChartInfo { ExtendedPart = extPart, DocProperties = docProps, Inline = inline });
                }
                catch { /* skip invalid references */ }
            }
        }

        return result;
    }

    /// <summary>
    /// Apply <c>width</c> / <c>height</c> to a Word inline chart's
    /// <c>wp:extent</c>. Accepts unit-qualified sizes (`6cm`, `2in`,
    /// `720pt`) or raw EMU integers via EmuConverter.
    ///
    /// CONSISTENCY(chart-position-set): mirrors the PPTX and Excel path.
    /// Word inline charts have no absolute x/y (they flow with text), so
    /// those keys — if provided — are appended to <paramref name="unsupported"/>
    /// rather than silently dropped.
    /// </summary>
    private static void ApplyWordChartPositionSet(
        DW.Inline inline, Dictionary<string, string> properties, List<string> unsupported)
    {
        var extent = inline.Extent;
        if (extent == null) return;

        // x/y are meaningless for inline charts.
        foreach (var k in new[] { "x", "y" })
        {
            var matched = properties.Keys
                .FirstOrDefault(key => key.Equals(k, StringComparison.OrdinalIgnoreCase));
            if (matched == null) continue;
            unsupported.Add(matched);
            Console.Error.WriteLine(
                $"Warning: '{matched}' is ignored on Word inline charts — inline elements have no absolute position. " +
                "For positioned charts, switch to anchor mode (not currently supported).");
        }

        if (properties.TryGetValue("width", out var wStr))
        {
            try { extent.Cx = OfficeCli.Core.EmuConverter.ParseEmu(wStr); }
            catch { unsupported.Add("width"); }
        }

        if (properties.TryGetValue("height", out var hStr))
        {
            try { extent.Cy = OfficeCli.Core.EmuConverter.ParseEmu(hStr); }
            catch { unsupported.Add("height"); }
        }
    }

    /// <summary>
    /// Get the relationship ID from an extended chart inline Drawing element.
    /// </summary>
    private static string? GetWordExtendedChartRelId(DW.Inline inline)
    {
        var gd = inline.Descendants<A.GraphicData>().FirstOrDefault(g => g.Uri == WordChartExUri);
        if (gd == null) return null;
        var typed = gd.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.RelId>().FirstOrDefault();
        if (typed?.Id?.Value != null) return typed.Id.Value;
        foreach (var child in gd.ChildElements)
        {
            var rId = child.GetAttributes().FirstOrDefault(a =>
                a.LocalName == "id" && a.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            if (rId.Value != null) return rId.Value;
        }
        return null;
    }

    /// <summary>
    /// Get current document protection mode and enforcement status.
    /// </summary>
    private (string mode, bool enforced) GetDocumentProtection()
    {
        var settings = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings;
        var docProtection = settings?.GetFirstChild<DocumentProtection>();
        if (docProtection == null)
            return ("none", false);

        var mode = docProtection.Edit?.InnerText switch
        {
            "readOnly" => "readOnly",
            "comments" => "comments",
            "trackedChanges" => "trackedChanges",
            "forms" => "forms",
            _ => "none"
        };
        var enforced = docProtection.Enforcement?.Value == true
            || (docProtection.Enforcement?.Value == null && docProtection.Edit != null);
        return (mode, enforced);
    }

    /// <summary>
    /// Check if an SDT element is editable based on its lock attribute and the current document protection.
    /// </summary>
    private bool IsSdtEditable(SdtProperties? sdtProps)
    {
        var (mode, enforced) = GetDocumentProtection();

        // No protection or not enforced → all SDTs are editable
        if (!enforced || mode == "none")
            return true;

        // readOnly protection → SDTs are not editable (unless in permRange, P2)
        if (mode == "readOnly")
            return false;

        // forms protection → SDTs are editable unless content-locked
        if (mode == "forms")
        {
            var lockEl = sdtProps?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Lock>();
            var lockVal = lockEl?.Val?.Value;
            return lockVal != LockingValues.ContentLocked && lockVal != LockingValues.SdtContentLocked;
        }

        // comments/trackedChanges → not typically editable
        return false;
    }

    /// <summary>
    /// Generate a unique 8-character uppercase hex ID for w14:paraId / w14:textId.
    /// OOXML spec requires value &lt; 0x80000000 (MaxExclusive).
    /// Uses deterministic increment from _nextParaId, wraps around on overflow,
    /// skips IDs already in use.
    /// </summary>
    private string GenerateParaId()
    {
        const int maxExclusive = 0x7FFFFFFF; // OOXML spec limit
        const int minStartId = 0x100000;
        var startId = _nextParaId;
        while (true)
        {
            var id = _nextParaId.ToString("X8");
            _nextParaId++;
            if (_nextParaId > maxExclusive)
                _nextParaId = minStartId;
            if (_usedParaIds.Add(id))
                return id;
            // Safety: if we've wrapped all the way around, something is very wrong
            if (_nextParaId == startId)
                throw new InvalidOperationException("No available paraId slots");
        }
    }

    /// <summary>
    /// Assign paraId and textId to a paragraph if not already set.
    /// </summary>
    private void AssignParaId(Paragraph para)
    {
        if (string.IsNullOrEmpty(para.ParagraphId?.Value))
            para.ParagraphId = GenerateParaId();
        if (string.IsNullOrEmpty(para.TextId?.Value))
            para.TextId = GenerateParaId();
    }

    /// <summary>
    /// Ensure all paragraphs in the document have w14:paraId and w14:textId.
    /// Called on document open.
    /// </summary>
    private void EnsureAllParaIds()
    {
        var mainPart = _doc.MainDocumentPart;
        if (mainPart?.Document?.Body == null) return;

        _usedParaIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // CONSISTENCY(paraid-global-uniqueness): paraId is allocated from a
        // single _nextParaId counter shared across the entire handler, so
        // EVERY part that can hold paragraphs must contribute to the
        // collision set. Body + headers + footers were already covered;
        // footnotes/endnotes/comments were missed, letting newly generated
        // paraIds collide with paraIds Word had already written into those
        // parts (rare in practice but a real correctness gap).
        var allParagraphs = mainPart.Document.Body.Descendants<Paragraph>().AsEnumerable();
        foreach (var headerPart in mainPart.HeaderParts)
            if (headerPart.Header != null)
                allParagraphs = allParagraphs.Concat(headerPart.Header.Descendants<Paragraph>());
        foreach (var footerPart in mainPart.FooterParts)
            if (footerPart.Footer != null)
                allParagraphs = allParagraphs.Concat(footerPart.Footer.Descendants<Paragraph>());
        if (mainPart.FootnotesPart?.Footnotes != null)
            allParagraphs = allParagraphs.Concat(mainPart.FootnotesPart.Footnotes.Descendants<Paragraph>());
        if (mainPart.EndnotesPart?.Endnotes != null)
            allParagraphs = allParagraphs.Concat(mainPart.EndnotesPart.Endnotes.Descendants<Paragraph>());
        if (mainPart.WordprocessingCommentsPart?.Comments != null)
            allParagraphs = allParagraphs.Concat(mainPart.WordprocessingCommentsPart.Comments.Descendants<Paragraph>());

        var paragraphs = allParagraphs.ToList();

        // Collect existing IDs, detect duplicates, and track max for deterministic increment
        var paraIdSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int maxId = 0;

        foreach (var para in paragraphs)
        {
            // Fix duplicate paraId: if already seen, clear it so it gets reassigned below
            if (!string.IsNullOrEmpty(para.ParagraphId?.Value))
            {
                if (!paraIdSeen.Add(para.ParagraphId.Value))
                {
                    para.ParagraphId = null!; // duplicate — will be reassigned
                }
                else
                {
                    _usedParaIds.Add(para.ParagraphId.Value);
                    if (int.TryParse(para.ParagraphId.Value, System.Globalization.NumberStyles.HexNumber, null, out var numId) && numId > maxId)
                        maxId = numId;
                }
            }
            if (!string.IsNullOrEmpty(para.TextId?.Value))
            {
                _usedParaIds.Add(para.TextId.Value);
                if (int.TryParse(para.TextId.Value, System.Globalization.NumberStyles.HexNumber, null, out var numId) && numId > maxId)
                    maxId = numId;
            }
        }

        // Start deterministic increment from max+1, minimum 0x100000 to avoid conflicts with small IDs
        const int minStartId = 0x100000;
        _nextParaId = Math.Max(maxId + 1, minStartId);
        if (_nextParaId > 0x7FFFFFFF) _nextParaId = minStartId;

        // Assign IDs to paragraphs that don't have them (including cleared duplicates)
        foreach (var para in paragraphs)
        {
            if (string.IsNullOrEmpty(para.ParagraphId?.Value))
                para.ParagraphId = GenerateParaId();
            if (string.IsNullOrEmpty(para.TextId?.Value))
                para.TextId = GenerateParaId();
        }

        // Ensure mc:Ignorable includes "w14" so Word 2007 skips w14:paraId/textId attributes
        var doc = mainPart.Document;
        const string mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        if (doc.LookupNamespace("mc") == null)
            doc.AddNamespaceDeclaration("mc", mcNs);
        if (doc.LookupNamespace("w14") == null)
            doc.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
        var ignorable = doc.MCAttributes?.Ignorable?.Value ?? "";
        if (!ignorable.Contains("w14"))
        {
            doc.MCAttributes ??= new DocumentFormat.OpenXml.MarkupCompatibilityAttributes();
            doc.MCAttributes.Ignorable = string.IsNullOrEmpty(ignorable) ? "w14" : $"{ignorable} w14";
        }
    }

    // ==================== SDT IDs (content controls) ====================

    /// <summary>
    /// Generate a deterministic unique SdtId by scanning max existing value + 1.
    /// </summary>
    private int NextSdtId()
    {
        const int overflowReset = 872011;
        int maxId = 0;
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body != null)
        {
            foreach (var sdtId in body.Descendants<SdtId>())
            {
                if (sdtId.Val?.HasValue == true && sdtId.Val.Value > maxId)
                    maxId = sdtId.Val.Value;
            }
        }
        var next = maxId + 1;
        return next > int.MaxValue - 1 ? overflowReset : next;
    }

    // ==================== DocPr IDs (pictures, charts) ====================

    /// <summary>
    /// Ensure all DocProperties in the document have unique IDs.
    /// Called on document open.
    /// </summary>
    private void EnsureDocPropIds()
    {
        var mainPart = _doc.MainDocumentPart;
        if (mainPart?.Document?.Body == null) return;

        var allDocProps = mainPart.Document.Body.Descendants<DW.DocProperties>().ToList();

        foreach (var headerPart in mainPart.HeaderParts)
            if (headerPart.Header != null)
                allDocProps.AddRange(headerPart.Header.Descendants<DW.DocProperties>());
        foreach (var footerPart in mainPart.FooterParts)
            if (footerPart.Footer != null)
                allDocProps.AddRange(footerPart.Footer.Descendants<DW.DocProperties>());

        var usedIds = new HashSet<uint>();
        var duplicates = new List<DW.DocProperties>();

        foreach (var dp in allDocProps)
        {
            if (dp.Id?.HasValue == true && !usedIds.Add(dp.Id.Value))
                duplicates.Add(dp);
            else if (dp.Id?.HasValue != true)
                duplicates.Add(dp);
        }

        foreach (var dp in duplicates)
        {
            uint newId = 1;
            while (!usedIds.Add(newId)) newId++;
            dp.Id = newId;
        }
    }
}
