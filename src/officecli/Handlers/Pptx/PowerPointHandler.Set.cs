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
    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        path = NormalizePptxPathSegmentCasing(path);
        path = NormalizeCellPath(path);
        path = ResolveIdPath(path);

        // Batch Set: if path looks like a selector (not starting with /), Query → Set each
        if (!string.IsNullOrEmpty(path) && !path.StartsWith("/"))
        {
            var unsupported = new List<string>();
            var targets = Query(path);
            if (targets.Count == 0)
                throw new ArgumentException($"No elements matched selector: {path}");
            foreach (var target in targets)
            {
                var targetUnsupported = Set(target.Path, properties);
                foreach (var u in targetUnsupported)
                    if (!unsupported.Contains(u)) unsupported.Add(u);
            }
            return unsupported;
        }

        if (path.Equals("/theme", StringComparison.OrdinalIgnoreCase))
            return SetThemeProperties(properties);

        // Unified find: if 'find' key is present, route to ProcessPptFind
        if (properties.TryGetValue("find", out var findText))
        {
            var replace = properties.TryGetValue("replace", out var r) ? r : null;
            var formatProps = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
            formatProps.Remove("find");
            formatProps.Remove("replace");
            formatProps.Remove("scope");
            formatProps.Remove("regex");

            if (replace == null && formatProps.Count == 0)
                throw new ArgumentException("'find' requires either 'replace' and/or format properties (e.g. bold, color, size).");

            // Support regex=true as an alternative to r"..." prefix.
            // CONSISTENCY(find-regex): mirror of WordHandler.Set.cs:60-61. grep
            // "CONSISTENCY(find-regex)" for every project-wide call site.
            if (properties.TryGetValue("regex", out var regexFlag) && ParseHelpers.IsTruthySafe(regexFlag) && !findText.StartsWith("r\"") && !findText.StartsWith("r'"))
                findText = $"r\"{findText}\"";

            var matchCount = ProcessPptFind(path, findText, replace, formatProps);
            LastFindMatchCount = matchCount;
            return [];
        }

        // Presentation-level properties: / or /presentation
        if (path is "/" or "" or "/presentation")
        {

            var presentation = _doc.PresentationPart?.Presentation
                ?? throw new InvalidOperationException("No presentation");
            var unsupported = new List<string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "slidewidth" or "width":
                        var sldSz = presentation.GetFirstChild<SlideSize>()
                            ?? presentation.AppendChild(new SlideSize());
                        sldSz.Cx = Core.EmuConverter.ParseEmuAsInt(value);
                        sldSz.Type = SlideSizeValues.Custom;
                        break;
                    case "slideheight" or "height":
                        var sldSz2 = presentation.GetFirstChild<SlideSize>()
                            ?? presentation.AppendChild(new SlideSize());
                        sldSz2.Cy = Core.EmuConverter.ParseEmuAsInt(value);
                        sldSz2.Type = SlideSizeValues.Custom;
                        break;
                    case "slidesize":
                        var sz = presentation.GetFirstChild<SlideSize>()
                            ?? presentation.AppendChild(new SlideSize());
                        if (SlideSizeDefaults.Presets.TryGetValue(value, out var preset))
                        {
                            sz.Cx = (int)preset.Cx;
                            sz.Cy = (int)preset.Cy;
                            sz.Type = preset.Type;
                        }
                        else
                        {
                            unsupported.Add(key);
                        }
                        break;
                    // Core document properties
                    case "title":
                        _doc.PackageProperties.Title = value;
                        break;
                    case "author" or "creator":
                        _doc.PackageProperties.Creator = value;
                        break;
                    case "subject":
                        _doc.PackageProperties.Subject = value;
                        break;
                    case "description":
                        _doc.PackageProperties.Description = value;
                        break;
                    case "category":
                        _doc.PackageProperties.Category = value;
                        break;
                    case "keywords":
                        _doc.PackageProperties.Keywords = value;
                        break;
                    case "lastmodifiedby":
                        _doc.PackageProperties.LastModifiedBy = value;
                        break;
                    case "revision":
                        _doc.PackageProperties.Revision = value;
                        break;
                    case "defaultfont" or "font":
                    {
                        var masterPart = _doc.PresentationPart?.SlideMasterParts?.FirstOrDefault();
                        var theme = masterPart?.ThemePart?.Theme;
                        var fontScheme = theme?.ThemeElements?.FontScheme;
                        if (fontScheme != null)
                        {
                            if (fontScheme.MajorFont?.LatinFont != null)
                                fontScheme.MajorFont.LatinFont.Typeface = value;
                            if (fontScheme.MinorFont?.LatinFont != null)
                                fontScheme.MinorFont.LatinFont.Typeface = value;
                            masterPart!.ThemePart!.Theme!.Save();
                        }
                        break;
                    }
                    default:
                        var lowerKey = key.ToLowerInvariant();
                        if (!TrySetPresentationSetting(lowerKey, value)
                            && !Core.ThemeHandler.TrySetTheme(
                                _doc.PresentationPart?.SlideMasterParts?.FirstOrDefault()?.ThemePart, lowerKey, value)
                            && !Core.ExtendedPropertiesHandler.TrySetExtendedProperty(
                                Core.ExtendedPropertiesHandler.GetOrCreateExtendedPart(_doc), lowerKey, value))
                        {
                            if (unsupported.Count == 0)
                                unsupported.Add($"{key} (valid presentation props: slideWidth, slideHeight, slideSize, title, author, defaultFont, firstSlideNum, rtl, compatMode, print.*, show.*)");
                            else
                                unsupported.Add(key);
                        }
                        break;
                }
            }
            presentation.Save();
            return unsupported;
        }

        // Try slidemaster/slidelayout bg-aware path first (case-insensitive):
        // /slidemaster[N], /slidemaster[N]/slidelayout[M], /slidelayout[N]
        // Handles background and name props. Falls through for shape-nested paths.
        {
            var masterBgMatch = Regex.Match(path, @"^/slidemaster\[(\d+)\](?:/slidelayout\[(\d+)\])?$", RegexOptions.IgnoreCase);
            var layoutBgMatch = Regex.Match(path, @"^/slidelayout\[(\d+)\]$", RegexOptions.IgnoreCase);
            if (masterBgMatch.Success || layoutBgMatch.Success)
                return SetMasterOrLayoutBackgroundByPath(masterBgMatch, layoutBgMatch, properties);
        }

        // Try slideMaster/slideLayout shape editing: /slideMaster[N]/shape[M] or /slideLayout[N]/shape[M]
        var masterShapeMatch = Regex.Match(path, @"^/(slideMaster|slideLayout)\[(\d+)\](?:/(\w+)\[(\d+)\])?$");
        if (masterShapeMatch.Success) return SetMasterShapeByPath(masterShapeMatch, properties);

        // Try notes path: /slide[N]/notes
        var notesSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/notes$");
        if (notesSetMatch.Success) return SetNotesByPath(notesSetMatch, properties);

        // Try animation path: /slide[N]/shape[M]/animation[A]
        var animSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/animation\[(\d+)\]$");
        if (animSetMatch.Success) return SetShapeAnimationByPath(animSetMatch, properties);

        // Try run-level path: /slide[N]/shape[M]/run[K]
        var runMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/run\[(\d+)\]$");
        if (runMatch.Success) return SetShapeRunByPath(runMatch, properties);

        // Try paragraph/run path: /slide[N]/shape[M]/paragraph[P]/run[K]
        var paraRunMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/paragraph\[(\d+)\]/run\[(\d+)\]$");
        if (paraRunMatch.Success) return SetParagraphRunByPath(paraRunMatch, properties);

        // Try paragraph-level path: /slide[N]/shape[M]/paragraph[P]
        var paraMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/paragraph\[(\d+)\]$");
        if (paraMatch.Success) return SetParagraphByPath(paraMatch, properties);

        // Try chart axis-by-role sub-path: /slide[N]/chart[M]/axis[@role=ROLE].
        // Routed separately from the chart[]/series[] path because the role capture
        // needs to drive a different forwarder (SetAxisProperties, not series-prefix).
        var chartAxisSetMatch = Regex.Match(path,
            @"^/slide\[(\d+)\]/chart\[(\d+)\]/axis\[@role=([a-zA-Z0-9_]+)\]$");
        if (chartAxisSetMatch.Success) return SetChartAxisByPath(chartAxisSetMatch, properties);

        // Try chart path: /slide[N]/chart[M] or /slide[N]/chart[M]/series[K]
        var chartSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/chart\[(\d+)\](?:/series\[(\d+)\])?$");
        if (chartSetMatch.Success) return SetChartByPath(chartSetMatch, properties);

        // Try table cell path: /slide[N]/table[M]/tr[R]/tc[C]
        var tblCellMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]/tc\[(\d+)\]$");
        if (tblCellMatch.Success) return SetTableCellByPath(tblCellMatch, properties);

        // Try table-level path: /slide[N]/table[M]
        var tblMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]$");
        if (tblMatch.Success) return SetTableByPath(tblMatch, properties);

        // Try table row path: /slide[N]/table[M]/tr[R]
        var tblRowMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]$");
        if (tblRowMatch.Success) return SetTableRowByPath(tblRowMatch, properties);

        // Try placeholder path: /slide[N]/placeholder[M] or /slide[N]/placeholder[type]
        var phMatch = Regex.Match(path, @"^/slide\[(\d+)\]/placeholder\[(\w+)\]$");
        if (phMatch.Success) return SetPlaceholderByPath(phMatch, properties);

        // Try video/audio path: /slide[N]/video[M] or /slide[N]/audio[M]
        var mediaSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/(video|audio)\[(\d+)\]$");
        if (mediaSetMatch.Success) return SetMediaByPath(mediaSetMatch, properties);

        // Try picture path: /slide[N]/picture[M] or /slide[N]/pic[M]
        // OLE set path: /slide[N]/ole[M]
        // Replace backing embedded part + refresh ProgID automatically
        // when the extension changes. Cleans up the old part to avoid
        // storage bloat (mirrors picture path clean-up).
        var oleSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/(?:ole|object|embed)\[(\d+)\]$");
        if (oleSetMatch.Success) return SetOleByPath(oleSetMatch, properties);

        var picSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/(?:picture|pic)\[(\d+)\]$");
        if (picSetMatch.Success) return SetPictureByPath(picSetMatch, properties);

        // Try slide-level path: /slide[N]
        var slideOnlyMatch = Regex.Match(path, @"^/slide\[(\d+)\]$");
        if (slideOnlyMatch.Success) return SetSlideByPath(slideOnlyMatch, properties);

        // Try model3d-level path: /slide[N]/model3d[M]
        var model3dSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/model3d\[(\d+)\]$");
        if (model3dSetMatch.Success) return SetModel3DByPath(model3dSetMatch, properties);

        // Try zoom-level path: /slide[N]/zoom[M]
        var zoomSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/zoom\[(\d+)\]$");
        if (zoomSetMatch.Success) return SetZoomByPath(zoomSetMatch, properties);

        // Try shape-level path: /slide[N]/shape[M]
        var match = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]$");
        if (match.Success) return SetShapeByPath(match, properties);

        // Try connector path: /slide[N]/connector[M] or /slide[N]/connection[M]
        var cxnMatch = Regex.Match(path, @"^/slide\[(\d+)\]/(?:connector|connection)\[(\d+)\]$");
        if (cxnMatch.Success) return SetConnectorByPath(cxnMatch, properties);

        // Try group inner shape path: /slide[N]/group[M]/shape[K]
        // CONSISTENCY(group-inner-shape): Get supports this; Set must too.
        var grpInnerShapeMatch = Regex.Match(path, @"^/slide\[(\d+)\]/group\[(\d+)\]/shape\[(\d+)\]$");
        if (grpInnerShapeMatch.Success) return SetGroupInnerShapeByPath(grpInnerShapeMatch, properties);

        // Try group path: /slide[N]/group[M]
        var grpMatch = Regex.Match(path, @"^/slide\[(\d+)\]/group\[(\d+)\]$");
        if (grpMatch.Success) return SetGroupByPath(grpMatch, properties);

        // BUG-R36-B11: comment path /slide[N]/comment[M].
        var cmtMatch = Regex.Match(path, @"^/slide\[(\d+)\]/comment\[(\d+)\]$");
        if (cmtMatch.Success)
        {
            var resolved = ResolveSlideComment(path)
                ?? throw new ArgumentException($"Comment not found: {path}");
            var unsupported = SetSlideCommentProperties(resolved.comment, properties);
            resolved.slide.SlideCommentsPart!.CommentList!.Save();
            return unsupported;
        }

        // Generic XML fallback: navigate to element and set attributes
        {
            SlidePart fbSlidePart;
            OpenXmlElement target;

            // Try logical path resolution first (table/placeholder paths)
            var logicalResult = ResolveLogicalPath(path);
            if (logicalResult.HasValue)
            {
                fbSlidePart = logicalResult.Value.slidePart;
                target = logicalResult.Value.element;
            }
            else
            {
                var allSegments = GenericXmlQuery.ParsePathSegments(path);
                if (allSegments.Count == 0 || !allSegments[0].Name.Equals("slide", StringComparison.OrdinalIgnoreCase) || !allSegments[0].Index.HasValue)
                    throw new ArgumentException($"Path must start with /slide[N]: {path}");

                var fbSlideIdx = allSegments[0].Index!.Value;
                var fbSlideParts = GetSlideParts().ToList();
                if (fbSlideIdx < 1 || fbSlideIdx > fbSlideParts.Count)
                    throw new ArgumentException($"Slide {fbSlideIdx} not found (total: {fbSlideParts.Count})");

                fbSlidePart = fbSlideParts[fbSlideIdx - 1];
                var remaining = allSegments.Skip(1).ToList();
                target = GetSlide(fbSlidePart);
                if (remaining.Count > 0)
                {
                    target = GenericXmlQuery.NavigateByPath(target, remaining)
                        ?? throw new ArgumentException($"Element not found: {path}");
                }
            }

            var unsup = new List<string>();
            foreach (var (key, value) in properties)
            {
                if (!GenericXmlQuery.SetGenericAttribute(target, key, value))
                    unsup.Add(key);
            }
            GetSlide(fbSlidePart).Save();
            return unsup;
        }
    }

    // Per-element-type Set helpers live in sibling partial-class files:
    //   PowerPointHandler.Set.Slide.cs    — slide / master / layout / notes
    //   PowerPointHandler.Set.Shape.cs    — shape / paragraph / run / placeholder / group / connector
    //   PowerPointHandler.Set.Table.cs    — table / row / cell
    //   PowerPointHandler.Set.Chart.cs    — chart / chartAxis
    //   PowerPointHandler.Set.Media.cs    — picture / media / OLE / 3D model / zoom
}
