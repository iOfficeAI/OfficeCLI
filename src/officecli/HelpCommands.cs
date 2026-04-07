// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;

namespace OfficeCli;

/// <summary>
/// Format-specific help: docx, xlsx, pptx with nested verb help (view, get, set, add, etc.).
/// These are help-only commands — they do not execute any document operations.
/// Real format-prefixed commands are rewritten in Program.cs before this handler runs.
/// Args are intercepted before System.CommandLine parsing so --help works naturally.
/// </summary>
internal static class HelpCommands
{
    /// <summary>
    /// Register short-description commands so they appear in root --help listing.
    /// </summary>
    internal static void Register(RootCommand rootCommand)
    {
        rootCommand.Add(new Command("docx", "Word (.docx) help — run 'officecli docx [view|get|query|set|add|raw]' for details"));
        rootCommand.Add(new Command("xlsx", "Excel (.xlsx) help — run 'officecli xlsx [view|get|query|set|add|raw]' for details"));
        rootCommand.Add(new Command("pptx", "PowerPoint (.pptx) help — run 'officecli pptx [view|get|query|set|add|raw]' for details"));
    }

    /// <summary>
    /// Intercept args before System.CommandLine. Returns true if handled.
    /// Supports three-layer navigation:
    ///   officecli pptx set              → Layer 1: all elements overview
    ///   officecli pptx set shape        → Layer 2: element detail
    ///   officecli pptx set shape.fill   → Layer 3: property detail
    /// </summary>
    internal static bool TryHandle(string[] args)
    {
        if (args.Length == 0) return false;
        var format = args[0].ToLowerInvariant();
        if (format is not ("docx" or "xlsx" or "pptx")) return false;

        // Extract verb and optional element.property (skip --help flags)
        string? verb = null;
        string? elementArg = null;
        string? extraArg = null;
        foreach (var arg in args.Skip(1))
        {
            if (arg is "--help" or "-h" or "-?") continue;
            if (arg.StartsWith("--")) continue;
            if (verb == null) { verb = arg.ToLowerInvariant(); continue; }
            if (elementArg == null) { elementArg = arg.ToLowerInvariant(); continue; }
            extraArg = arg;
            break;
        }

        static bool LooksLikeOfficeFilePath(string arg)
        {
            if (string.IsNullOrWhiteSpace(arg) || arg.StartsWith("--"))
                return false;

            var ext = System.IO.Path.GetExtension(arg).ToLowerInvariant();
            return ext is ".doc" or ".docx" or ".xls" or ".xlsx" or ".ppt" or ".pptx";
        }

        if (verb is "add" or "set" or "get" or "query" or "remove" or "view" or "raw" or "raw-set")
        {
            if (elementArg != null && LooksLikeOfficeFilePath(elementArg))
                return false;

            // If extra args beyond element look like file paths or document paths,
            // this is a real command — don't intercept for help
            if (extraArg != null && (LooksLikeOfficeFilePath(extraArg) || extraArg.Contains('/') || extraArg.Contains('\\')))
                return false;
        }

        // Parse element.property syntax
        string? element = null;
        string? property = null;
        if (elementArg != null)
        {
            var dotIdx = elementArg.IndexOf('.');
            if (dotIdx >= 0)
            {
                element = elementArg[..dotIdx];
                property = elementArg[(dotIdx + 1)..];
            }
            else
            {
                element = elementArg;
            }
        }

        var help = GetHelp(format, verb, element, property);
        Console.WriteLine(help);
        return true;
    }

    static string GetHelp(string format, string? verb, string? element = null, string? property = null)
    {
        // Try wiki-based help first (from local directory or embedded resource)
        var wikiHelp = WikiHelpLoader.TryGetHelp(format, verb, element, property);
        if (wikiHelp != null) return wikiHelp;

        // Fallback to hardcoded help (no element/property layer support in fallback)
        return GetFallbackHelp(format, verb);
    }

    static string GetFallbackHelp(string format, string? verb) => format switch
    {
        "docx" => verb switch
        {
            "view" => DocxView,
            "get" => DocxGet,
            "query" => DocxQuery,
            "set" => DocxSet,
            "add" => DocxAdd,
            "raw" => DocxRaw,
            _ => DocxOverview,
        },
        "xlsx" => verb switch
        {
            "view" => XlsxView,
            "get" => XlsxGet,
            "query" => XlsxQuery,
            "set" => XlsxSet,
            "add" => XlsxAdd,
            "raw" => XlsxRaw,
            _ => XlsxOverview,
        },
        "pptx" => verb switch
        {
            "view" => PptxView,
            "get" => PptxGet,
            "query" => PptxQuery,
            "set" => PptxSet,
            "add" => PptxAdd,
            "raw" => PptxRaw,
            _ => PptxOverview,
        },
        _ => ""
    };

    // ======================== DOCX ========================

    const string DocxOverview = """
Word (.docx) Reference
======================

Path system (1-based):
  /                          Document root (page settings, default font)
  /body/p[N]                 Paragraph N
  /body/p[N]/r[M]            Run M in paragraph N
  /body/tbl[N]               Table N
  /body/tbl[N]/tr[R]/tc[C]   Cell at row R, column C
  /header[N]                 Header N
  /footer[N]                 Footer N
  /bookmark[Name]            Bookmark by name

Common workflow:
  1. officecli view doc.docx outline              # understand structure
  2. officecli view doc.docx text --max-lines 50  # read content
  3. officecli get  doc.docx '/body/p[1]' --depth 2  # inspect element
  4. officecli set  doc.docx '/body/p[1]/r[1]' --prop bold=true
  5. officecli validate doc.docx                  # verify changes

Run 'officecli docx <command>' for details:
  view    View modes and options
  get     DOM path navigation
  query   CSS-like selectors
  set     Property reference with examples
  add     Element types and properties
  raw     Raw XML parts reference
""";

    const string DocxView = """
Word (.docx) — view
====================

Modes:
  text (t)       Plain text with element paths (e.g. /body/p[1]). Use --start/--end/--max-lines to paginate.
  annotated (a)  Text with element paths and formatting details (font, size, bold, style, LaTeX formulas).
  outline (o)    Hierarchical structure by heading styles.
  stats (s)      Statistics: paragraph/table/image/equation counts, style inheritance.
  issues (i)     Formatting/content/structure problems. Filter with --type and --limit.

Options:
  --start N       Start at paragraph/line N
  --end N         End at paragraph/line N
  --max-lines N   Limit output lines (shows total count when truncated)
  --type T        Issue filter: format, content, structure (issues mode only)
  --limit N       Max number of issues to return
  --json          Output as JSON

Examples:
  officecli view doc.docx text --max-lines 100
  officecli view doc.docx annotated --start 10 --end 20
  officecli view doc.docx outline
  officecli view doc.docx stats --json
  officecli view doc.docx issues --type format --limit 5
""";

    const string DocxGet = """
Word (.docx) — get
===================

Get a document node by DOM path. Returns node type, properties, and children.

Standard paths:
  /                        Document root (core properties: title, author, ...)
  /body/p[3]               Paragraph 3
  /body/p[1]/r[1]          Run 1 (format: font, size, bold, italic, superscript, subscript)
  /body/tbl[1]/tr[1]/tc[1] Table cell
  /header[N]               Header N (type, text, font, size, bold, italic, color, alignment)
  /footer[N]               Footer N (same as header)
  /bookmark[Name]          Bookmark by name (text between start/end, id)
  /footnote[N]             Footnote N (N = id from add, returns text)
  /endnote[N]              Endnote N (N = id from add, returns text)
  /toc[N]                  TOC N (returns levels, hyperlinks, pageNumbers)
  /section[N]              Section N (returns type, pageWidth/Height, orientation, margins)
  /styles/StyleId          Style (returns font, size, bold, color, alignment, ...)

Also supports any XML path via element localName:
  /body/tbl[1]/tblPr       Table 1 properties
  /body/p[1]/r[1]/rPr      Run properties of first run

Options:
  --depth N   Depth of child nodes to include (default 1)
  --json      Output as JSON

Examples:
  officecli get doc.docx /                     # document root + core properties
  officecli get doc.docx '/body/p[1]' --depth 3
  officecli get doc.docx '/body/tbl[1]/tr[1]/tc[1]' --json
  officecli get doc.docx '/footnote[1]'        # footnote text
  officecli get doc.docx '/toc[1]'             # TOC field properties
  officecli get doc.docx '/section[1]'         # section page setup
  officecli get doc.docx '/styles/Heading1'    # style definition
""";

    const string DocxQuery = """
Word (.docx) — query
=====================

Element types:  paragraph (p), run (r), table (tbl), picture, equation, header, footer, bookmark, chart
Attribute filters:  [attr=value], [attr!=value], [attr~=text], [attr>=value], [attr<=value]
Pseudo-selectors:   :contains("text"), :empty, :no-alt, :has(formula)
Child combinator:   paragraph > run[bold=true]
Generic XML:        Falls back to any XML element name (e.g. wsp, srgbClr[val=0070C0])

Filter operators (apply to any Format property or text):
  =          Exact match              [style=Heading1]
  !=         Not equal                [font!=Arial]
  ~=         Contains text            [text~=季度报告]
  >=         Numeric greater/equal    [size>=24pt]
  <=         Numeric less/equal       [size<=12pt]
  Combined:  paragraph[style=Heading1][size>=14pt]

Examples:
  officecli query doc.docx 'paragraph[style=Heading1]'
  officecli query doc.docx 'run[bold=true]'
  officecli query doc.docx 'run[size>=14pt]'
  officecli query doc.docx 'paragraph[text~=报告]'
  officecli query doc.docx 'paragraph:contains("error")'
  officecli query doc.docx 'paragraph:empty'
  officecli query doc.docx 'picture:no-alt'
  officecli query doc.docx 'paragraph > run[font!=Arial]'
  officecli query doc.docx 'paragraph[alignment=center]'
  officecli query doc.docx 'header'
  officecli query doc.docx 'footer'
  officecli query doc.docx 'bookmark'
  officecli query doc.docx 'bookmark:contains("important")'
  officecli query doc.docx 'chart'
  officecli query doc.docx 'chart:contains("Sales")'
  officecli query doc.docx 'field'
  officecli query doc.docx 'field:contains("PAGE")'
  officecli query doc.docx 'field[fieldType=date]'
""";

    const string DocxSet = """
Word (.docx) — set
===================

Usage: officecli set <file> <path> --prop key=value [--prop key=value ...]

Run properties (/body/p[N]/r[M]):
  text, font, size, bold, italic, color, underline, strike, highlight,
  caps, smallCaps, superscript, subscript, dstrike, vanish, outline,
  shadow, emboss, imprint, noProof, rtl,
  shd (format: "fill" or "pattern;fill" or "pattern;fill;color")
  w14 text effects (Word 2010+, use ; as separator, - also accepted):
    textOutline   Text stroke: "WIDTHpt;COLOR" (e.g. "0.5pt;FF0000"), "none" to remove
    textFill      Text fill gradient: "C1;C2[;ANGLE]", "radial:C1;C2", or solid "COLOR"
    w14shadow     Advanced text shadow: "COLOR[;BLUR[;ANGLE[;DIST[;OPACITY]]]]"
    w14glow       Text glow: "COLOR[;RADIUS[;OPACITY]]"
    w14reflection Text reflection: "tight", "half", "full"
  For images in runs: alt, width, height (cm/in/pt/px or raw EMU),
    path/src (replace image file, removes old image part to avoid bloat)

Paragraph properties (/body/p[N]):
  style, alignment (left|center|right|justify),
  firstLineIndent, leftIndent, rightIndent, hangingIndent (twips),
  shd, spaceBefore, spaceAfter (unit-qualified: '12pt', '0.5cm', '0.5in', or bare twips),
  lineSpacing (unit-qualified: '1.5x' multiplier, '150%', '18pt' fixed, or bare twips),
  numId, numLevel/ilvl,
  listStyle (bullet|numbered|none), start (numbering start value),
  keepNext, keepLines, pageBreakBefore, widowControl (bool)

Table cell properties (/body/tbl[N]/tr[R]/tc[C]):
  text, font, size, bold, italic, color, shd, alignment,
  valign (top|center|bottom), width, vmerge (restart|continue), gridspan,
  padding (all sides, twips), padding.top, padding.bottom, padding.left, padding.right,
  textDirection (btlr|tbrl|lrtb|horizontal|vertical|vertical-rl), nowrap (bool),
  border.all, border.top, border.bottom, border.left, border.right
    Border format: style[;size[;color[;space]]]
    Styles: none,single,thick,double,dotted,dashed,dotDash,triple,wave,doubleWave,
            thinThickSmallGap,thickThinSmallGap,thinThickThinSmallGap,
            thinThickMediumGap,thickThinMediumGap,3dEmboss,3dEngrave

Table row properties (/body/tbl[N]/tr[R]):
  height (twips, at-least), height.exact (twips, exact), header (bool)

Table properties (/body/tbl[N]):
  alignment (left|center|right), width (twips or "100%"),
  indent (twips), cellSpacing (twips), layout (fixed|auto),
  padding (default cell margin, twips),
  border.all, border.top, border.bottom, border.left, border.right,
  border.insideH, border.insideV (same format as cell borders)

Document root (/):
  defaultFont, pageBackground, pageWidth, pageHeight,
  marginTop, marginBottom, marginLeft, marginRight,
  title, author, subject, keywords, description, category,
  lastModifiedBy, revision,
  find + replace (text search & replace across document),
    scope (all|body|headers|footers, default: all)

Batch Set (selector path):
  Use a query selector as path to set properties on all matching elements.
  E.g.: set doc.docx 'paragraph[style=Heading1]' --prop font=Arial

Footnote (/footnote[N]):
  text

Endnote (/endnote[N]):
  text

Field (/field[N]):
  instruction (field code, e.g. " PAGE ", " DATE \\@ \"yyyy-MM-dd\" ")
  text/result (cached display text), dirty (bool, force recalc on open)
  Get returns: instruction, fieldType, dirty, text (cached result)

TOC (/toc[N]):
  levels (e.g. "1-3"), hyperlinks (bool), pagenumbers (bool)

Section (/section[N]):
  type (nextPage|continuous|evenPage|oddPage),
  pagewidth, pageheight (twips), orientation (portrait|landscape),
  margintop, marginbottom, marginleft, marginright (twips),
  columns (equal-width: "3" or "3,720" for count,space),
  colWidths (custom: "3000,720,2000,720,3000" alternating width,space),
  separator (bool, draw line between columns)

SDT / Content Control (/body/sdt[N] or /body/p[N]/sdt[N]):
  alias (display name), tag (identifier), lock (unlocked|content|sdt|both),
  text (replace content text)

Header (/header[N]):
  text, font, size, bold, italic, color, alignment

Footer (/footer[N]):
  text, font, size, bold, italic, color, alignment

Bookmark (/bookmark[Name]):
  name (rename), text (replace content between start/end)

Style (/styles/StyleId):
  name, basedon, next, font, size, bold, italic, color,
  alignment, spacebefore, spaceafter (unit-qualified: '12pt', '0.5cm', or bare twips)

Watermark (/watermark):
  text, color, font, opacity, rotation

Any XML attribute is also settable via element path (use get --depth N to find paths).
Composite props (pBdr, tabs, lang, bdr) -> use raw-set instead.

Examples:
  officecli set doc.docx '/body/p[1]/r[1]' --prop bold=true --prop color=FF0000
  officecli set doc.docx '/body/p[2]' --prop style=Heading1 --prop alignment=center
  officecli set doc.docx '/body/tbl[1]/tr[1]/tc[1]' --prop text="Hello" --prop shd=4472C4
  officecli set doc.docx '/body/tbl[1]/tr[1]/tc[1]' --prop valign=center --prop gridspan=2
  officecli set doc.docx / --prop defaultFont=Arial --prop title="My Doc" --prop author="John"
  officecli set doc.docx / --prop find="甲方" --prop replace="乙方"
  officecli set doc.docx / --prop find="OldCo" --prop replace="NewCo" --prop scope=body
  officecli set doc.docx 'paragraph[style=Heading1]' --prop font=Arial --prop size=18
  officecli set doc.docx 'run[bold=true]' --prop color=FF0000
  officecli set doc.docx '/footnote[1]' --prop text="Updated footnote"
  officecli set doc.docx '/field[1]' --prop instruction=" DATE \\@ \"yyyy/MM/dd\" " --prop dirty=true
  officecli set doc.docx '/toc[1]' --prop levels="1-2" --prop pagenumbers=false
  officecli set doc.docx '/section[1]' --prop orientation=landscape --prop margintop=720
  officecli set doc.docx '/styles/Heading1' --prop font=Arial --prop size=16 --prop bold=true
  officecli set doc.docx '/header[1]' --prop text="New Header" --prop bold=true
  officecli set doc.docx '/footer[1]' --prop text="Page Footer" --prop alignment=center
  officecli set doc.docx '/bookmark[MyBookmark]' --prop text="Updated text"
  officecli set doc.docx '/chart[1]' --prop title="Revenue" --prop legend=top
  officecli set doc.docx '/chart[1]' --prop colors=FF0000,00FF00 --prop dataLabels=value

Chart (/chart[N]):
  title        Chart title text (or "none" to remove)
  legend       Legend position: top/bottom/left/right or "none" to remove
  categories   Update category labels (comma-separated)
  data         Replace all series: "S1:1,2;S2:3,4"
  series1..N   Update individual series: "NewName:1,2,3" or just "1,2,3"
  colors       Series colors (comma-separated hex): "FF0000,00FF00,0000FF"
  dataLabels   Data labels: value, category, series, percent, all, none
  labelPos     Data label position: center, insideEnd, insideBase, outsideEnd, top, bottom, left, right
  labelFont    Data label font: "size:color:bold" (e.g. "10:FF0000:true")
  axisTitle    Value axis title (alias: vtitle)
  catTitle     Category axis title (alias: htitle)
  axisMin, axisMax  Value axis scale bounds
  majorUnit, minorUnit  Tick mark spacing
  axisNumFmt   Value axis number format (e.g. "0.0", "$#,##0")
  gridlines    Major gridlines: true/none or "color:widthPt:dash" (e.g. "CCCCCC:0.5:dash")
  minorGridlines  Minor gridlines: true/none or "color:widthPt:dash"
  plotFill     Plot area background color (hex)
  chartFill    Chart area background color (hex)
  lineWidth    Series line width in pt (e.g. "2.5")
  lineDash     Series line dash: solid, dot, dash, dashdot, longdash
  marker       Series marker: "style:size:color" (e.g. "circle:8:FF0000"), styles: circle, diamond, square, triangle, star, x, plus, none
  markerSize   Marker size (2-72)
  style        Chart style ID (1-48, built-in Excel/PowerPoint styles)
  transparency Series fill transparency (0-100, %)
  gradient     Series gradient fill: "color1-color2:angle" (e.g. "FF0000-0000FF:90")
  gradients    Per-series gradients (semicolon-separated): "FF0000-0000FF;00FF00-FFFF00"
  secondary    Secondary axis: series indices (e.g. "2,3")

Document-level properties (via set / path):
  title, author, subject, keywords, description, category,
  lastModifiedBy, revision, defaultFont

DocDefaults (via set / path):
  docDefaults.font           Default font for all scripts
  docDefaults.font.latin     Default Latin font only
  docDefaults.font.eastAsia  Default East Asian font only
  docDefaults.font.cs        Default Complex Script font only
  docDefaults.fontSize       Default font size (e.g. "12pt")
  docDefaults.color          Default font color (hex)
  docDefaults.bold, docDefaults.italic
  docDefaults.alignment      Default paragraph alignment
  docDefaults.spaceBefore, docDefaults.spaceAfter, docDefaults.lineSpacing

DocGrid / CJK layout (via set / path):
  docGrid.type, docGrid.linePitch, docGrid.charSpace
  autoSpaceDE, autoSpaceDN, kinsoku, overflowPunct, charSpacingControl

Compatibility (via set / path):
  compatibility.mode, compatibility.preset (word2019/word2010/css-layout)
  compatibility.<flag>       Individual compat flag (e.g. compatibility.useFarEastLayout)

Settings flags (via set / path):
  embedFonts, mirrorMargins, evenAndOddHeaders, defaultTabStop, etc.

Theme (via set / — all formats):
  theme.color.<slot>         dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink
  theme.font.major.latin, theme.font.minor.latin, etc.

Extended properties (via set / — all formats):
  extended.template, extended.manager, extended.company
""";

    const string DocxAdd = """
Word (.docx) — add
===================

Usage: officecli add <file> <parent> --type <type> [--index N] [--prop key=value ...]
   or: officecli add <file> <parent> --from <path> [--index N]  (clone existing element)

Types and properties:

  paragraph (p)  -- parent: /body or /body/tbl[N]/tr[R]/tc[C]
    text, font, size, bold, italic, color, underline, strike, highlight,
    caps, smallCaps, superscript, subscript, style, alignment,
    firstLineIndent, leftIndent, rightIndent, hangingIndent,
    spaceBefore, spaceAfter (unit-qualified: '12pt', '0.5cm', or bare twips),
    lineSpacing (unit-qualified: '1.5x', '150%', '18pt', or bare twips),
    numId, numLevel, shd,
    listStyle, start, keepNext, keepLines, pageBreakBefore, widowControl

  run (r)  -- parent: /body/p[N]
    text, font, size, bold, italic, color, underline, strike, highlight,
    caps, smallCaps, superscript, subscript, shd

  table (tbl)  -- parent: /body
    rows (int), cols (int), colWidths (comma-separated twips, e.g. "3000,2000,5000"),
    alignment (left|center|right), width (twips or "100%"),
    indent (twips), cellSpacing (twips), layout (fixed|auto), padding (twips),
    border.* (same as Set)

  row (tr)  -- parent: /body/tbl[N]
    cols (int, default: match existing), height (twips),
    c1, c2, ... (cell text shortcuts)

  cell (tc)  -- parent: /body/tbl[N]/tr[M]
    text, width (twips)

  picture (image, img)  -- parent: /body/p[N] or /body
    path (required), width, height (cm/in/pt/px/EMU), alt
    Floating: anchor=true, wrap (none|square|tight|through|topAndBottom),
      hposition, vposition (cm/in/pt/EMU), hrelative (margin|page|column|character),
      vrelative (margin|page|paragraph|line), behindText (bool)

  chart  -- parent: /body
    chartType (column|bar|line|pie|doughnut|area|scatter|bubble|radar|stock|combo, default: column)
    title, categories ("Q1,Q2,Q3"), data ("S1:1,2,3;S2:4,5,6")
    series1..N ("Revenue:100,200"), colors ("FF0000,00FF00"), legend (top|bottom|left|right|none)
    width, height (cm/in/pt/EMU, default: 15cm x 10cm)

  equation (formula, math)  -- parent: /body/p[N] or /body
    formula (required, LaTeX subset), mode (display|inline)
    Supported: \frac{}{}, \sqrt{}, ^{}, _{}, \sum, \int, Greek letters

  hyperlink (link)  -- parent: /body/p[N]
    url (required), text (display text, defaults to url)
    font, size (optional run formatting)

  comment  -- parent: /body/p[N] or /body/p[N]/r[M]
    text (required), author, initials, date (ISO format)

  section (sectionbreak)  -- parent: /body
    type (nextPage|continuous|evenPage|oddPage, default: nextPage),
    pagewidth, pageheight (twips), orientation (portrait|landscape)

  footnote  -- parent: /body/p[N]
    text (required)

  endnote  -- parent: /body/p[N]
    text (required)

  toc (tableofcontents)  -- parent: /body
    levels (default "1-3"), title, hyperlinks (true|false),
    pagenumbers (true|false)

  header  -- parent: /
    text, type (default|first|even), font, size, bold, italic, color, alignment

  footer  -- parent: /
    text, type (default|first|even), font, size, bold, italic, color, alignment

  field (pagenum, pagenumber, numpages, date)  -- parent: /body/p[N] or /body
    instruction (field code, e.g. " PAGE ", " NUMPAGES ", " DATE \\@ \"yyyy-MM-dd\" ")
    text (placeholder value), font, size, bold, color, alignment (body-level)

  pagebreak (break)  -- parent: /body/p[N] or /body
    type (page|column|textwrapping, default: page)

  columnbreak  -- parent: /body/p[N] or /body
    (shorthand for break with type=column)

  sdt (contentcontrol)  -- parent: /body or /body/p[N]
    sdtType (text|richtext|dropdown|combobox|date, default: text)
    alias (display name), tag (identifier), lock (unlocked|content|sdt|both)
    text (content), items (comma-separated, for dropdown/combobox)
    format (date format, e.g. "yyyy-MM-dd", for date type)

  bookmark  -- parent: /body/p[N]
    name (required), text (optional, creates run between start/end)

  style  -- parent: /body (creates in styles part)
    name (required), id, type (paragraph|character|table),
    basedon, next, font, size, bold, italic, color,
    alignment, spacebefore, spaceafter (unit-qualified: '12pt', '0.5cm', or bare twips)

  watermark  -- parent: / (creates VML text watermark in headers)
    text (default: DRAFT), color (default: silver), font (default: Calibri),
    size (default: 1pt), opacity (default: .5), rotation (default: 315),
    width (default: 415pt), height (default: 207.5pt)
    Path: /watermark — use get, set, remove

--index is 0-based. If omitted, appends to end.
--from clones an element (cross-part relationships handled automatically).

Document properties (via set / path):
  title, author, subject, keywords, description, category,
  lastModifiedBy, revision
  defaultFont          Legacy alias for docDefaults.font

DocDefaults (via set / path):
  docDefaults.font           Default font for all scripts (Ascii, HighAnsi, EastAsia, ComplexScript)
  docDefaults.font.latin     Default Latin/Ascii font only
  docDefaults.font.eastAsia  Default East Asian font only
  docDefaults.font.cs        Default Complex Script font only
  docDefaults.fontSize       Default font size (e.g. "12pt")
  docDefaults.color          Default font color (hex)
  docDefaults.bold           Default bold (true/false)
  docDefaults.italic         Default italic (true/false)
  docDefaults.alignment      Default paragraph alignment (left/center/right/justify)
  docDefaults.spaceBefore    Default space before (e.g. "6pt")
  docDefaults.spaceAfter     Default space after (e.g. "6pt")
  docDefaults.lineSpacing    Default line spacing (e.g. "1.5x", "18pt")

DocGrid / CJK layout (via set / path):
  docGrid.type               Grid type: default, lines, linesAndChars, snapToChars
  docGrid.linePitch          Grid line pitch in twips
  docGrid.charSpace          Grid character spacing
  autoSpaceDE                Auto-space between East Asian and Latin (true/false)
  autoSpaceDN                Auto-space between East Asian and numbers (true/false)
  kinsoku                    Kinsoku line-breaking rules (true/false)
  overflowPunct              Overflow punctuation (true/false)
  charSpacingControl         compressPunctuation, compressPunctuationAndJapaneseKana, doNotCompress

Compatibility (via set / path):
  compatibility.mode         Compatibility mode version (e.g. 15 for Word 2013+)
  compatibility.preset       Preset: word2019, word2010, css-layout
  compatibility.<flag>       Any compat flag (e.g. compatibility.useFarEastLayout=true)

Settings flags (via set / path):
  embedFonts, embedSystemFonts, saveSubsetFonts
  mirrorMargins, gutterAtTop, bookFoldPrinting, bookFoldReversePrinting
  bookFoldPrintingSheets     Sheets per booklet
  evenAndOddHeaders          Different odd/even headers (true/false)
  defaultTabStop             Default tab stop in twips
  displayBackgroundShape, doNotDisplayPageBoundaries
  removePersonalInfo, removeDateAndTime

Theme (via set / path — all formats):
  theme.color.<slot>         Theme color (slot: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink)
  theme.font.major.latin     Major (heading) Latin font
  theme.font.major.eastAsia  Major East Asian font
  theme.font.minor.latin     Minor (body) Latin font
  theme.font.minor.eastAsia  Minor East Asian font

Extended properties (via set / path — all formats):
  extended.template          Template name
  extended.manager           Manager
  extended.company           Company

Examples:
  officecli add doc.docx /body --type paragraph --prop text="Hello World" --prop style=Heading1
  officecli add doc.docx '/body/p[1]' --type run --prop text="bold text" --prop bold=true
  officecli add doc.docx '/body/p[1]' --type run --prop text="2" --prop superscript=true
  officecli add doc.docx /body --type table --prop rows=3 --prop cols=4
  officecli add doc.docx '/body/tbl[1]' --type row --prop c1="Name" --prop c2="Value"
  officecli add doc.docx '/body/tbl[1]/tr[1]' --type cell --prop text="Extra"
  officecli add doc.docx /body --type picture --prop path=logo.png --prop width=5cm
  officecli add doc.docx /body --type picture --prop path=bg.png --prop anchor=true --prop wrap=square --prop hposition=2cm --prop vposition=3cm
  officecli add doc.docx '/body/p[1]' --type equation --prop formula="\frac{a}{b}"
  officecli add doc.docx '/body/p[3]' --type comment --prop text="Please review"
  officecli add doc.docx /body --type section --prop type=nextPage
  officecli add doc.docx '/body/p[1]' --type footnote --prop text="See reference 1"
  officecli add doc.docx '/body/p[1]' --type endnote --prop text="Additional info"
  officecli add doc.docx /body --type toc --prop levels="1-3" --prop title="Contents"
  officecli add doc.docx /body --type style --prop name=MyStyle --prop font=Arial --prop bold=true
  officecli set doc.docx / --prop title="My Doc" --prop author="John"
  officecli add doc.docx / --type header --prop text="My Header" --prop type=default --prop bold=true
  officecli add doc.docx / --type footer --prop text="Page Footer" --prop alignment=center
  officecli add doc.docx '/body/p[1]' --type bookmark --prop name=MyBookmark --prop text="marked text"
  officecli add doc.docx /body --from '/body/p[1]' --index 5
  officecli add doc.docx '/footer[1]/p[1]' --type pagenum
  officecli add doc.docx '/body/p[5]' --type pagebreak
  officecli add doc.docx '/body/p[5]' --type columnbreak
  officecli add doc.docx /body --type field --prop instruction=" NUMPAGES "
  officecli add doc.docx /body --type sdt --prop sdtType=dropdown --prop alias="Status" --prop items="Draft,Review,Final"
  officecli add doc.docx '/body/p[1]' --type sdt --prop sdtType=text --prop alias="Name" --prop text="Enter name"
  officecli set doc.docx '/section[1]' --prop columns=2 --prop separator=true
  officecli set doc.docx '/section[1]' --prop colWidths="3000,720,5000"
""";

    const string DocxRaw = """
Word (.docx) — raw
===================

Available parts:
  /document      Main document body (default)
  /styles        Style definitions
  /numbering     List/numbering definitions
  /settings      Document settings
  /header[N]     Header N (1-based)
  /footer[N]     Footer N (1-based)

raw-set actions: append, prepend, insertbefore, insertafter, replace, remove, setattr
No xmlns declarations needed -- prefixes auto-registered: w, a, r, mc, wp, wps, v, wp14

add-part types: chart, header, footer (returns relId for use with raw-set)

Examples:
  officecli raw doc.docx /document
  officecli raw doc.docx /styles
  officecli raw-set doc.docx /document --xpath "//w:body/w:p[1]" --action replace --xml '<w:p><w:r><w:t>New</w:t></w:r></w:p>'
  officecli raw-set doc.docx /styles --xpath "//w:style[@w:styleId='Heading1']/w:rPr/w:color" --action setattr --xml "w:val=FF0000"
  officecli add-part doc.docx / --type header
""";

    // ======================== XLSX ========================

    const string XlsxOverview = """
Excel (.xlsx) Reference
=======================

Path system (1-based for rows, standard cell refs):
  /                    Workbook root
  /Sheet1              Sheet by name
  /Sheet1/A1           Cell A1 in Sheet1
  /Sheet1/row[N]       Row N
  /Sheet1/col[A]       Column A
  /Sheet1/chart[N]     Chart N
  /Sheet1/picture[N]   Picture N
  /Sheet1/table[N]     Table (ListObject) N
  /Sheet1/comment[N]   Comment/note N
  /Sheet1/validation[N] Data validation N
  /namedrange[N]       Named range by index
  /namedrange[Name]    Named range by name

Common workflow:
  1. officecli view data.xlsx outline              # see sheets and dimensions
  2. officecli view data.xlsx text --max-lines 50 --cols A,B,C
  3. officecli get  data.xlsx '/Sheet1' --depth 1
  4. officecli set  data.xlsx '/Sheet1/A1' --prop value=100 --prop fill=4472C4
  5. officecli validate data.xlsx

Run 'officecli xlsx <command>' for details:
  view    View modes and column filtering
  get     Cell and sheet navigation
  query   Cell selectors and filters
  set     Cell property reference
  add     Sheets, rows, cells, data bars
  raw     Raw XML parts reference
""";

    const string XlsxView = """
Excel (.xlsx) — view
=====================

Modes:
  text (t)       Tab-separated cell values with row numbers. Paginate with --start/--end/--max-lines.
  annotated (a)  Cell ref, value, and annotation (formula, type). Shows errors and empty cells.
  outline (o)    Sheet names with row/column counts and formula counts.
  stats (s)      Total/empty/formula/error cell counts, data type distribution.
  issues (i)     Formula errors (#REF!, #VALUE!, #NAME?, #DIV/0!).

Options:
  --start N       Start at row N
  --end N         End at row N
  --max-lines N   Limit output rows
  --cols A,B,C    Show only specific columns (comma-separated)
  --type T        Issue filter (issues mode only)
  --limit N       Max issues
  --json          Output as JSON

Examples:
  officecli view data.xlsx text --max-lines 20 --cols A,B,C
  officecli view data.xlsx annotated --start 1 --end 50
  officecli view data.xlsx outline
  officecli view data.xlsx stats
  officecli view data.xlsx issues --limit 10
""";

    const string XlsxGet = """
Excel (.xlsx) — get
====================

Paths:
  /              Workbook root (lists sheets)
  /Sheet1        Sheet (lists rows/cells summary)
  /Sheet1/A1     Specific cell
  /Sheet1/col[A]       Column (width, hidden)
  /Sheet1/row[N]       Row (height, hidden)
  /Sheet1/chart[N]     Chart (title, type, legend)
  /Sheet1/picture[N]   Picture (alt, name, position, size)
  /Sheet1/table[N]     Table (name, ref, style, columns)
  /Sheet1/comment[N]   Comment (ref, text, author)
  /Sheet1/validation[N] Data validation (sqref, type, formula1, ...)
  /Sheet1/cf[N]        Conditional formatting
  /Sheet1/autofilter   AutoFilter range
  /namedrange[N]       Named range by index or name

Options:
  --depth N   Depth of child nodes (default 1)
  --json      Output as JSON

Examples:
  officecli get data.xlsx / --depth 1
  officecli get data.xlsx /Sheet1 --depth 2
  officecli get data.xlsx '/Sheet1/A1' --json
""";

    const string XlsxQuery = """
Excel (.xlsx) — query
======================

Cell selectors:
  cell                          All cells
  Sheet1!cell                   Cells in specific sheet
  A                             All cells in column A
  cell[value='text']            Exact value match
  cell[value!='text']           Value not equal
  cell[formula=true]            Cells with formulas
  cell[formula=false]           Cells without formulas
  cell[type='String']           Filter by type (String, Number, Boolean)
  cell[empty=true]              Empty cells
  cell[empty=false]             Non-empty cells

Pseudo-selectors:
  :contains("text")             Contains text (case-insensitive)
  :empty                        Empty cells
  :has(formula)                 Cells with formulas

Filter operators (apply to any Format property or text):
  =          Exact match              [value=100]
  !=         Not equal                [type!=Number]
  ~=         Contains text            [text~=季度]
  >=         Numeric greater/equal    [font.size>=14pt]
  <=         Numeric less/equal       [font.size<=10pt]
  Combined:  cell[font.bold=true][font.size>=14pt]

Falls back to generic XML element navigation for advanced queries.

Examples:
  officecli query data.xlsx 'cell[formula=true]'
  officecli query data.xlsx 'Sheet1!cell[empty=false]'
  officecli query data.xlsx 'A'
  officecli query data.xlsx 'cell:contains("error")'
  officecli query data.xlsx 'cell[type=Number]'
  officecli query data.xlsx 'cell[text~=季度]'
  officecli query data.xlsx 'validation'
  officecli query data.xlsx 'comment'
  officecli query data.xlsx 'table'
  officecli query data.xlsx 'chart'
  officecli query data.xlsx 'chart:contains("Sales")'
""";

    const string XlsxSet = """
Excel (.xlsx) — set
====================

Usage: officecli set <file> '<path>' --prop key=value [--prop ...]

Cell properties (/SheetName/A1):
  value, formula, type (string|number|boolean), clear, link ("none" to remove)

Cell style properties (/SheetName/A1):
  font.bold, font.italic, font.strike, font.underline (true/false or single/double)
  font.color (hex), font.size (pt), font.name
  fill (hex RGB), numFmt (format string)
  alignment.horizontal (left|center|right|justify)
  alignment.vertical (top|center|bottom), alignment.wrapText (bool)
  rotation       Text rotation angle (0-180, or 255 for vertical)
  indent         Text indent level (0-15)
  shrinkToFit    Shrink text to fit cell (bool)
  border.all (thin|medium|thick|double|dashed|dotted|none)
  border.left, border.right, border.top, border.bottom (style)
  border.color (hex), border.left.color, ... (per-side color)

Merge/Unmerge (/SheetName/A1:D1):
  merge          true = merge range, false = unmerge

Column properties (/SheetName/col[A]):
  width          Column width (number), hidden (bool)
  outline        Outline/group level (0-7), collapsed (bool)

Row properties (/SheetName/row[1]):
  height         Row height in points, hidden (bool)
  outline        Outline/group level (0-7), collapsed (bool)

Sheet properties (/SheetName):
  name           Rename the sheet
  freeze         Freeze panes (e.g. "A2" = freeze row 1, "B2" = freeze row 1 + col A)
  zoom           Zoom scale percentage (10-400, default 100)
  tabColor       Sheet tab color (hex RGB, or "none" to remove)
  protect        true/false — protect the sheet (prevents editing locked cells)
  password       Protection password (use with protect=true)
  printArea      Print area range (e.g. "A1:F20")
  orientation    portrait / landscape
  paperSize      Paper size number (1=Letter, 9=A4)
  fitToPage      true/false — fit content to page when printing
  header         Print header text
  footer         Print footer text
  sort           Sort range and column (e.g. "A1:D10,B" or "A1:D10,B,desc")

AutoFilter (/SheetName/autofilter):
  range          Update filter range (e.g. A1:F100)

Data validation (/SheetName/validation[N]):
  sqref, type (list|whole|decimal|date|time|textLength|custom),
  operator (between|equal|greaterThan|...), formula1, formula2,
  allowBlank, showError, errorTitle, error, showInput, promptTitle, prompt

Shape (/SheetName/shape[N]):
  text           Replace text content (supports \n)
  font           Font typeface
  size           Font size in points
  bold           true/false
  italic         true/false
  color          Hex RGB text color (e.g. FF0000)
  fill           Hex RGB shape fill or "none"
  align          Text alignment: left/center/right
  name           Shape name
  shadow         Shadow: "COLOR-BLUR-ANGLE-DIST-OPACITY" or "none"
  glow           Glow: "COLOR-RADIUS-OPACITY" or "none"
  reflection     Reflection: tight/half/full or "none"
  softEdge       Soft edge radius in pt or "none"
  x, y, width, height  Position (col/row offset) and size (col/row span)

Picture (/SheetName/picture[N]):
  x, y (col/row offset), width, height (col/row span), alt
  rotation       Rotation angle in degrees (e.g. "45", "90")
  shadow         Shadow: "COLOR-BLUR-DIST-DIR" or "none"
  glow           Glow: "COLOR-RADIUS" or "none"
  reflection     Reflection: tight/half/full or "none"
  softEdge       Soft edge radius in pt or "none"

Table (/SheetName/table[N]):
  name, displayName, style, ref

Comment (/SheetName/comment[N]):
  text, author

Named range (/namedrange[N] or /namedrange[Name]):
  ref, name, comment

Chart (/SheetName/chart[N]):
  title        Chart title text (or "none" to remove)
  title.font   Title font typeface
  title.size   Title font size in pt
  title.color  Title font color (hex)
  title.bold   Title bold (true/false)
  title.glow   Title glow: "COLOR-RADIUS-OPACITY" (e.g. "00D2FF-6-60")
  title.shadow Title shadow: "COLOR-BLUR-ANGLE-DIST-OPACITY"
  legend       Legend position: top/bottom/left/right or "none"
  legendFont   Legend font: "size:color:fontname" (e.g. "9:8B949E:Helvetica Neue")
  axisFont     Axis label font: "size:color:fontname" (e.g. "10:58626E:Arial")
  categories   Update category labels (comma-separated)
  data         Replace all series: "S1:1,2;S2:3,4"
  series1..N   Update individual series: "NewName:1,2,3" or just "1,2,3"
  colors       Series colors (comma-separated hex): "FF0000,00FF00,0000FF"
  dataLabels   Data labels: value, category, series, percent, all, none
  labelPos     Data label position: center, insideEnd, insideBase, outsideEnd, top, bottom, left, right
  labelFont    Data label font: "size:color:bold" (e.g. "10:FF0000:true")
  axisTitle    Value axis title (alias: vtitle)
  catTitle     Category axis title (alias: htitle)
  axisMin, axisMax  Value axis scale bounds
  majorUnit, minorUnit  Tick mark spacing
  axisNumFmt   Value axis number format (e.g. "0.0", "$#,##0")
  gridlines    Major gridlines: true/none or "color:widthPt:dash"
  minorGridlines  Minor gridlines: true/none or "color:widthPt:dash"
  plotFill     Plot area fill: hex solid or gradient "C1-C2[:angle]"
  chartFill    Chart area fill: hex solid or gradient "C1-C2[:angle]"
  lineWidth    Series line width in pt (e.g. "2.5")
  lineDash     Series line dash: solid, dot, dash, dashdot, longdash
  marker       Series marker: "style:size:color" (e.g. "circle:8:FF0000")
  markerSize   Marker size (2-72)
  style        Chart style ID (1-48)
  opacity      Series fill opacity (0.0-1.0)
  transparency Series fill transparency (0-100, %)
  gradient     Series gradient: "C1-C2[:angle]" (all series)
  gradients    Per-series gradients (semicolon-separated): "C1-C2;C3-C4"
  series.shadow  Series bar shadow: "COLOR-BLUR-ANGLE-DIST-OPACITY" or "none"
  series.outline Series bar outline: "COLOR-WIDTH" (e.g. "FFFFFF-0.5") or "none"
  gap          Bar gap width (0-500, default 150)
  overlap      Bar overlap (-100 to 100)
  view3d       3D camera: "rotX,rotY,perspective" (e.g. "15,20,30")
  areafill     Area chart series fill: gradient "C1-C2[:angle]"
  secondary    Secondary axis: series indices (e.g. "2,3")

PivotTable (/SheetName/pivottable[N]):
  name         Pivot table name
  style        Style name (e.g. "PivotStyleMedium9")

Workbook properties (via set / path):
  workbook.date1904          Use 1904 date system (true/false)
  workbook.codeName          VBA code name
  workbook.filterPrivacy     Filter personal info on save (true/false)
  workbook.showObjects       Object display: all, placeholders, none
  workbook.backupFile        Create backup on save (true/false)
  workbook.dateCompatibility Date compatibility mode (true/false)

Calculation properties (via set / path):
  calc.mode              Calculation mode: auto, manual, autoExceptTables
  calc.iterate           Enable iterative calculation (true/false)
  calc.iterateCount      Max iterations (number)
  calc.iterateDelta      Max change threshold (number)
  calc.fullPrecision     Full precision calculation (true/false)
  calc.fullCalcOnLoad    Force full recalc on open (true/false)
  calc.refMode           Reference mode: A1, R1C1

Workbook protection (via set / path):
  workbook.protection    Enable/disable workbook protection (true/false/none)
  workbook.lockStructure Lock workbook structure (true/false)

Examples:
  officecli set data.xlsx '/Sheet1/A1' --prop value=100 --prop font.bold=true
  officecli set data.xlsx '/Sheet1/A1' --prop border.all=thin --prop border.color=000000
  officecli set data.xlsx '/Sheet1/A1:D1' --prop merge=true
  officecli set data.xlsx '/Sheet1/col[A]' --prop width=20
  officecli set data.xlsx '/Sheet1/row[1]' --prop height=30
  officecli set data.xlsx /Sheet1 --prop freeze=A2
  officecli set data.xlsx '/Sheet1/autofilter' --prop range=A1:F100
  officecli set data.xlsx '/Sheet1/validation[1]' --prop formula1="Yes,No,Maybe"
  officecli set data.xlsx '/Sheet1/picture[1]' --prop x=5 --prop y=3
  officecli set data.xlsx '/Sheet1/table[1]' --prop style=TableStyleLight1
  officecli set data.xlsx '/Sheet1/comment[1]' --prop text="Updated note"
  officecli set data.xlsx '/namedrange[MyRange]' --prop ref="Sheet1!$A$1:$E$20"
""";

    const string XlsxAdd = """
Excel (.xlsx) — add
=====================

Types and properties:

  sheet  -- parent: /
    name (default: Sheet{N})

  row  -- parent: /SheetName
    cols (int, number of empty cells)

  cell  -- parent: /SheetName
    ref (e.g. A1), value, formula, type,
    plus all style properties from 'set' (font.*, fill, alignment.*, numFmt)

  databar (conditionalformatting)  -- parent: /SheetName
    sqref (e.g. A1:A10), min, max, color (hex)

  colorscale  -- parent: /SheetName
    sqref (e.g. A1:A10), mincolor (hex, default F8696B),
    maxcolor (hex, default 63BE7B), midcolor (hex, optional for 3-color)

  iconset  -- parent: /SheetName
    sqref (e.g. A1:A10), iconset (default 3TrafficLights1),
    reverse (bool), showvalue (bool, default true)
    Icon sets: 3Arrows, 3ArrowsGray, 3Flags, 3TrafficLights1, 3TrafficLights2,
      3Signs, 3Symbols, 3Symbols2, 4Arrows, 4ArrowsGray, 4Rating, 4RedToBlack,
      4TrafficLights, 5Arrows, 5ArrowsGray, 5Rating, 5Quarters

  formulacf  -- parent: /SheetName
    sqref (e.g. A1:A10), formula (e.g. $A1>100),
    fill (hex), font.color (hex), font.bold (bool)

  validation (datavalidation)  -- parent: /SheetName
    sqref (required), type (list|whole|decimal|date|time|textLength|custom),
    formula1, formula2, operator (between|equal|greaterThan|...),
    allowBlank (bool), showError (bool), errorTitle, error,
    showInput (bool), promptTitle, prompt

  shape (textbox)  -- parent: /SheetName
    text (supports \n), name, font, size, bold, italic, color,
    fill (hex or "none"), line (hex or "none"),
    align (left/center/right), preset (roundRect, ellipse, etc.),
    shadow, glow, reflection, softEdge (same format as PPTX),
    x, y (col/row offset), width, height (col/row span)

  picture (image)  -- parent: /SheetName
    path (required), x, y (col/row offset, default 0),
    width, height (col/row span, default 5), alt

  table (listobject)  -- parent: /SheetName
    ref (required, e.g. A1:D10), name, displayName,
    style (default TableStyleMedium2), headerRow (bool), totalRow (bool),
    columns (comma-separated, auto-detected from header if omitted)

  comment (note)  -- parent: /SheetName
    ref (required, e.g. A1), text, author (default "Author")

  namedrange (definedname)  -- parent: /
    name (required), ref (e.g. Sheet1!$A$1:$D$10), scope (sheet name), comment

  chart  -- parent: /SheetName
    chartType (column|bar|line|pie|doughnut|area|scatter|bubble|radar|stock|combo|column3d|bar3d)
    title, categories (comma-separated), legend (top|bottom|left|right|none)
    data ("Series1:1,2,3;Series2:4,5,6") or series1/series2/... ("Name:1,2,3")
    x, y (col/row offset), width, height (col/row span)

  pivottable (pivot)  -- parent: /SheetName
    source (required): "Sheet1!A1:D100" or "A1:D100" (same sheet)
    position: anchor cell for pivot table (e.g. "F1"), auto-placed if omitted
    rows: row fields (comma-separated header names or 0-based indices)
    cols: column fields
    values: data fields with aggregation ("Sales:sum,Qty:count")
      Functions: sum, count, average, max, min, product, stddev, var
    filters: page/filter fields
    name: pivot table name (auto-generated if omitted)
    style: style name (default: PivotStyleLight16)

--index is 0-based. --from clones an existing element.

Examples:
  officecli add data.xlsx / --type sheet --prop name=Summary
  officecli add data.xlsx /Sheet1 --type row --prop cols=5
  officecli add data.xlsx /Sheet1 --type cell --prop ref=A1 --prop value=100 --prop fill=4472C4
  officecli add data.xlsx /Sheet1 --type databar --prop sqref=B2:B20 --prop color=63C384
  officecli add data.xlsx /Sheet1 --type colorscale --prop sqref=A1:A20 --prop mincolor=F8696B --prop maxcolor=63BE7B
  officecli add data.xlsx /Sheet1 --type iconset --prop sqref=B1:B10 --prop iconset=3Arrows
  officecli add data.xlsx /Sheet1 --type formulacf --prop sqref=A1:A10 --prop formula="$A1>100" --prop fill=FF0000
  officecli add data.xlsx /Sheet1 --type chart --prop chartType=column --prop title="Sales" --prop data="Q1:100,200;Q2:150,250" --prop categories="Jan,Feb"
  officecli add data.xlsx /Sheet1 --type validation --prop sqref=A1:A100 --prop type=list --prop formula1="Yes,No,Maybe"
  officecli add data.xlsx /Sheet1 --type picture --prop path=logo.png --prop x=1 --prop y=1 --prop width=4 --prop height=3
  officecli add data.xlsx /Sheet1 --type table --prop ref=A1:D10 --prop name=SalesData --prop style=TableStyleMedium2
  officecli add data.xlsx /Sheet1 --type comment --prop ref=A1 --prop text="Check this value" --prop author=John
  officecli add data.xlsx / --type namedrange --prop name=MyRange --prop ref="Sheet1!$A$1:$D$10"
  officecli set data.xlsx '/Sheet1/A1' --prop link="https://example.com"
""";

    const string XlsxRaw = """
Excel (.xlsx) — raw
=====================

Available parts:
  / or /workbook           Workbook XML root
  /styles                  Stylesheet (fonts, fills, number formats)
  /sharedstrings           Shared string table
  /SheetName               Worksheet XML
  /SheetName/drawing       Drawing/chart container
  /SheetName/chart[N]      Specific chart (1-based)

Row filtering (worksheet parts only):
  --start N   Start row
  --end N     End row
  --cols A,B  Column filter

raw-set actions: append, prepend, insertbefore, insertafter, replace, remove, setattr
No xmlns needed -- prefixes auto-registered: x, a, r, c, xdr

add-part types: chart (returns relId for use with raw-set)

Excel chart workflow:
  1. officecli add-part data.xlsx /Sheet1 --type chart   # get relId
  2. officecli raw-set data.xlsx /Sheet1/chart[1] ...    # fill chart XML
  3. officecli raw-set data.xlsx /Sheet1/drawing ...     # add anchor

Examples:
  officecli raw data.xlsx /Sheet1 --start 1 --end 100 --cols A,B
  officecli raw data.xlsx /styles
  officecli raw data.xlsx /sharedstrings
""";

    // ======================== PPTX ========================

    const string PptxOverview = """
PowerPoint (.pptx) Reference
=============================

Path system (1-based):
  /                          Presentation root
  /slide[N]                  Slide N
  /slide[N]/notes            Speaker notes for slide N
  /slide[N]/shape[M]         Shape M on slide N
  /slide[N]/picture[M]       Picture M on slide N
  /slide[N]/table[M]          Table M on slide N
  /slide[N]/table[M]/tr[R]/tc[C]  Table cell
  /slide[N]/placeholder[M]   Placeholder M (by ordinal or type name)
  /slide[N]/shape[M]/run[K]               Run K (flat index across paragraphs)
  /slide[N]/shape[M]/paragraph[P]        Paragraph P in shape
  /slide[N]/shape[M]/paragraph[P]/run[K]  Run K in paragraph P

XML paths also work (via get --depth N):
  /slide[1]/cSld/spTree/sp[1]/spPr/xfrm[1]/off[1]   Shape offset
  /slide[1]/cSld/spTree/sp[1]/txBody/p[1]/r[1]/rPr[1]  Run properties

Create a blank presentation:
  officecli create pres.pptx
  Output includes slide dimensions (slideWidth, slideHeight) for positioning:
    Created: pres.pptx
      slideWidth: 33.87cm
      slideHeight: 19.05cm

Common workflow:
  1. officecli create pres.pptx                     # create blank (shows slide size)
  2. officecli view pres.pptx outline               # slide titles overview
  3. officecli view pres.pptx annotated             # shapes, fonts, sizes
  4. officecli get  pres.pptx '/slide[1]' --depth 2
  5. officecli set  pres.pptx '/slide[1]/shape[1]' --prop text="Title"
  6. officecli validate pres.pptx

Run 'officecli pptx <command>' for details:
  view    View modes
  get     Slide/shape/table/placeholder navigation
  query   Shape selectors
  set     Shape/table/placeholder property reference
  add     Slides, shapes, tables, pictures, equations
  raw     Raw XML parts reference
""";

    const string PptxView = """
PowerPoint (.pptx) — view
==========================

Modes:
  text (t)       Plain text from all shapes with slide paths (e.g. /slide[1]).
  annotated (a)  Shape details with paths: type, text, font, size, pictures with alt text, equations.
  outline (o)    Slide-by-slide outline with titles and content summary.
  stats (s)      Slide/shape counts, text boxes, pictures, missing titles, font usage.
  issues (i)     Missing titles, inconsistent fonts, missing alt text on pictures.

Options:
  --start N       Start at slide N
  --end N         End at slide N
  --max-lines N   Limit output slides
  --type T        Issue filter: format, content, structure
  --limit N       Max issues
  --json          Output as JSON

Examples:
  officecli view pres.pptx outline
  officecli view pres.pptx annotated --start 1 --end 5
  officecli view pres.pptx text --max-lines 10
  officecli view pres.pptx stats
  officecli view pres.pptx issues --type format
""";

    const string PptxGet = """
PowerPoint (.pptx) — get
=========================

Paths:
  /                                          Presentation root (lists slides)
  /slide[1]                                  Slide 1 (lists shapes, tables, placeholders)
  /slide[1]/shape[1]                         Shape or text box
  /slide[1]/table[1]                         Table
  /slide[1]/table[1]/tr[1]                   Table row
  /slide[1]/table[1]/tr[1]/tc[1]             Table cell
  /slide[1]/placeholder[1]                   Placeholder by ordinal
  /slide[1]/placeholder[title]               Placeholder by type name
  /slide[1]/notes                            Speaker notes (text)
  /slide[1]/shape[1]/paragraph[1]            Paragraph in shape
  /slide[1]/shape[1]/paragraph[1]/run[1]     Run in paragraph
  /slide[1]/shape[1]/run[1]                  Run shortcut (flat index across paragraphs)
  /slide[1]/cSld/spTree/sp[1]/spPr           Shape properties XML element

Format keys returned by Get:

  Slide (/slide[N]):
    layout        Layout name (e.g. "Blank", "Title Slide", "Title and Content")
    layoutType    Layout type (e.g. "blank", "title", "obj")
    background    Solid hex, gradient (C1-C2[-angle]), or "image"
    transition    Transition type name (fade, wipe, push, etc.)
    advanceTime   Auto-advance time in ms (if set)
    advanceClick  false if click-advance is disabled

  Shape/textbox (/slide[N]/shape[M]):
    text, name, type (textbox/title)
    x, y, width, height    Position and size (e.g. "2cm")
    font, size, bold, italic
    underline              Underline style (sng, dbl, heavy, dotted, dash, wavy)
    strikethrough          Strike style (sngStrike, dblStrike)
    color                  Text color hex (from first run)
    fill                   Shape fill hex or "none"
    opacity                Fill opacity 0.0–1.0 (if Alpha set)
    gradient               Linear "C1-C2[-angle]", radial "radial:C1-C2[-focus]" (focus: tl/tr/bl/br/center)
    image                  Shape image fill (path to image file)
    line, lineWidth, lineDash, lineOpacity (0.0–1.0)
    preset / geometry       Shape geometry name (e.g. roundRect, ellipse, diamond)
    align, valign
    lineSpacing            Unit-qualified: "1.5x" (multiplier) or "18pt" (fixed)
    spaceBefore, spaceAfter  Unit-qualified: "12pt" from first paragraph
    margin                 Text padding
    rotation               Degrees
    autoFit                normal / shape / none
    list                   Bullet char or auto-number type
    link                   Hyperlink URL (from first run)
    shadow, glow           Effect color hex
    reflection             "true" if reflection effect applied
    animation              Flexible: "effect[-class][-direction][-duration][-trigger]" (e.g. "fly-left-400")
                           Trigger default: 1st on slide = click, subsequent = after (sequential)

  Chart (/slide[N]/chart[M]):
    chartType              column, bar, line, pie, doughnut, area, scatter, bubble, radar, stock
    title                  Chart title text
    legend                 Legend position (t/b/l/r)
    seriesCount            Number of data series
    categories             Comma-separated category labels
    x, y, width, height    Position and size
    (depth>0: children = series nodes with name + values)

  Video/Audio (/slide[N]/picture[M] with type=video or type=audio):
    name, volume, autoplay, trimStart, trimEnd
    x, y, width, height    Position and size

  Table cell (/slide[N]/table[M]/tr[R]/tc[C]):
    text, fill, font, size, bold, italic, color

  Run (/slide[N]/shape[M]/…/run[K]):
    text, font, size, bold, italic, color, link

Use --depth N to explore deeper. Any XML localName works as path segment.

Examples:
  officecli get pres.pptx / --depth 1
  officecli get pres.pptx '/slide[1]' --depth 2
  officecli get pres.pptx '/slide[1]/shape[1]' --depth 3 --json
  officecli get pres.pptx '/slide[1]/table[1]' --depth 2
  officecli get pres.pptx '/slide[1]/chart[1]' --depth 1
  officecli get pres.pptx '/slide[1]/placeholder[title]'
  officecli get pres.pptx '/slide[1]/notes'
  officecli get pres.pptx '/slide[1]/shape[1]/paragraph[1]/run[1]'
  officecli get pres.pptx '/slide[1]/cSld/spTree/sp[1]/spPr' --depth 3
""";

    const string PptxQuery = """
PowerPoint (.pptx) — query
============================

Element types:
  shape (textbox)    Text shapes / text boxes
  title              Title shapes
  picture (pic)      Images
  equation (math, formula)  Mathematical equations
  table              Tables
  placeholder        Placeholder shapes (shows phType)
  notes              Slides with speaker notes

Filters:
  [font="Arial"]     Shapes with specific font
  [font!="Arial"]    Shapes without specific font
  [title=true]       Title shapes only
  :contains("text")  Shapes/tables containing text (case-insensitive)
  :no-alt            Pictures without alt text

Filter operators (apply to any Format property or text):
  =          Exact match              [fill=FF0000]
  !=         Not equal                [type!=placeholder]
  ~=         Contains text            [text~=季度报告]
  >=         Numeric greater/equal    [size>=24pt]
  <=         Numeric less/equal       [size<=12pt]
  Combined:  shape[fill=FF0000][size>=24pt]

Scope by slide:
  slide[N] shape     Only shapes in slide N

Falls back to generic XML element name for advanced queries.

Examples:
  officecli query pres.pptx 'shape'
  officecli query pres.pptx 'slide[1] shape'
  officecli query pres.pptx 'title'
  officecli query pres.pptx 'title[fill=FF0000][size>=24pt]'
  officecli query pres.pptx 'shape[text~=季度报告]'
  officecli query pres.pptx 'shape[bold=true][fill=FF0000]'
  officecli query pres.pptx 'table'
  officecli query pres.pptx 'table:contains("revenue")'
  officecli query pres.pptx 'placeholder'
  officecli query pres.pptx 'picture:no-alt'
  officecli query pres.pptx 'shape:contains("hello")'
  officecli query pres.pptx 'equation'
""";

    const string PptxSet = """
PowerPoint (.pptx) — set
==========================

Usage: officecli set <file> <path> --prop key=value [--prop ...]

Colors: hex RGB (e.g. FF0000) or theme color names:
  accent1..accent6, dk1, dk2, lt1, lt2, tx1, tx2, bg1, bg2, hyperlink, followedhyperlink

Shape properties (/slide[N]/shape[M]) -- applies to all runs:
  text       Replace all text content (supports \n for line breaks, preserves first run's formatting)
  font       Font typeface
  size       Font size in points
  bold       true/false
  italic     true/false
  underline  true/single/double/heavy/dotted/dash/wavy/false
  strikethrough  true/single/double/false (alias: strike)
  color      Hex RGB text color (e.g. FF0000)
  fill       Hex RGB shape fill (e.g. 4472C4) or "none"
  line       Hex RGB border color (e.g. FF0000) or "none" (alias: linecolor, line.color)
  lineWidth  Border width (EMU or cm/pt, e.g. 2pt) (alias: line.width)
  lineDash   Border dash style: solid/dot/dash/dashdot/longdash (alias: line.dash)
  lineOpacity  Border opacity 0.0-1.0 (alias: line.opacity)
  preset     Shape geometry (e.g. roundRect, ellipse, rightArrow, diamond, star5). Also accepts: geometry=roundRect, shape=roundRect
  geometry   Custom freeform path: SVG-like syntax "M x,y L x,y C x1,y1 x2,y2 x,y Z"
             M=moveTo, L=lineTo, C=cubicBezier, Q=quadBezier, Z=close
             Example: "M 0,100 L 50,0 L 100,100 Z" (triangle in 100x100 space)
  textFill   Text gradient fill (same format as gradient): C1-C2[-angle], radial:C1-C2[-focus]
             Example: "FF0000-0000FF-90" (red to blue horizontal gradient on text)
  margin     Text padding inside shape (e.g. 0.5cm or left,top,right,bottom: 0.5cm,0.3cm,0.5cm,0.3cm)
  align      Text horizontal alignment: left (l), center (c), right (r), justify (j) — applies to all paragraphs
  valign     Text vertical alignment: top (t), center/middle (c/m), bottom (b)
  lineSpacing  Unit-qualified: '1.5x' (multiplier), '150%', '18pt' (fixed), or bare multiplier
  spaceBefore  Unit-qualified: '12pt', '0.5cm', '0.5in', or bare points
  spaceAfter   Unit-qualified: '6pt', '0.5cm', '0.5in', or bare points
  gradient   Linear: C1-C2[-angle], Radial: radial:C1-C2[-focus] (focus: tl/tr/bl/br/center)
  image      Shape image fill (path to image file, e.g. /tmp/bg.png)
  list       List style: bullet/numbered/alpha/roman/none or a custom character (e.g. ✓)
  spacing    Character spacing in points (e.g. 2 for +2pt, -1 for condensed) (alias: charSpacing, letterSpacing)
  indent     Paragraph first-line indent (EMU or cm/pt, negative for hanging indent)
  marginLeft Paragraph left margin (EMU or cm/pt) (alias: marL)
  marginRight Paragraph right margin (EMU or cm/pt) (alias: marR)
  rotation   Rotation angle in degrees (e.g. 45) (alias: rotate)
  opacity    Fill opacity 0.0-1.0 (e.g. 0.5 for 50%)
  textWarp   WordArt text effect (alias: wordart): textWave1, textChevron, textArchUp, etc. or "none"
  autoFit    Text auto-fit: true/normal, shape, false/none
  x          Horizontal position (EMU or cm/in/pt/px, e.g. 2cm)
  y          Vertical position (EMU or cm/in/pt/px, e.g. 3cm)
  width      Shape width (EMU or cm/in/pt/px, e.g. 10cm)
  height     Shape height (EMU or cm/in/pt/px, e.g. 2cm)
  flipH      Horizontal flip (true/false)
  flipV      Vertical flip (true/false)

3D effects (/slide[N]/shape[M]):
  rot3d      3D rotation: "rotX,rotY,rotZ" in degrees (e.g. "45,30,0") or "none" (alias: rotation3d)
  rotX       X-axis rotation in degrees
  rotY       Y-axis rotation in degrees
  rotZ       Z-axis rotation in degrees
  bevel      Top bevel: preset or preset-width-height in pt (e.g. "circle", "circle-6-6") (alias: bevelTop)
  bevelBottom  Bottom bevel (same format as bevel)
             Presets: circle, relaxedInset, cross, coolSlant, angle, softRound, convex, slope, divot, riblet, hardEdge, artDeco
  depth      3D extrusion depth in points (e.g. "10") or "none" (alias: extrusion)
  material   3D material: plastic, metal, warmMatte, matte, flat, clear, softMetal, powder, translucentPowder, darkEdge
  lighting   Light rig: threePt, balanced, soft, harsh, flood, contrasting, morning, sunrise, sunset, flat, glow, brightRoom (alias: lightRig)
  softEdge   Soft edge blur radius in points (e.g. "5") or "none"

Text effects (auto-applied to text runs when fill=none):
  shadow     "COLOR-BLUR-ANGLE-DIST-OPACITY" (e.g. "FFFFFF-6-135-4-60") or "none"
  glow       "COLOR-RADIUS-OPACITY" (e.g. "FF0000-8-75") or "none"
  reflection tight/half/full or "none"
  softEdge   Radius in pt (e.g. "5") or "none"
  Note: when shape has fill=none, these effects are applied to a:rPr (text runs)
  instead of spPr (shape), so they render on the text itself.

Chart properties (/slide[N]/chart[M]):
  title        Chart title text (or "none" to remove)
  title.font   Title font typeface
  title.size   Title font size in pt
  title.color  Title font color (hex)
  title.bold   Title bold (true/false)
  title.glow   Title glow: "COLOR-RADIUS-OPACITY"
  title.shadow Title shadow: "COLOR-BLUR-ANGLE-DIST-OPACITY"
  legend       Legend position: top/bottom/left/right or "none"
  legendFont   Legend font: "size:color:fontname" (e.g. "9:8B949E:Helvetica Neue")
  axisFont     Axis label font: "size:color:fontname"
  categories   Update category labels (comma-separated)
  data         Replace all series: "S1:1,2;S2:3,4"
  series1..N   Update individual series: "NewName:1,2,3" or just "1,2,3"
  colors       Series colors (comma-separated hex/theme): "FF0000,00FF00,accent3"
  dataLabels   Data labels: value, category, series, percent, all, none
  labelPos     Data label position: center, insideEnd, outsideEnd, top, bottom, left, right
  labelFont    Data label font: "size:color:bold" (e.g. "10:FF0000:true")
  axisTitle    Value axis title (alias: vtitle)
  catTitle     Category axis title (alias: htitle)
  axisMin, axisMax  Value axis scale bounds
  majorUnit, minorUnit  Tick mark spacing
  axisNumFmt   Value axis number format (e.g. "0.0", "$#,##0")
  gridlines    Major gridlines: true/none or "color:widthPt:dash"
  minorGridlines  Minor gridlines: true/none or "color:widthPt:dash"
  plotFill     Plot area fill: hex solid or gradient "C1-C2[:angle]"
  chartFill    Chart area fill: hex solid or gradient "C1-C2[:angle]"
  lineWidth    Series line width in pt (e.g. "2.5")
  lineDash     Series line dash: solid, dot, dash, dashdot, longdash
  marker       Series marker: "style:size:color" (e.g. "circle:8:FF0000")
  markerSize   Marker size (2-72)
  style        Chart style ID (1-48)
  opacity      Series fill opacity (0.0-1.0)
  transparency Series fill transparency (0-100, %)
  gradient     Series gradient: "C1-C2[:angle]" (all series)
  gradients    Per-series gradients (semicolon-separated)
  series.shadow  Series bar shadow: "COLOR-BLUR-ANGLE-DIST-OPACITY" or "none"
  series.outline Series bar outline: "COLOR-WIDTH" or "none"
  gap          Bar gap width (0-500)
  overlap      Bar overlap (-100 to 100)
  view3d       3D camera: "rotX,rotY,perspective"
  areafill     Area chart series fill: gradient "C1-C2[:angle]"
  secondary    Secondary axis: series indices (e.g. "2,3")
  x, y, width, height  Chart position and size (EMU or cm/in/pt/px)
  name         Chart name

Picture properties (/slide[N]/picture[M]):
  alt          Alternative text
  path         Replace image source (file path)
  crop         Crop (percentage): "left,top,right,bottom" (e.g. "10,10,10,10")
  cropLeft, cropTop, cropRight, cropBottom  Individual crop sides (percentage)
  x, y, width, height  Position and size

Video/Audio properties (/slide[N]/video[M] or /slide[N]/audio[M]):
  volume       Playback volume 0-100
  autoplay     true/false — auto-play on slide enter
  trimStart    Start time in ms (e.g. 5000)
  trimEnd      End time in ms
  x, y, width, height  Position and size

Connector properties (/slide[N]/connector[M]):
  name         Connector name
  x, y, width, height  Position and size
  lineWidth    Line width (EMU or cm/pt, e.g. 2pt) (alias: line.width)
  lineColor    Line color hex (e.g. FF0000) (alias: line.color)
  preset       Connector geometry (e.g. straightConnector1, bentConnector3, curvedConnector3)

Group properties (/slide[N]/group[M]):
  name         Group name
  x, y, width, height  Position and size

Master/Layout editing (/slideMaster[N] or /slideLayout[N]):
  name         Change layout/master name
  /slideMaster[N]/shape[M] or /slideLayout[N]/shape[M]  — set shape properties

Presentation properties (/ or /presentation):
  slideSize    Preset: 16:9, 4:3, 16:10, a4
  slideWidth   Custom width (EMU or cm/in/pt/px)
  slideHeight  Custom height (EMU or cm/in/pt/px)
  firstSlideNum    First slide number (default 1)
  rtl              Right-to-left layout (true/false)
  compatMode       Compatibility mode (true/false)
  removePersonalInfo  Remove personal info on save (true/false)

Show properties (/ or /presentation):
  show.loop         Loop continuously (true/false)
  show.narration    Play narration (true/false)
  show.animation    Show animations (true/false)
  show.useTimings   Use slide timings (true/false)

Printing properties (/ or /presentation):
  print.what           Print output: slides, handouts, notes, outline
  print.colorMode      Color mode: color, grayscale, blackAndWhite
  print.hiddenSlides   Print hidden slides (true/false)
  print.scaleToFitPaper Scale to fit paper (true/false)
  print.frameSlides    Frame slides (true/false)

Table properties (/slide[N]/table[M]):
  tableStyle   Built-in style: medium1..4, light1..3, dark1..2, none, or GUID
  x, y, width, height, name

Notes properties (/slide[N]/notes):
  text         Speaker notes text (multi-line supported with \n)

Slide properties (/slide[N]):
  background   Solid color (RRGGBB), gradient (C1-C2 or C1-C2-angle or C1-C2-C3),
               image fill (image:/path/to/file.png), or "none" to remove
               Examples: FF0000  |  FF0000-0000FF  |  FF0000-0000FF-45  |  image:/tmp/bg.png
  transition   Slide transition: fade, push, wipe, split, reveal, random, cover, uncover, zoom, morph, none
               Suffix with speed: fade-fast, push-slow (slow=1200ms, fast=300ms, default=700ms)
               Morph options: morph (default byObject), morph-byWord, morph-byChar
               Morph matches same-named shapes across slides (slide N-1 → slide N).
               Setting transition=morph automatically adds "!!" prefix to shape names on
               the current and previous slide, so shapes morph even when text changes.
               To pair specific shapes: give them the same name on both slides.
               Removing morph (e.g. transition=fade) auto-removes the "!!" prefix.
  advanceTime  Auto-advance after time: "3000" (ms) to advance 3 s after last animation
  advanceClick true/false — advance on click (default true)

Shape animation (/slide[N]/shape[M]):
  link         Hyperlink URL for the shape (applied to all runs). "none" to remove.
               Example: "https://example.com"

  animation    Segments can appear in any order after EFFECT (parsed by content type):
               EFFECT:    appear, fade, fly, zoom, wipe, bounce, float, split, wheel,
                          spin, grow, swivel, checkerboard, blinds, bars, dissolve, flash, none
               CLASS:     entrance/in (default), exit/out, emphasis/emph
               DIRECTION: left/right/up/down/top/bottom (for fly, wipe, etc.)
               DURATION:  milliseconds (default 400)
               TRIGGER:   click, after/afterprevious, with/withprevious
                          Default: 1st on slide = click, subsequent = after (sequential)
               Examples: "fade"  |  "fly-left"  |  "fly-entrance-left-350"
                         "zoom-exit-800"  |  "fade-in-500-after"  |  "wipe-right-300-with"

Table properties (/slide[N]/table[M]):
  x, y, width, height, name

Table row properties (/slide[N]/table[M]/tr[R]):
  height; other props apply to all cells in the row

Table cell properties (/slide[N]/table[M]/tr[R]/tc[C]):
  text, font, size, bold, italic, color, fill, align,
  gridspan/colspan (horizontal merge), rowspan (vertical merge),
  hmerge (true for continuation cell in horizontal merge),
  vmerge (true for continuation cell in vertical merge)

Placeholder properties (/slide[N]/placeholder[M] or /slide[N]/placeholder[type]):
  Same as shape properties. Types: title, body, subtitle, date, footer, slidenum
  If placeholder not on slide, it is auto-created from layout.

Paragraph properties (/slide[N]/shape[M]/paragraph[P]):
  align      left (l), center (c), right (r), justify (j)
  indent     First-line indent (EMU or cm/pt, negative for hanging)
  marginLeft Left margin (alias: marL)
  marginRight Right margin (alias: marR)
  lineSpacing (unit-qualified: '1.5x', '150%', '18pt', or bare multiplier)
  spaceBefore, spaceAfter (unit-qualified: '12pt', '0.5cm', or bare points)
  Plus all run-level properties above

Run properties (/slide[N]/shape[M]/run[K] or /slide[N]/shape[M]/paragraph[P]/run[K]):
  text, font, size, bold, italic, color, spacing,
  baseline    Baseline offset in percent (30 = superscript, -25 = subscript) or "super"/"sub"/"none"
  superscript true/false shorthand (baseline +30%)
  subscript   true/false shorthand (baseline -25%)

Shape z-order (/slide[N]/shape[M]):
  zorder     Z-order: front, back, forward (+1), backward (-1), or absolute position (1=back, N=front)
             Aliases: z-order, order

Any XML attribute is settable via element path (find paths with get --depth N):
  Color:     /slide[1]/cSld/spTree/sp[1]/txBody/p[1]/r[1]/rPr[1]/solidFill[1]/srgbClr[1]  -> val

Examples:
  officecli set pres.pptx '/slide[1]/shape[1]' --prop text="New Title" --prop font=Arial
  officecli set pres.pptx '/slide[1]/shape[1]' --prop fill=4472C4 --prop preset=roundRect
  officecli set pres.pptx '/slide[1]/shape[1]' --prop x=2cm --prop y=3cm --prop width=10cm --prop height=2cm
  officecli set pres.pptx '/slide[1]/table[1]' --prop x=2cm --prop y=3cm --prop width=20cm
  officecli set pres.pptx '/slide[1]/table[1]/tr[1]' --prop bold=true --prop fill=4472C4
  officecli set pres.pptx '/slide[1]/table[1]/tr[1]/tc[1]' --prop text="Header" --prop bold=true --prop fill=4472C4
  officecli set pres.pptx '/slide[1]/placeholder[title]' --prop text="My Title"
  officecli set pres.pptx '/slide[1]/shape[2]/paragraph[1]' --prop align=center
  officecli set pres.pptx '/slide[1]' --prop background=1F3864
  officecli set pres.pptx '/slide[1]' --prop background=FF0000-0000FF-45
  officecli set pres.pptx '/slide[1]' --prop transition=fade --prop advanceTime=3000
  officecli set pres.pptx '/slide[1]/shape[1]' --prop animation=fly-entrance-500
  officecli set pres.pptx '/slide[1]/notes' --prop text="Speaker notes here"
  officecli set pres.pptx '/slide[1]/shape[1]' --prop link="https://example.com"
  officecli set pres.pptx '/slide[1]/shape[1]/run[1]' --prop link="https://example.com"
""";

    const string PptxAdd = """
PowerPoint (.pptx) — add
==========================

Types and properties:

  slide  -- parent: /
    title (optional), text (optional),
    layout (optional) — by name, type, or index:
      Name: "Title Slide", "Title and Content", "Two Content", "Blank", etc.
      Type: blank, title, titleonly, twocontent, titlecontent, section, comparison, caption
      Index: 1, 2, 3, ... (1-based index of available layouts)
    background (optional) — RRGGBB, gradient (C1-C2[-angle]), or image:/path/to/file.png
    transition (optional) — morph, fade, push, wipe, split, reveal, random, none, etc.
    advanceTime (optional) — auto-advance time in ms (e.g. "3000")
    advanceClick (optional) — advance on click (true/false)

  notes  -- parent: /slide[N]
    text (required) — speaker notes text (multi-line with \n)

  shape (textbox)  -- parent: /slide[N]
    text (supports \n for line breaks), name, font, size, bold, italic,
    underline, strikethrough, color, fill,
    line (border color), lineWidth, lineDash, lineOpacity,
    margin (text padding: 0.5cm or left,top,right,bottom),
    align (left/center/right/justify), valign (top/center/bottom),
    gradient (e.g. FF0000-0000FF-90), list (bullet/numbered/alpha/roman),
    lineSpacing ('1.5x', '150%', '18pt', or bare multiplier),
    spaceBefore, spaceAfter ('12pt', '0.5cm', or bare points),
    rotation, opacity, autoFit,
    preset (shape geometry: rect, roundRect, ellipse, triangle, diamond, pentagon, hexagon,
            star5, rightArrow, leftArrow, chevron, plus, heart, cloud, cube, can, line,
            callout, process, decision, smiley, frame, gear6, ...),
    x, y, width, height (EMU or cm/in/pt/px, default: full-width text box)

  chart  -- parent: /slide[N]
    chartType (column|bar|line|pie|doughnut|area|scatter|bubble|radar|stock|combo|column3d|bar3d), title, legend (top|bottom|left|right|none),
    colors (comma-separated series colors, e.g. "FF0000,00FF00,0000FF"),
    comboSplit (for combo: how many series are columns, default 1),
    categories (comma-separated labels, e.g. "Q1,Q2,Q3,Q4"),
    Data format (choose one):
      data = "Series1:1,2,3;Series2:4,5,6"  (compact: all series in one prop)
      series1 = "Revenue:100,200,300"        (numbered: one prop per series)
      series2 = "Cost:80,150,250"
    x, y, width, height (EMU or cm/in/pt/px), name

  table  -- parent: /slide[N]
    rows (default 3), cols (default 3), name,
    x, y, width, height (EMU or cm/in/pt/px)

  row (tr)  -- parent: /slide[N]/table[M]
    cols (int, default: match existing), height (EMU or cm/in/pt/px),
    c1, c2, ... (cell text shortcuts)

  cell (tc)  -- parent: /slide[N]/table[M]/tr[R]
    text

  picture (image, img)  -- parent: /slide[N]
    path (required), name, alt, width, height, x, y
    Formats: .png, .jpg, .jpeg, .gif, .bmp, .tif, .tiff, .emf, .wmf, .svg
    Default position: centered. Default size: ~6x4 inches.

  connector (connection, line)  -- parent: /slide[N]
    x, y, width, height (position and extent), name,
    preset (straight|elbow|curve, default: straight),
    line (color: hex or theme), lineWidth,
    startShape, endShape (shape IDs to connect to)

  group  -- parent: /slide[N]
    shapes (required) — comma-separated shape indices to group (e.g. shapes=2,3)
    name (optional)

  video (audio, media)  -- parent: /slide[N]
    path (required), name, x, y, width, height,
    poster (cover image path, optional),
    volume (0-100, default 80), autoplay (true|false, default false),
    trimStart (start time in ms, e.g. 5000), trimEnd (end time in ms)
    Formats: .mp4, .avi, .wmv, .mpg, .mov, .mp3, .wav, .wma, .m4a

  equation (formula, math)  -- parent: /slide[N]
    formula (required, LaTeX subset), name

  paragraph (para)  -- parent: /slide[N]/shape[M]
    text, font, size, bold, italic, color, spacing,
    align, indent (EMU/cm/pt), marginLeft (marL), marginRight (marR),
    list (bullet/numbered)

  run  -- parent: /slide[N]/shape[M] or /slide[N]/shape[M]/paragraph[P]
    text, font, size, bold, italic, underline, strike, color, spacing
    When parent is a shape, appends to the last paragraph.
    Use to create rich text with mixed formatting within a single text box.

--index is 0-based. --from clones existing elements (cross-slide relationships handled).

Examples:
  officecli add pres.pptx / --type slide --prop title="Agenda" --prop text="Topics for today"
  officecli add pres.pptx / --type slide --prop layout=blank
  officecli add pres.pptx / --type slide --prop layout="Title Slide" --prop title="Welcome"
  officecli add pres.pptx / --type slide --prop layout=twocontent --prop title="Comparison"
  officecli add pres.pptx / --type slide --prop title="Dark Slide" --prop background=1F3864
  officecli add pres.pptx '/slide[1]' --type shape --prop text="Hello" --prop font=Arial --prop size=18
  officecli add pres.pptx '/slide[1]' --type shape --prop text="Go" --prop preset=rightArrow --prop fill=4472C4
  officecli add pres.pptx '/slide[1]' --type chart --prop chartType=column --prop title="Q1 Sales" --prop categories="Q1,Q2,Q3,Q4" --prop data="Revenue:100,200,300,400;Cost:80,150,250,350"
  officecli add pres.pptx '/slide[1]' --type chart --prop chartType=pie --prop title="Market Share" --prop categories="Apple,Google,MS" --prop series1="Share:40,30,30"
  officecli add pres.pptx '/slide[1]' --type chart --prop chartType=line --prop series1="Trend:1,3,2,5" --prop legend=top
  officecli add pres.pptx '/slide[1]' --type table --prop rows=3 --prop cols=4
  officecli add pres.pptx '/slide[1]/table[1]' --type row --prop c1="Name" --prop c2="Value"
  officecli add pres.pptx '/slide[1]/table[1]/tr[1]' --type cell --prop text="Extra"
  officecli add pres.pptx '/slide[1]' --type picture --prop path=photo.jpg --prop width=8cm --prop alt="Team photo"
  officecli add pres.pptx '/slide[1]' --type equation --prop formula="\frac{-b \pm \sqrt{b^2-4ac}}{2a}"
  officecli add pres.pptx '/slide[1]/shape[1]' --type paragraph --prop text="New para" --prop bold=true
  officecli add pres.pptx '/slide[1]/shape[1]/paragraph[1]' --type run --prop text=" bold red" --prop bold=true --prop color=FF0000
  officecli add pres.pptx '/slide[1]/shape[1]' --type run --prop text=" (appended)" --prop italic=true
  officecli add pres.pptx / --from '/slide[1]' --index 0
  officecli add pres.pptx '/slide[2]' --from '/slide[1]/shape[2]'
  officecli add pres.pptx '/slide[1]' --type notes --prop text="Key talking points\nRemember to pause here"
""";

    const string PptxRaw = """
PowerPoint (.pptx) — raw
==========================

Available parts:
  / or /presentation     Presentation root (slide size, slide list)
  /slide[N]              Slide N XML
  /slideMaster[N]        Slide master N
  /slideLayout[N]        Slide layout N
  /noteSlide[N]          Notes page for slide N

raw-set actions: append, prepend, insertbefore, insertafter, replace, remove, setattr
No xmlns needed -- prefixes auto-registered: p, a, r

PPT slide workflow (add content beyond L2 capabilities):
  1. officecli raw pres.pptx /presentation    # check sldSz for dimensions
  2. officecli add pres.pptx / --type slide   # create slide via L2
  3. officecli raw-set pres.pptx /slide[N] ...  # fill with raw XML

Chart workflow:
  1. officecli add-part pres.pptx /slide[1] --type chart   # get relId
  2. officecli raw-set pres.pptx /slide[1]/chart[1] ...    # fill chart XML

Examples:
  officecli raw pres.pptx /presentation
  officecli raw pres.pptx '/slide[1]'
  officecli raw pres.pptx '/slideMaster[1]'
  officecli raw-set pres.pptx '/slide[1]' --xpath "//p:cSld/p:spTree" --action append --xml '<p:sp>...</p:sp>'
""";
}
