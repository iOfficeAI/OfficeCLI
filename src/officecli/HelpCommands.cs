// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;

namespace OfficeCli;

/// <summary>
/// Format-specific help: docx, xlsx, pptx with nested verb help (view, get, set, add, etc.).
/// These are help-only commands — they do not execute any document operations.
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
    /// Matches: docx [verb] [--help], xlsx [verb] [--help], pptx [verb] [--help]
    /// </summary>
    internal static bool TryHandle(string[] args)
    {
        if (args.Length == 0) return false;
        var format = args[0].ToLowerInvariant();
        if (format is not ("docx" or "xlsx" or "pptx")) return false;

        // Extract verb (skip --help flags)
        string? verb = null;
        foreach (var arg in args.Skip(1))
        {
            if (arg is "--help" or "-h" or "-?") continue;
            verb = arg.ToLowerInvariant();
            break;
        }

        var help = GetHelp(format, verb);
        Console.WriteLine(help);
        return true;
    }

    static string GetHelp(string format, string? verb) => format switch
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
  text (t)       Plain text with line numbers. Use --start/--end/--max-lines to paginate.
  annotated (a)  Text with formatting details (font, size, bold, style, LaTeX formulas).
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

Supports any XML path via element localName:
  /body/p[3]              Paragraph 3
  /body/tbl[1]/tblPr      Table 1 properties
  /body/p[1]/r[1]/rPr     Run properties of first run

Options:
  --depth N   Depth of child nodes to include (default 1)
  --json      Output as JSON

Examples:
  officecli get doc.docx /                     # document root
  officecli get doc.docx '/body/p[1]' --depth 3
  officecli get doc.docx '/body/tbl[1]/tr[1]/tc[1]' --json
""";

    const string DocxQuery = """
Word (.docx) — query
=====================

Element types:  paragraph (p), run (r), table (tbl), picture, equation
Attribute filters:  [attr=value], [attr!=value]
Pseudo-selectors:   :contains("text"), :empty, :no-alt, :has(formula)
Child combinator:   paragraph > run[bold=true]
Generic XML:        Falls back to any XML element name (e.g. wsp, srgbClr[val=0070C0])

Examples:
  officecli query doc.docx 'paragraph[style=Heading1]'
  officecli query doc.docx 'run[bold=true]'
  officecli query doc.docx 'paragraph:contains("error")'
  officecli query doc.docx 'paragraph:empty'
  officecli query doc.docx 'picture:no-alt'
  officecli query doc.docx 'paragraph > run[font!=Arial]'
  officecli query doc.docx 'paragraph[alignment=center]'
""";

    const string DocxSet = """
Word (.docx) — set
===================

Usage: officecli set <file> <path> --prop key=value [--prop key=value ...]

Run properties (/body/p[N]/r[M]):
  text, font, size, bold, italic, color, underline, strike, highlight,
  caps, smallCaps, dstrike, vanish, outline, shadow, emboss, imprint,
  noProof, rtl, shd (format: "fill" or "pattern;fill" or "pattern;fill;color")
  For images in runs: alt, width, height (cm/in/pt/px or raw EMU)

Paragraph properties (/body/p[N]):
  style, alignment (left|center|right|justify), firstLineIndent,
  shd, spaceBefore, spaceAfter, lineSpacing, numId, numLevel/ilvl,
  listStyle (bullet|numbered|none), start (numbering start value)

Table cell properties (/body/tbl[N]/tr[R]/tc[C]):
  text, font, size, bold, italic, color, shd, alignment,
  valign (top|center|bottom), width, vmerge (restart|continue), gridspan

Table row properties (/body/tbl[N]/tr[R]):
  height, header (bool)

Table properties (/body/tbl[N]):
  alignment (left|center|right), width

Document root (/):
  defaultFont, pageBackground, pageWidth, pageHeight,
  marginTop, marginBottom, marginLeft, marginRight

Any XML attribute is also settable via element path (use get --depth N to find paths).
Composite props (pBdr, tabs, lang, bdr) -> use raw-set instead.

Examples:
  officecli set doc.docx '/body/p[1]/r[1]' --prop bold=true --prop color=FF0000
  officecli set doc.docx '/body/p[2]' --prop style=Heading1 --prop alignment=center
  officecli set doc.docx '/body/tbl[1]/tr[1]/tc[1]' --prop text="Hello" --prop shd=4472C4
  officecli set doc.docx '/body/tbl[1]/tr[1]/tc[1]' --prop valign=center --prop gridspan=2
  officecli set doc.docx / --prop defaultFont=Arial --prop marginTop=1440
  officecli set doc.docx '/body/p[5]/r[1]' --prop width=5cm --prop alt="Logo"
  officecli set doc.docx '/body/p[1]/r[1]' --prop link="https://example.com"
""";

    const string DocxAdd = """
Word (.docx) — add
===================

Usage: officecli add <file> <parent> --type <type> [--index N] [--prop key=value ...]
   or: officecli add <file> <parent> --from <path> [--index N]  (clone existing element)

Types and properties:

  paragraph (p)  -- parent: /body or /body/tbl[N]/tr[R]/tc[C]
    text, font, size, bold, italic, color, underline, strike, highlight,
    caps, smallCaps, style, alignment, firstLineIndent, spaceBefore,
    spaceAfter, lineSpacing, numId, numLevel, shd, listStyle, start

  run (r)  -- parent: /body/p[N]
    text, font, size, bold, italic, color, underline, strike, highlight,
    caps, smallCaps, shd

  table (tbl)  -- parent: /body
    rows (int), cols (int)

  picture (image, img)  -- parent: /body/p[N] or /body
    path (required), width, height (cm/in/pt/px/EMU), alt

  equation (formula, math)  -- parent: /body/p[N] or /body
    formula (required, LaTeX subset), mode (display|inline)
    Supported: \frac{}{}, \sqrt{}, ^{}, _{}, \sum, \int, Greek letters

  hyperlink (link)  -- parent: /body/p[N]
    url (required), text (display text, defaults to url)
    font, size (optional run formatting)

  comment  -- parent: /body/p[N] or /body/p[N]/r[M]
    text (required), author, initials, date (ISO format)

--index is 0-based. If omitted, appends to end.
--from clones an element (cross-part relationships handled automatically).

Examples:
  officecli add doc.docx /body --type paragraph --prop text="Hello World" --prop style=Heading1
  officecli add doc.docx '/body/p[1]' --type run --prop text="bold text" --prop bold=true
  officecli add doc.docx /body --type table --prop rows=3 --prop cols=4
  officecli add doc.docx /body --type picture --prop path=logo.png --prop width=5cm
  officecli add doc.docx '/body/p[1]' --type equation --prop formula="\frac{a}{b}"
  officecli add doc.docx '/body/p[3]' --type comment --prop text="Please review"
  officecli add doc.docx /body --from '/body/p[1]' --index 5
""";

    const string DocxRaw = """
Word (.docx) — raw
===================

Available parts:
  /document      Main document body (default)
  /styles        Style definitions
  /numbering     List/numbering definitions
  /settings      Document settings
  /header[N]     Header N (0-based)
  /footer[N]     Footer N (0-based)

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

Falls back to generic XML element navigation for advanced queries.

Examples:
  officecli query data.xlsx 'cell[formula=true]'
  officecli query data.xlsx 'Sheet1!cell[empty=false]'
  officecli query data.xlsx 'A'
  officecli query data.xlsx 'cell:contains("error")'
  officecli query data.xlsx 'cell[type=Number]'
""";

    const string XlsxSet = """
Excel (.xlsx) — set
====================

Usage: officecli set <file> '/SheetName/CellRef' --prop key=value [--prop ...]

Content properties:
  value          Cell value (auto-detects number vs string)
  formula        Formula (e.g. =SUM(A1:A10))
  type           Force type: string|str, number|num, boolean|bool
  clear          Clear cell value and formula

Font properties (prefix with font.):
  font.bold, font.italic, font.strike, font.underline (true/false or single/double)
  font.color     Hex RGB (e.g. FF0000)
  font.size      Point size (e.g. 11)
  font.name      Font family (e.g. Calibri)

Fill:
  fill           Hex RGB background color (e.g. 4472C4)

Alignment (prefix with alignment.):
  alignment.horizontal    left, center, right, justify
  alignment.vertical      top, center, bottom
  alignment.wrapText      true/false

Number format:
  numFmt         Format string (e.g. "0%", "0.00", "#,##0")

Hyperlink:
  link           URL for the cell hyperlink. "none" to remove.

Examples:
  officecli set data.xlsx '/Sheet1/A1' --prop value=100
  officecli set data.xlsx '/Sheet1/B2' --prop formula="=SUM(A1:A10)"
  officecli set data.xlsx '/Sheet1/A1' --prop font.bold=true --prop font.color=FF0000
  officecli set data.xlsx '/Sheet1/C3' --prop fill=4472C4 --prop font.size=14
  officecli set data.xlsx '/Sheet1/A1' --prop alignment.horizontal=center --prop alignment.wrapText=true
  officecli set data.xlsx '/Sheet1/D1' --prop numFmt="0.00%"
  officecli set data.xlsx '/Sheet1/A1' --prop clear=true
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

--index is 0-based. --from clones an existing element.

Examples:
  officecli add data.xlsx / --type sheet --prop name=Summary
  officecli add data.xlsx /Sheet1 --type row --prop cols=5
  officecli add data.xlsx /Sheet1 --type cell --prop ref=A1 --prop value=100 --prop fill=4472C4
  officecli add data.xlsx /Sheet1 --type databar --prop sqref=B2:B20 --prop color=63C384
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

Common workflow:
  1. officecli view pres.pptx outline              # slide titles overview
  2. officecli view pres.pptx annotated             # shapes, fonts, sizes
  3. officecli get  pres.pptx '/slide[1]' --depth 2
  4. officecli set  pres.pptx '/slide[1]/shape[1]' --prop text="Title"
  5. officecli validate pres.pptx

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
  text (t)       Plain text from all shapes, slide by slide.
  annotated (a)  Shape details: type, text, font, size, pictures with alt text, equations.
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
    preset                 Shape geometry name
    align, valign
    lineSpacing            Multiplier (e.g. 1.5) from first paragraph
    spaceBefore, spaceAfter  Points from first paragraph
    margin                 Text padding
    rotation               Degrees
    autoFit                normal / shape / none
    list                   Bullet char or auto-number type
    link                   Hyperlink URL (from first run)
    shadow, glow           Effect color hex
    reflection             "true" if reflection effect applied
    animation              "effectName-class-durationMs" (e.g. "fade-entrance-500")

  Chart (/slide[N]/chart[M]):
    chartType              column, bar, line, pie, doughnut, area, scatter
    title                  Chart title text
    legend                 Legend position (t/b/l/r)
    seriesCount            Number of data series
    categories             Comma-separated category labels
    x, y, width, height    Position and size
    (depth>0: children = series nodes with name + values)

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

Scope by slide:
  slide[N] shape     Only shapes in slide N

Falls back to generic XML element name for advanced queries.

Examples:
  officecli query pres.pptx 'shape'
  officecli query pres.pptx 'slide[1] shape'
  officecli query pres.pptx 'title'
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
  preset     Shape geometry (e.g. roundRect, ellipse, rightArrow, diamond, star5)
  margin     Text padding inside shape (e.g. 0.5cm or left,top,right,bottom: 0.5cm,0.3cm,0.5cm,0.3cm)
  align      Text horizontal alignment: left (l), center (c), right (r), justify (j) — applies to all paragraphs
  valign     Text vertical alignment: top (t), center/middle (c/m), bottom (b)
  lineSpacing  Line spacing multiplier (e.g. 1.5 for 150%)
  spaceBefore  Space before paragraphs in points (e.g. 6)
  spaceAfter   Space after paragraphs in points (e.g. 6)
  gradient   Linear: C1-C2[-angle], Radial: radial:C1-C2[-focus] (focus: tl/tr/bl/br/center)
  image      Shape image fill (path to image file, e.g. /tmp/bg.png)
  list       List style: bullet/numbered/alpha/roman/none or a custom character (e.g. ✓)
  rotation   Rotation angle in degrees (e.g. 45) (alias: rotate)
  opacity    Fill opacity 0.0-1.0 (e.g. 0.5 for 50%)
  autoFit    Text auto-fit: true/normal, shape, false/none
  x          Horizontal position (EMU or cm/in/pt/px, e.g. 2cm)
  y          Vertical position (EMU or cm/in/pt/px, e.g. 3cm)
  width      Shape width (EMU or cm/in/pt/px, e.g. 10cm)
  height     Shape height (EMU or cm/in/pt/px, e.g. 2cm)

Chart properties (/slide[N]/chart[M]):
  title        Chart title text (or "none" to remove)
  legend       Legend position: top/bottom/left/right or "none" to remove
  x, y, width, height  Chart position and size (EMU or cm/in/pt/px)
  name         Chart name

Notes properties (/slide[N]/notes):
  text         Speaker notes text (multi-line supported with \n)

Slide properties (/slide[N]):
  background   Solid color (RRGGBB), gradient (C1-C2 or C1-C2-angle or C1-C2-C3),
               image fill (image:/path/to/file.png), or "none" to remove
               Examples: FF0000  |  FF0000-0000FF  |  FF0000-0000FF-45  |  image:/tmp/bg.png
  transition   Slide transition: fade, push, wipe, split, reveal, random, cover, uncover, zoom, none
               Suffix with speed: fade-fast, push-slow (slow=1200ms, fast=300ms, default=700ms)
  advanceTime  Auto-advance after time: "3000" (ms) to advance 3 s after last animation
  advanceClick true/false — advance on click (default true)

Shape animation (/slide[N]/shape[M]):
  link         Hyperlink URL for the shape (applied to all runs). "none" to remove.
               Example: "https://example.com"

  animation    EFFECT[-CLASS[-DURATION[-TRIGGER]]]
               EFFECT:  appear, fade, fly, zoom, wipe, bounce, float, split, wheel,
                        spin, grow, swivel, checkerboard, blinds, bars, dissolve, flash, none
               CLASS:   entrance/in (default), exit/out, emphasis/emph
               DURATION: milliseconds (default 500)
               TRIGGER: click (default), after/afterprevious, with/withprevious
               Examples: "fade"  |  "fly-entrance"  |  "zoom-exit-800"  |  "fade-in-500-after"

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
  Plus all run-level properties above

Run properties (/slide[N]/shape[M]/run[K] or /slide[N]/shape[M]/paragraph[P]/run[K]):
  text, font, size, bold, italic, color

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

  notes  -- parent: /slide[N]
    text (required) — speaker notes text (multi-line with \n)

  shape (textbox)  -- parent: /slide[N]
    text (supports \n for line breaks), name, font, size, bold, italic,
    underline, strikethrough, color, fill,
    line (border color), lineWidth, lineDash, lineOpacity,
    margin (text padding: 0.5cm or left,top,right,bottom),
    align (left/center/right/justify), valign (top/center/bottom),
    gradient (e.g. FF0000-0000FF-90), list (bullet/numbered/alpha/roman),
    lineSpacing, spaceBefore, spaceAfter, rotation, opacity, autoFit,
    preset (shape geometry: rect, roundRect, ellipse, triangle, diamond, pentagon, hexagon,
            star5, rightArrow, leftArrow, chevron, plus, heart, cloud, cube, can, line,
            callout, process, decision, smiley, frame, gear6, ...),
    x, y, width, height (EMU or cm/in/pt/px, default: full-width text box)

  chart  -- parent: /slide[N]
    chartType (column|bar|line|pie|doughnut|area|scatter), title, legend (top|bottom|left|right|none),
    categories (comma-separated labels, e.g. "Q1,Q2,Q3,Q4"),
    Data format (choose one):
      data = "Series1:1,2,3;Series2:4,5,6"  (compact: all series in one prop)
      series1 = "Revenue:100,200,300"        (numbered: one prop per series)
      series2 = "Cost:80,150,250"
    x, y, width, height (EMU or cm/in/pt/px), name

  table  -- parent: /slide[N]
    rows (default 3), cols (default 3), name,
    x, y, width, height (EMU or cm/in/pt/px)

  picture (image, img)  -- parent: /slide[N]
    path (required), name, alt, width, height, x, y
    Formats: .png, .jpg, .jpeg, .gif, .bmp, .tif, .tiff, .emf, .wmf, .svg
    Default position: centered. Default size: ~6x4 inches.

  equation (formula, math)  -- parent: /slide[N]
    formula (required, LaTeX subset), name

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
  officecli add pres.pptx '/slide[1]' --type picture --prop path=photo.jpg --prop width=8cm --prop alt="Team photo"
  officecli add pres.pptx '/slide[1]' --type equation --prop formula="\frac{-b \pm \sqrt{b^2-4ac}}{2a}"
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
