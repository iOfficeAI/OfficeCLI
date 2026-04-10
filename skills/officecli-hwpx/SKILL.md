---
name: officecli-hwpx
description: "Use this skill any time a .hwpx file is involved -- as input or for analysis. This includes: reading, parsing, or extracting text from any .hwpx file; editing or modifying existing HWPX documents; querying document structure; validating HWPX integrity; working with Korean (한글) office documents. Trigger whenever the user mentions 'HWP', 'HWPX', '한글 문서', '한글 파일', 'Hancom', or references a .hwpx filename. NOTE: creating new .hwpx files is NOT supported -- use python-hwpx or hwp_origin tools instead."
---

# OfficeCLI HWPX Skill

## Quick Decision

| Task | Supported? | Action |
|------|-----------|--------|
| Read / analyze .hwpx | ✅ Yes | Use `view`, `get`, `query` commands below |
| Edit existing .hwpx | ✅ Yes | Use `set`, `raw-set`, `move`, `remove`, `copy` |
| Validate .hwpx | ✅ Yes | Use `validate` command |
| View raw XML parts | ✅ Yes | Use `raw` command |
| Create new .hwpx | ❌ No | Use `python-hwpx` or `hwp_origin` Python tools |
| Open .hwp (binary) | ❌ No | Convert to .hwpx first (Hancom Office or LibreOffice) |
| Visual preview (html/svg) | ❌ No | Open in Hancom Office or LibreOffice for visual check |
| Watch for changes | ❌ No | Not implemented for HWPX |

---

## Core Command Model

### Binary Location

```bash
OFFICECLI="700_projects/cli-jaw/officecli/build-local/officecli"

# Build if missing
if [ ! -f "$OFFICECLI" ]; then
    cd 700_projects/cli-jaw/officecli && dotnet publish -c Release -o build-local
fi
```

### Text Extraction

```bash
officecli view file.hwpx text
officecli view file.hwpx text --max-lines 200
officecli view file.hwpx text --start 1 --end 50
```

`text` mode extracts body text with automatic Korean normalization: PUA character stripping, shape alt-text removal (e.g., "사각형입니다."), and uniform-distribution spacing collapse (현 장 대 응 → 현장대응).

### Structure Overview

```bash
officecli view file.hwpx outline
```

Shows section count, paragraph/table counts, and document structure tree.

### Detailed Inspection

```bash
officecli view file.hwpx annotated
```

Shows per-element detail with namespace-aware rendering of `hp:`, `hs:`, `hh:`, `hc:` elements.

### Statistics

```bash
officecli view file.hwpx stats
```

Element counts, section distribution, and structural summary.

### Issue Detection

```bash
officecli view file.hwpx issues
officecli view file.hwpx issues --type format
officecli view file.hwpx issues --type content
officecli view file.hwpx issues --type structure
```

### Element Inspection (get)

```bash
# Document root
officecli get file.hwpx /

# Specific paragraph in a section (1-based)
officecli get file.hwpx "/section[1]/p[1]"

# Table in a section
officecli get file.hwpx "/section[2]/tbl[1]" --depth 3

# Table cell
officecli get file.hwpx "/section[1]/tbl[1]/tr[2]/tc[3]"

# Search across all sections (omit section index)
officecli get file.hwpx "/p[5]"
officecli get file.hwpx "/tbl[1]"
```

### CSS-like Queries

```bash
# Find paragraphs containing text
officecli query file.hwpx 'p:contains("분기")'

# Find empty elements
officecli query file.hwpx 'p:empty'

# Find elements with specific children
officecli query file.hwpx 'tbl:has(tr)'
```

**Available pseudo-selectors:** `:contains("text")`, `:empty`, `:has(child)`
**NOT available:** `:heading` (not implemented for HWPX)

### Set Properties

```bash
# Set paragraph text (use --prop, NOT --props)
officecli set file.hwpx "/section[1]/p[1]" --prop text="새로운 텍스트"

# Set multiple properties
officecli set file.hwpx "/section[1]/p[2]" --prop text="Hello" --prop bold=true
```

**Critical:** The flag is `--prop` (singular), not `--props`.

### Raw XML Access

```bash
# View a raw XML part inside the HWPX archive
officecli raw file.hwpx 'Contents/section0.xml'

# Modify raw XML
officecli raw-set file.hwpx 'Contents/section0.xml' \
    --xpath '//hp:p[1]' \
    --action replace \
    --xml '<hp:p>replacement content</hp:p>'
```

**`raw-set` actions (7 total):**

| Action | Description |
|--------|-------------|
| `append` | Add as last child of matched element |
| `prepend` | Add as first child of matched element |
| `insertbefore` | Insert before matched element |
| `insertafter` | Insert after matched element |
| `replace` | Replace matched element entirely |
| `remove` | Remove matched element |
| `setattr` | Set attribute on matched element |

### Mutation Commands

```bash
# Move element
officecli move file.hwpx "/section[1]/p[3]" --to "/section[1]/p[1]"

# Remove element
officecli remove file.hwpx "/section[1]/p[2]"

# Copy element
officecli copy file.hwpx "/section[1]/p[1]" --to "/section[2]"
```

### Validation

```bash
officecli validate file.hwpx
```

Six-level validation:
1. **ZIP integrity** — archive is valid and extractable
2. **Required files** — `mimetype`, `container.xml`, `content.hpf` present
3. **XML well-formedness** — all XML parts parse correctly
4. **ID reference consistency** — internal ID references resolve
5. **Table structure** — row/cell counts are consistent
6. **Namespace declarations** — `hp:`, `hs:`, `hh:`, `hc:` prefixes declared

---

## Path Syntax (HWPX-Specific)

HWPX uses section-based paths, unlike DOCX's `/body/` root:

```
/section[N]/p[M]                  # Mth paragraph in Nth section (1-based)
/section[N]/tbl[M]                # Mth table in Nth section
/section[N]/tbl[M]/tr[R]/tc[C]    # Cell at row R, column C
/p[N]                             # Nth paragraph (searches all sections)
/tbl[N]                           # Nth table (searches all sections)
```

**Always quote paths in shell** to prevent glob expansion of `[N]`:
```bash
officecli get file.hwpx "/section[1]/p[1]"   # ✅ correct
officecli get file.hwpx /section[1]/p[1]      # ❌ shell glob expansion
```

---

## Common Workflows

### 1. Analyze an Unknown HWPX Document

```bash
# Step 1: Quick overview
officecli view file.hwpx outline

# Step 2: Full text extraction
officecli view file.hwpx text

# Step 3: Check for issues
officecli view file.hwpx issues

# Step 4: Validate structure
officecli validate file.hwpx
```

### 2. Find and Replace Text

```bash
# Step 1: Find paragraphs containing target text
officecli query file.hwpx 'p:contains("기존 텍스트")'

# Step 2: Note the path from results, then set new text
officecli set file.hwpx "/section[1]/p[3]" --prop text="새로운 텍스트"
```

### 3. Inspect Table Data

```bash
# Find all tables
officecli query file.hwpx 'tbl:has(tr)'

# Inspect table structure
officecli get file.hwpx "/section[1]/tbl[1]" --depth 3

# Read specific cell
officecli get file.hwpx "/section[1]/tbl[1]/tr[1]/tc[1]"
```

### 4. Raw XML Surgery

```bash
# Step 1: View the raw XML to understand structure
officecli raw file.hwpx 'Contents/section0.xml'

# Step 2: Modify with precise XPath
officecli raw-set file.hwpx 'Contents/section0.xml' \
    --xpath '//hp:run[1]/hp:t' \
    --action replace \
    --xml '<hp:t>수정된 텍스트</hp:t>'
```

---

## CJK / Korean Handling

HWPX is the native format for Korean (한글) documents. Korean text normalization is **automatic** during text extraction:

| Normalization | Example | Notes |
|--------------|---------|-------|
| PUA character stripping | U+E000–U+F8FF removed | Hancom proprietary chars |
| Uniform spacing collapse | 현 장 대 응 → 현장대응 | Common in scanned/OCR docs |
| Shape alt-text removal | "사각형입니다." stripped | Auto-generated by Hancom |

**When working with Korean text in commands:**
- Use UTF-8 terminal encoding
- Quote text arguments: `--prop text="한글 텍스트"`
- `:contains()` queries work with Korean: `'p:contains("보고서")'`

---

## QA Verification (Required)

**Assume there are problems. Your job is to find them.**

### Verification Loop

1. Run `view outline` — check section/element counts
2. Run `view text` — verify content extraction is complete
3. Run `view issues` — review all flagged issues
4. Run `validate` — ensure structural integrity
5. Fix any issues found
6. Re-verify — one fix often creates another problem
7. Repeat until a full pass reveals no new issues

**Do not declare success until you've completed at least one fix-and-verify cycle.**

### Pre-Delivery Checklist

- [ ] All text extracted cleanly (no garbled characters)
- [ ] Korean normalization applied correctly
- [ ] No structural issues in `view issues`
- [ ] Validation passes all 6 levels
- [ ] Table structures are consistent
- [ ] No orphaned ID references
- [ ] Namespace declarations are complete

**NOTE:** There is no visual preview mode (html/svg) for HWPX. Content verification relies on `view text`, `view annotated`, `view outline`, `view issues`, and `validate`. For visual verification, the user must open the file in Hancom Office or LibreOffice.

---

## Common Pitfalls

| Pitfall | Correct Approach |
|---------|-----------------|
| `--props text=Hello` | Use `--prop text=Hello` — singular `--prop` flag |
| Using `/body/p[1]` path | HWPX uses `/section[1]/p[1]` — section-based, not body-based |
| Opening `.hwp` (binary) | Convert to `.hwpx` first — binary HWP is not supported |
| Trying `officecli create *.hwpx` | Not supported — use `python-hwpx` or `hwp_origin` |
| Unquoted `[N]` in shell | Always quote: `"/section[1]/p[1]"` to prevent glob expansion |
| Assuming `:heading` works | `:heading` pseudo-selector is NOT available for HWPX |
| Forgetting namespace prefixes in raw XML | HWPX uses `hp:`, `hs:`, `hh:`, `hc:` prefixes |
| Using `add-part` for HWPX | Not supported — HWPX uses OPF packaging, not OPC |
| Expecting `watch/unwatch` | Not implemented for HWPX yet |

---

## HWPX Format Reference

- **Standard:** OWPML (Open Word-Processor Markup Language) by Hancom
- **Packaging:** ZIP archive with XML parts (similar to OOXML but OPF-based)
- **Namespace prefixes:** `hp:` (paragraph), `hs:` (section), `hh:` (header), `hc:` (character)
- **Key archive entries:**
  - `mimetype` — MIME type declaration
  - `META-INF/container.xml` — container manifest
  - `Contents/content.hpf` — OPF content package file (manifest + spine)
  - `Contents/header.xml` — styles, fonts, numbering definitions
  - `Contents/section0.xml`, `section1.xml`, ... — section content
- **Encoding:** UTF-8

---

## Dependencies

- **Binary:** `700_projects/cli-jaw/officecli/build-local/officecli`
- **Build:** `cd 700_projects/cli-jaw/officecli && dotnet publish -c Release -o build-local`
- **Runtime:** .NET 10+ SDK
- **For creating HWPX:** Use `python-hwpx` or `hwp_origin` (separate Python tools)

---

## Essential Rules

1. **View mode is REQUIRED** — `officecli view file.hwpx` alone will error; specify `text`, `annotated`, `outline`, `stats`, or `issues`
2. **Paths are 1-based** — `/section[1]/p[1]` is the first paragraph in the first section
3. **Always quote paths** — prevent shell glob expansion of bracket syntax
4. **Use `--prop` not `--props`** — singular flag for property setting
5. **No create support** — use external Python tools for HWPX creation
6. **Binary HWP is rejected** — officecli throws a helpful error suggesting .hwpx conversion
7. **Korean normalization is automatic** — no flags needed for PUA stripping or spacing collapse
8. **Verify after every edit** — run `view issues` + `validate` after mutations
