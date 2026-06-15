---
name: officecli-track-changes
description: "Track Changes / 修订模式 for .docx files. Use when the user wants to review, audit, redline, or revise a document with tracked changes (修订模式, 审查, redline, revision). MUST load alongside officecli skill."
---

# Track Changes / Revision Mode (Word .docx)

When the user asks to review, audit, redline, or revise a document with tracked changes, use `--prop revision.type=...` / `--prop revision.author=...`. Do NOT use python-docx, unpack/pack, raw-set, or any other approach — officecli v1.0.98+ handles OOXML revision markup natively via the `revision.*` API.

**API migration**: Previous `trackChange=ins/del/format` has been replaced by `revision.type=ins/del/format`. See the property table below.

## Preferred: find + revision (Word-style Find & Replace with Track Changes)

One command does it all — the handler auto-creates paired ins+del markers:

```bash
# Replace text with tracked change (one command)
officecli set "$FILE" /body --prop find="old text" --prop replace="new text" --prop revision.author=AI

# Delete text (tracked deletion — no insertion)
officecli set "$FILE" /body --prop find="text to remove" --prop replace= --prop revision.author=AI

# Format change on matched text (bold all occurrences of "keyword")
officecli set "$FILE" /body --prop find="keyword" --prop bold=true --prop revision.author=AI

# Regex-based find+replace
officecli set "$FILE" /body --prop 'find=\$\d+' --prop regex=true --prop replace="[PRICE]" --prop revision.author=AI
```

## Alternate: Direct run wrapping with `set` + `revision.type`

```bash
# Mark existing text as inserted
officecli set "$FILE" '/body/p[1]/r[1]' --prop revision.type=ins --prop revision.author=AI

# Mark existing text as deleted
officecli set "$FILE" '/body/p[1]/r[1]' --prop revision.type=del --prop revision.author=AI

# Move revision (paired — same revision.id)
officecli set "$FILE" '/body/p[1]/r[1]' --prop revision.type=moveFrom --prop revision.author=AI --prop revision.id=100
officecli set "$FILE" '/body/p[2]/r[1]' --prop revision.type=moveTo --prop revision.author=AI --prop revision.id=100
```

## Legacy: Split+remove+add workflow (complex multi-change paragraphs)

```bash
# Step 0: Read run structure FIRST
officecli get "$FILE" '/body/p[N]' --depth 2

# Step 1: Isolate target into own run
officecli set "$FILE" '/body/p[N]' --prop find="8" --prop color=auto

# Step 2: Remove original run
officecli remove "$FILE" '/body/p[N]/r[2]'

# Step 3: Add ins run
officecli add "$FILE" '/body/p[N]' --type run --after '/body/p[N]/r[1]'   --prop revision.type=ins --prop revision.author=AI --prop text="4"

# Step 4: Add del run
officecli add "$FILE" '/body/p[N]' --type run --after '/body/p[N]/r[1]'   --prop revision.type=del --prop revision.author=AI --prop text="8"
```

## Revision properties (creation)

| Property | Values | Where |
|----------|--------|-------|
| `revision.type` | `ins`, `del`, `format`, `moveFrom`, `moveTo` | set on run/paragraph/table/row/cell/section |
| `revision.author` | Any string | paired with `revision.type` |
| `revision.date` | ISO 8601 (optional) | paired with `revision.type` |
| `revision.id` | Integer (auto if omitted) | required for moveFrom/moveTo pairs |

## Revision actions (accept/reject)

```bash
# Accept all / reject all
officecli set "$FILE" /revision --prop revision.action=accept
officecli set "$FILE" /revision --prop revision.action=reject

# Filtered accept/reject
officecli set "$FILE" '/revision[@author=AI]' --prop revision.action=accept
officecli set "$FILE" '/revision[@type=ins]' --prop revision.action=accept
officecli set "$FILE" '/revision[@id=42]' --prop revision.action=accept
```

## Verify

```bash
officecli query "$FILE" revision          # list all revisions
officecli query "$FILE" revision --json   # JSON output
```

## Preview after modifications

After completing revisions, **show the user a visual preview** so they can see revision marks rendered (red strikethrough, green underline, yellow highlight). Choose the right method for your environment:

**If your client supports displaying web pages** (browser tool, webview, iframe):
```bash
officecli view "$FILE" html -o /tmp/preview.html
# Then open /tmp/preview.html in your client's browser tool
```

**If the user is in AionUI**, the preview panel already shows live updates via `officecli watch` — no action needed. **Do NOT call `view html` in AionUI** (redundant — opens a second tab).

**If the user is in a CLI-only environment**, fall back to text summary:
```bash
officecli query "$FILE" revision          # show revision list as text
```

## Contract review pattern

```bash
FILE="contract_reviewed.docx"
cp original.docx "$FILE"
officecli open "$FILE"

# One command per change with find + revision
officecli set "$FILE" /body --prop find="3个工作日" --prop replace="5个工作日" --prop revision.author=甲方
officecli set "$FILE" /body --prop find="8小时" --prop replace="4小时" --prop revision.author=甲方
officecli set "$FILE" /body --prop find="24小时" --prop replace="12小时" --prop revision.author=甲方

officecli query "$FILE" revision
officecli close "$FILE"
```

## Critical rules

1. **Prefer `find + revision`** — one command replaces old→new with ins+del pair.
2. **Use `set` + `revision.type=...`** to wrap existing content — no need to remove/add.
3. **Never fall back to python-docx, unpack/pack, or raw XML** — v1.0.98+ handles all revision types natively.
4. **`revision.type=format` needs a real property change** — pair with `bold=true`, `font.color=...`, etc.
5. **moveFrom/moveTo must share `revision.id`** — same ID on both halves for paired move.
6. **`find` matches within a single run** — if 0 matches, use `get --depth 2` to see boundaries.
