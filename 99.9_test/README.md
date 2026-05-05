# 99.9 Phase A-H Manual Test Guide

> Run from: `cd /Users/jun/Developer/new/700_projects/cli-jaw/officecli`
> CLI: `dotnet run --project src/officecli -- <command>`
> Python: `python3 scripts/hwpx_form_edit.py <command>`

---

## Phase E — Security

### E1: Path Traversal
```bash
# Validate catches no errors on clean file
dotnet run --project src/officecli -- validate 99.9_test/test_phase_e.hwpx
```
**Pass**: No `path_traversal` errors

### E2: ZIP Bomb
**Pass**: Opening any test file completes without timeout or OOM

### E5: XXE Defense
**Pass**: `validate` does not throw on any test file (XXE would cause exception)

---

## Phase A — Quick Wins

### A1: Extended Keywords
```bash
dotnet run --project src/officecli -- view 99.9_test/test_phase_a.hwpx forms
```
**Pass**: All 8 keywords (성명, 주소, 생년월일, 전화번호, 이메일, 직업, 학력, 자격증) appear in form field recognition

### A3: False Positive Filter
```bash
dotnet run --project src/officecli -- view 99.9_test/test_phase_a.hwpx text
```
**Pass**: "접수시간: 10:30" — the time value "10:30" is NOT stripped (time-related labels preserve values)

### A4: Shape Alt Text
```bash
dotnet run --project src/officecli -- view 99.9_test/test_phase_a.hwpx text
```
**Pass**: "면적: 100m²" renders with superscript stripped cleanly

---

## Phase B — Form Enhancement

### B1-B2: In-Cell & KV Table
```bash
dotnet run --project src/officecli -- view 99.9_test/test_phase_b.hwpx forms
```
**Pass**: Table recognized as form with fields: 성명, 생년월일, 주소, 전화번호, 이메일, 비고

### B5: Checkbox Recognition
```bash
dotnet run --project src/officecli -- view 99.9_test/test_phase_b.hwpx text
```
**Pass**: "□남 □여" and "□동의 □미동의" detected as checkbox fields

### B6: Fill Test (Optional)
```bash
dotnet run --project src/officecli -- set 99.9_test/test_phase_b.hwpx fill --props '성명=홍길동'
dotnet run --project src/officecli -- view 99.9_test/test_phase_b.hwpx forms
```
**Pass**: 성명 field now shows "홍길동"

---

## Phase F — Text Quality

### F1: PUA Strip
```bash
dotnet run --project src/officecli -- view 99.9_test/test_phase_f.hwpx text
```
**Pass**: No PUA characters (U+E000-U+F8FF range) in output

### F3: Legal Heading Detection
```bash
dotnet run --project src/officecli -- view 99.9_test/test_phase_f.hwpx outline
```
**Pass**: "별표 1" appears as heading in outline
**Pass**: "별첨 서류는 반환하지 않습니다" does NOT appear as heading (it's body text)

### F5: 1x1 Cell
```bash
dotnet run --project src/officecli -- view 99.9_test/test_phase_f.hwpx markdown
```
**Pass**: Single-cell table rendered as structured text, not markdown table syntax

### F7: Phone Spacing (Python)
```bash
python3 scripts/hwpx_form_edit.py extract 99.9_test/test_phase_f.hwpx
```
**Pass**: "0 1 0 - 1 2 3 4 - 5 6 7 8" collapsed to "010-1234-5678"

---

## Phase G — Parser

### G1: Section File Regex
**Pass**: All test files open successfully (section discovery works)

### G3: Legal Heading Detection
```bash
dotnet run --project src/officecli -- view 99.9_test/test_phase_g.hwpx outline
```
**Pass**: "제1장 총칙" appears as h1 heading
**Pass**: "제2절 적용범위" appears as h2 heading
**Pass**: "제1장에서 언급한 바와 같이" does NOT appear as heading

### G4: Dublin Core Metadata
```bash
dotnet run --project src/officecli -- get 99.9_test/test_phase_g.hwpx /metadata
```
**Pass**: Returns metadata dict (may include dc:title, dc:creator if present in file)

### G5: MIME Validation
```bash
dotnet run --project src/officecli -- validate 99.9_test/test_phase_g.hwpx
```
**Pass**: No `package_mimetype_invalid` error

### G7: Markdown Import
```bash
# Create a test file and import markdown
dotnet run --project src/officecli -- create 99.9_test/test_import.hwpx --type hwpx
dotnet run --project src/officecli -- import 99.9_test/test_import.hwpx --markdown "# Heading\n\n> Quote text\n\n- List item 1\n- List item 2\n\n1. Ordered 1\n2. Ordered 2"
dotnet run --project src/officecli -- view 99.9_test/test_import.hwpx text
```
**Pass**: Heading, quote (with > prefix), list items all present in output

---

## Phase H — Diff/Compare

### H1: Text Compare
```bash
dotnet run --project src/officecli -- compare 99.9_test/test_phase_h_a.hwpx 99.9_test/test_phase_h_b.hwpx text
```
**Pass**: Output shows:
- "첫 번째 문장입니다" — `unchanged`
- "두 번째 문장입니다" → "두 번째 문장이 수정되었습니다" — `modified`
- "세 번째 문장입니다" — `unchanged`
- "네 번째 문장이 추가되었습니다" — `added`

### H5: Page Range Compare
```bash
dotnet run --project src/officecli -- compare 99.9_test/test_phase_h_a.hwpx 99.9_test/test_phase_h_b.hwpx text --pages "1"
```
**Pass**: Returns diff filtered to section 1 only

---

## Summary Checklist

| Phase | Items | Test File | Key Check |
|-------|-------|-----------|-----------|
| E | E1-E6 | test_phase_e.hwpx | `validate` clean |
| A | A1-A4 | test_phase_a.hwpx | `view forms` shows 8 keywords |
| B | B1-B7 | test_phase_b.hwpx | `view forms` shows table fields |
| F | F1-F8 | test_phase_f.hwpx | heading detection, 1x1 cell, phone spacing |
| G | G1-G7 | test_phase_g.hwpx | heading h1/h2, MIME check, import |
| H | H1-H5 | test_phase_h_{a,b}.hwpx | compare shows unchanged/modified/added |

### Build Status
```bash
cd officecli && dotnet build && dotnet test
```
**Expected**: 0 errors, 189 tests pass, 2 pre-existing failures (Plan703 tests)
