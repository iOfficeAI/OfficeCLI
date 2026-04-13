---
name: officecli-hwpx
description: "Use this skill any time a .hwpx file is involved -- as input, output, or for analysis. This includes: creating new HWPX from scratch or from Markdown; reading, parsing, or extracting text; editing or modifying existing documents; querying document structure; validating integrity; comparing documents; working with Korean (한글) office documents. Trigger whenever the user mentions 'HWP', 'HWPX', '한글 문서', '한글 파일', 'Hancom', or references a .hwpx filename."
---

# OfficeCLI HWPX Skill

## Quick Decision

| Task | Supported? | Command |
|------|-----------|---------|
| Create new .hwpx | ✅ Yes | `officecli create file.hwpx` |
| Create from Markdown | ✅ Yes | `officecli create file.hwpx --from-markdown input.md` |
| Read / analyze .hwpx | ✅ Yes | `view text`, `annotated`, `outline`, `stats`, `html`, `markdown`, `tables`, `forms`, `objects` |
| Edit existing .hwpx | ✅ Yes | `set`, `add`, `remove`, `move`, `swap` |
| Label-based fill | ✅ Yes | `set /table/fill --prop '라벨=값'` or `--prop 'fill:라벨=값'` |
| New form field creation (`text/checkbox/dropdown`) | 🟡 Blocked | source prototype exists, but Hancom golden/manual verification and published binary parity are not closed yet |
| Form recognize | ✅ Yes | `view forms --auto` (label-value auto-detect) |
| Table map | ✅ Yes | `view tables` (2D grid + labels) |
| Markdown export | ✅ Yes | `view markdown` |
| Object finder | ✅ Yes | `view objects` (picture/field/bookmark/equation) |
| Query (expanded) | ✅ Yes | `query 'tc[text~=홍길동]'`, `:has()`, `>` combinator |
| Template merge | ✅ Yes | `merge template.hwpx out.hwpx --data '{"key":"val"}'` |
| Swap elements | ✅ Yes | `swap file.hwpx '/p[1]' '/p[2]'` |
| Column break | ✅ Yes | `add --type columnbreak --prop cols=2` |
| Watermark (image) | 🟡 Plan 98 active | `build-local/officecli` 1.0.42 기준 동작 확인. Opaque RGB 권장, 밝은 자산은 `bright=0`, `contrast=0` 권장 |
| Image anchor / floating picture | ✅ Yes | `add --type picture --prop anchor=page --prop halign=center --prop valign=middle` |
| Field types | ✅ Yes | `add --type author\|title\|lastsaveby\|filename` |
| Compare documents | ✅ Yes | `compare a.hwpx b.hwpx --mode text\|outline\|table` |
| HTML preview | ✅ Yes | `view html --browser` |
| Watch live preview | ✅ Yes | `watch file.hwpx` |
| Validate .hwpx | ✅ Yes | `validate` (9-level: ZIP, package, XML, IDRef, table, NS, BinData, field, section) |
| Raw XML | ✅ Yes | `raw`, `raw-set` |
| Open .hwp (binary) | ❌ No | Convert to .hwpx first (Hancom Office) |

---

## Binary Location

```bash
OFFICECLI="700_projects/cli-jaw/build-local/officecli"
# Build: cd 700_projects/cli-jaw/officecli && dotnet publish -c Release -r osx-arm64 -o ../build-local
```

---

## Core Commands

### Create & Import & Merge

```bash
officecli create doc.hwpx                                    # 빈 문서
officecli create doc.hwpx --from-markdown input.md           # MD→HWPX (JUSTIFY 기본)
officecli create doc.hwpx --from-markdown input.md --align left  # 왼쪽 정렬
officecli merge template.hwpx output.hwpx --data '{"이름":"홍길동"}'  # 템플릿 {{키}} 치환
officecli merge template.hwpx output.hwpx --data data.json           # JSON 파일 데이터
```

### View Modes

```bash
officecli view doc.hwpx text                    # 줄번호 텍스트
officecli view doc.hwpx annotated               # 경로+스타일 상세
officecli view doc.hwpx outline                 # 제목만
officecli view doc.hwpx stats                   # 문서 통계
officecli view doc.hwpx html --browser          # A4 HTML 미리보기
officecli view doc.hwpx markdown                # GFM 마크다운 변환
officecli view doc.hwpx tables                  # 테이블 2D 그리드 + 라벨 맵
officecli view doc.hwpx forms --auto            # CLICK_HERE + label-value 자동 인식
officecli view doc.hwpx forms --auto --json     # AI 파이프라인용 JSON
officecli view doc.hwpx objects                 # picture/field/bookmark/equation 목록
officecli view doc.hwpx objects --object-type field  # 특정 타입 필터
officecli view doc.hwpx styles                  # charPr/paraPr 스타일
officecli view doc.hwpx issues                  # 9-level 검증 이슈
```

### Edit

```bash
officecli add doc.hwpx /section[1] --type paragraph --prop text="내용" --prop fontsize=11
officecli add doc.hwpx /section[1] --type table --prop rows=3 --prop cols=4
officecli set doc.hwpx '/section[1]/p[1]' --prop bold=true --prop align=CENTER
officecli set doc.hwpx / --prop find="old" --prop replace="new"
officecli remove doc.hwpx /section[1]/p[3]
```

### Image watermark

```bash
700_projects/cli-jaw/build-local/officecli add doc.hwpx /section[1] \
  --type watermark \
  --prop src=/path/to/watermark.png \
  --prop bright=0 \
  --prop contrast=0
```

Validation notes:
- Hancom에서 `v5`, `v5.1`, `v5.2` 세 변형 모두 표시 확인
- 실패 원인은 XML 불일치가 아니라 **raster 특성 + watermark filter 조합**이었다
- 투명 PNG는 피하고, **opaque RGB PNG**를 우선 사용
- 매우 밝은/단순한 자산은 기본 `bright=70`, `contrast=-50`에서 희미해질 수 있음
- 설치된 `~/.local/bin/officecli`가 `Unsupported element type: watermark`를 반환하면 최신 `build-local/officecli` 사용 또는 재설치

### Image anchor / floating picture

```bash
# 기본: inline (글자처럼 취급)
officecli add doc.hwpx /section[1] --type picture --prop path=/path/to/image.png

# 페이지 기준 정중앙
officecli add doc.hwpx /section[1] --type picture \
  --prop path=/path/to/image.png \
  --prop anchor=page \
  --prop halign=center \
  --prop valign=middle \
  --prop width=10000 \
  --prop height=5000

# 페이지 기준 중앙에서 약간 이동
officecli add doc.hwpx /section[1] --type picture \
  --prop path=/path/to/image.png \
  --prop anchor=page \
  --prop halign=center \
  --prop valign=middle \
  --prop x=1200 \
  --prop y=800

# 문단 기준 floating
officecli add doc.hwpx /section[1] --type picture \
  --prop path=/path/to/image.png \
  --prop anchor=para \
  --prop wrap=square \
  --prop halign=center \
  --prop y=1200

# 글 뒤로
officecli add doc.hwpx /section[1] --type picture \
  --prop path=/path/to/image.png \
  --prop wrap=behind

# 생성 후 위치/잠금 조정
officecli set doc.hwpx '/section[1]/p[2]/run[1]/pic[1]' \
  --prop x=1111 --prop y=2222 --prop lock=1 --prop wrap=topbottom
```

Rules:
- `path`가 기본이며 `src`도 허용됨
- `anchor=page`는 **용지 전체(PAPER)** 기준 offset 계산
- `halign`/`valign`은 별도 정렬 enum이 아니라 `horzOffset`/`vertOffset` 계산으로 처리됨
- `anchor=para`는 V1에서 본문 폭 기준 가로 배치 + `y` explicit only
- set 경로는 현재 `x`, `y`, `lock`, `wrap=topbottom`까지만 문서화한다
- picture path는 `'/section[1]/p[N]/run[1]/pic[1]'` 형태를 사용

### Label Fill (테이블 자동 채우기)

```bash
officecli set doc.hwpx / --prop 'fill:대표자=홍길동' --prop 'fill:연락처=010-1234'
officecli set doc.hwpx / --prop 'fill:주소>down=서울시'   # 방향: right(기본), down, left, up
officecli set doc.hwpx /table/fill --prop '이름=김서준'    # fill: prefix 생략
```

### Query (확장 문법)

```bash
officecli query doc.hwpx 'p'                          # 모든 단락
officecli query doc.hwpx 'tc[text~=홍길동]'           # 셀 텍스트 검색
officecli query doc.hwpx 'run[bold=true]'              # 굵은 글씨
officecli query doc.hwpx 'p:has(tbl)'                  # 테이블 포함 단락
officecli query doc.hwpx 'tbl > tr > tc[colSpan!=1]'   # 병합 셀
officecli query doc.hwpx 'run[fontsize>=20]'           # 20pt 이상
officecli query doc.hwpx 'p[heading=1]'                # heading 1
```

Operators: `=`, `!=`, `~=` (contains), `>=`, `<=`
Pseudo: `:empty`, `:contains(text)`, `:has(child)`, `:first`, `:last`
Virtual attrs: `text`, `bold`, `italic`, `fontsize`, `colSpan`, `rowSpan`, `heading`

### Compare

```bash
officecli compare a.hwpx b.hwpx                    # text diff (기본)
officecli compare a.hwpx b.hwpx --mode outline      # heading diff
officecli compare a.hwpx b.hwpx --mode table --json  # table diff JSON
```

### Watch

```bash
officecli watch doc.hwpx           # 파일 변경 시 HTML 자동 갱신
officecli unwatch doc.hwpx         # 중지
```

### Validate

```bash
officecli validate doc.hwpx
```

9-level: ZIP integrity, package (mimetype/rootfile/version), XML, IDRef, table structure, namespace, BinData orphan, field pairs, section count.

---

## Key Workflows

### 1. AI 양식 자동 채우기 (recognize → fill)

```bash
officecli view form.hwpx forms --auto --json > fields.json  # Step 1: 인식
# Step 2: AI가 label→value 매핑
officecli set form.hwpx /table/fill --prop '성 명=홍길동'   # Step 3: 채우기
```

> **Regulation docs** (운영지침 등): `forms --auto`로 label-value 인식 가능하지만,
> 체크박스 계층(□→○→-→*), 별첨 양식 참조, 비표준 heading은 Python 패턴매칭 필요.
> See "Document Classification & Pattern-Match Editing" section above.

### 2. 테이블 구조 파악 → 편집

```bash
officecli view doc.hwpx tables                              # 2D 그리드 맵
officecli query doc.hwpx 'tc[text~=대표자]'                # 셀 검색
officecli set doc.hwpx /table/fill --prop '대표자=홍길동'   # label fill
```

### 3. Markdown 왕복 변환

```bash
officecli view doc.hwpx markdown > output.md                # HWPX→MD
officecli create new.hwpx --from-markdown output.md         # MD→HWPX
```

### 4. 템플릿 대량 문서 생성

```bash
# 템플릿에 {{키}} 플레이스홀더 넣고 → 데이터로 치환
officecli merge template.hwpx 홍길동.hwpx --data '{"이름":"홍길동","날짜":"2026-04-12"}'
officecli merge template.hwpx 이지은.hwpx --data '{"이름":"이지은","날짜":"2026-04-12"}'
# 테이블 안의 {{키}}도 치환됨. 미해결 키는 보고됨.
```

### 5. 문서 비교

```bash
officecli compare before.hwpx after.hwpx --mode text
officecli compare before.hwpx after.hwpx --json > diff.json
```

---

## Document Classification & Pattern-Match Editing (Plan 90.999 + 99.7)

officecli `view forms --auto` handles standard label-value detection. For **complex templates**
(KICE exams, regulation docs, checkbox hierarchies), use the Python pattern-match fallback:

### Document Types (auto-classified)

| Type | Key Signals | Example |
|------|------------|---------|
| `exam` | equation 10+, rect objects | KICE 수능/모의고사 시험지 |
| `form` | table 3+, checkboxes (□/■) | 대학 신청서, 정부 양식 |
| `regulation` | ○ bullets 10+, 별첨/조항 refs, table 10+ | 운영지침, 내규, 시행세칙 |
| `report` | long text, few tables | 보고서, 논문 |
| `mixed` | none of above | 사업계획서 |

### Regulation-Specific Patterns (new)

- **Checkbox hierarchy**: `□` (section) → `○` (item) → `-` (detail) → `*` (footnote)
- **Appendix references**: `[별첨 제N호]`, `[별지 N]` — linked to form templates
- **Digit-concatenated headings**: `"3지원금 집행기준"` (no space between number and title)
- **Uniform footer**: repeated identical footers → org extraction (e.g., "크림슨창업지원단장 귀하")

### Verified Edit Workflow (lineseg strip)

For direct XML editing outside officecli, strip ALL linesegarray → edit → repack:
```bash
# Python one-liner
python -c "
import zipfile, re
from lxml import etree
# ... strip linesegarray, edit <t> nodes, repack ZIP
"
```
**Verified on**: KICE exam (193 lineseg), application form (472p), regulation doc (599 lineseg, HWP→HWPX).
All opened correctly in Hancom after full lineseg strip + text edits.

### 25 Regex Patterns Available (R1-R25)

Key patterns: lineseg strip (R1), checkbox (R6), label detect (R7-R8), uniform space normalization (R10),
checkbox hierarchy (R21), appendix ref (R22), digit-title concat (R23).
Full inventory → `devlog/_plan/office/hwp/plan/99.7-kice-regex-parsing-implementation.md`.

---

## Common Pitfalls

| Pitfall | Correct Approach |
|---------|-----------------|
| `--props text=Hello` | `--prop text=Hello` — 반드시 singular `--prop` |
| `/body/p[1]` path | HWPX는 `/section[1]/p[1]` — body가 아닌 section 기반 |
| `.hwp` (binary) 열기 | `.hwpx`로 변환 필수 |
| Unquoted `[N]` in shell | `"/section[1]/p[1]"` — 반드시 따옴표 |
| fontsize 미지정 | `--prop fontsize=11` 항상 명시 — charPr 오염 방지 |
| `--type formfield`를 build-local이 못 알아봄 | source tree prototype이 있어도 release acceptance 전까지는 blocked로 취급 |
| 테이블 수동 매핑 | `view tables` 한 줄로 대체 가능 |

---

## Essential Rules

1. **View mode 필수** — `officecli view file.hwpx` 만으로는 에러; `text`/`markdown`/`tables` 등 지정
2. **경로 1-based** — `/section[1]/p[1]`
3. **경로 따옴표** — shell glob 방지
4. **`--prop` singular** — `--props` 아님
5. **fontsize 항상 명시** — charPr 0 오염 방지
6. **편집 후 검증** — `view issues` + `validate` (9-level 동일 검사 범위)
7. **한글 자동 정규화** — PUA 제거, 균등 분배 축소 자동 적용
8. **Transport parity** — CLI/Resident/MCP 모두 같은 view 모드 지원 (tables, markdown, objects, forms)
