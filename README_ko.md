# OfficeCLI

> **OfficeCLI는 세계 최초이자 최고의, AI 에이전트를 위해 설계된 Office 도구입니다.**

**모든 AI 에이전트에게 Word, Excel, PowerPoint의 완전한 제어권을 — 단 한 줄의 코드로.**

오픈소스. 단일 바이너리. Office 설치 불필요. 의존성 제로. 모든 플랫폼 지원.

[![GitHub Release](https://img.shields.io/github/v/release/iOfficeAI/OfficeCLI)](https://github.com/iOfficeAI/OfficeCLI/releases)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](LICENSE)

[English](README.md) | [中文](README_zh.md) | [日本語](README_ja.md) | **한국어**

<p align="center">
  <img src="assets/ppt-process.gif" alt="AionUi에서 OfficeCLI로 PPT 제작 과정" width="100%">
</p>

<p align="center"><em><a href="https://github.com/iOfficeAI/AionUi">AionUi</a>에서 OfficeCLI로 PPT 제작 과정</em></p>

<p align="center"><strong>PowerPoint 프레젠테이션</strong></p>

<table>
<tr>
<td width="33%"><img src="assets/designwhatmovesyou.gif" alt="OfficeCLI 디자인 프레젠테이션 (PowerPoint)"></td>
<td width="33%"><img src="assets/horizon.gif" alt="OfficeCLI 비즈니스 프레젠테이션 (PowerPoint)"></td>
<td width="33%"><img src="assets/efforless.gif" alt="OfficeCLI 테크 프레젠테이션 (PowerPoint)"></td>
</tr>
<tr>
<td width="33%"><img src="assets/blackhole.gif" alt="OfficeCLI 우주 프레젠테이션 (PowerPoint)"></td>
<td width="33%"><img src="assets/first-ppt-aionui.gif" alt="OfficeCLI 게임 프레젠테이션 (PowerPoint)"></td>
<td width="33%"><img src="assets/shiba.gif" alt="OfficeCLI 크리에이티브 프레젠테이션 (PowerPoint)"></td>
</tr>
</table>

<p align="center">—</p>
<p align="center"><strong>Word 문서</strong></p>

<table>
<tr>
<td width="33%"><img src="assets/showcase/word1.gif" alt="OfficeCLI 학술 논문 (Word)"></td>
<td width="33%"><img src="assets/showcase/word2.gif" alt="OfficeCLI 프로젝트 제안서 (Word)"></td>
<td width="33%"><img src="assets/showcase/word3.gif" alt="OfficeCLI 연간 보고서 (Word)"></td>
</tr>
</table>

<p align="center">—</p>
<p align="center"><strong>Excel 스프레드시트</strong></p>

<table>
<tr>
<td width="33%"><img src="assets/showcase/excel1.gif" alt="OfficeCLI 예산 관리 (Excel)"></td>
<td width="33%"><img src="assets/showcase/excel2.gif" alt="OfficeCLI 성적 관리 (Excel)"></td>
<td width="33%"><img src="assets/showcase/excel3.gif" alt="OfficeCLI 매출 대시보드 (Excel)"></td>
</tr>
</table>

<p align="center"><em>위의 모든 문서는 AI 에이전트가 OfficeCLI를 사용하여 완전 자동으로 생성 — 템플릿 없음, 수동 편집 없음.</em></p>

## AI 에이전트용 — 한 줄로 시작

이 한 줄을 AI 에이전트 채팅에 붙여넣기만 하면 — 스킬 파일을 자동으로 읽고 설치를 완료합니다:

```
curl -fsSL https://officecli.ai/SKILL.md
```

이게 전부입니다. 스킬 파일이 에이전트에게 바이너리 설치 방법과 모든 명령어 사용법을 알려줍니다.

> **기술 세부사항:** OfficeCLI에는 [SKILL.md](SKILL.md)(239줄, 약 8K 토큰)가 포함되어 있으며, 명령어 구문, 아키텍처, 자주 발생하는 실수를 다룹니다. 설치 후 에이전트는 즉시 Office 문서를 생성, 읽기, 수정할 수 있습니다.

## 일반 사용자용 — AionUi를 설치하여 체험

명령어를 작성하고 싶지 않다면 [**AionUi**](https://github.com/iOfficeAI/AionUi)를 설치하세요 — 자연어로 Office 문서를 만들고 편집할 수 있는 데스크톱 앱입니다. 내부적으로 OfficeCLI가 구동됩니다.

원하는 것을 설명하기만 하면 AionUi가 모든 것을 처리합니다.

## 개발자용 — 30초 만에 라이브로 확인

[설치](#설치) 후:

```bash
# 1. 빈 PowerPoint 생성
officecli create deck.pptx

# 2. 라이브 미리보기 시작 — 브라우저에서 http://localhost:26315 이 열립니다
officecli watch deck.pptx --port 26315

# 3. 다른 터미널을 열고 슬라이드 추가 — 브라우저가 즉시 업데이트됩니다
officecli add deck.pptx / --type slide --prop title="Hello, World!"
```

이게 전부입니다. `add`, `set`, `remove` 명령을 실행할 때마다 미리보기가 실시간으로 갱신됩니다. 계속 실험해 보세요 — 브라우저가 바로 여러분의 라이브 피드백 루프입니다.

## 빠른 시작

```bash
# 프레젠테이션을 생성하고 콘텐츠 추가
officecli create deck.pptx
officecli add deck.pptx / --type slide --prop title="Q4 Report" --prop background=1A1A2E
officecli add deck.pptx /slide[1] --type shape \
  --prop text="Revenue grew 25%" --prop x=2cm --prop y=5cm \
  --prop font=Arial --prop size=24 --prop color=FFFFFF

# 개요 보기
officecli view deck.pptx outline
# → Slide 1: Q4 Report
# →   Shape 1 [TextBox]: Revenue grew 25%

# HTML로 보기 — 서버 없이 브라우저에서 렌더링된 미리보기를 엽니다
officecli view deck.pptx html

# 모든 요소의 구조화된 JSON 가져오기
officecli get deck.pptx /slide[1]/shape[1] --json
```

```json
{
  "tag": "shape",
  "path": "/slide[1]/shape[1]",
  "attributes": {
    "name": "TextBox 1",
    "text": "Revenue grew 25%",
    "x": "720000",
    "y": "1800000"
  }
}
```

## 왜 OfficeCLI인가?

이전에는 50줄의 Python과 3개의 라이브러리가 필요했습니다:

```python
from pptx import Presentation
from pptx.util import Inches, Pt
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])
title = slide.shapes.title
title.text = "Q4 Report"
# ... 45줄 더 ...
prs.save('deck.pptx')
```

이제 명령어 하나면 됩니다:

```bash
officecli add deck.pptx / --type slide --prop title="Q4 Report"
```

**OfficeCLI로 할 수 있는 것:**

- **생성** 문서 -- 빈 문서 또는 콘텐츠 포함
- **읽기** 텍스트, 구조, 스타일, 수식 -- 일반 텍스트 또는 구조화된 JSON
- **분석** 서식 문제, 스타일 불일치, 구조적 결함
- **수정** 모든 요소 -- 텍스트, 글꼴, 색상, 레이아웃, 수식, 차트, 이미지
- **재구성** 콘텐츠 -- 요소 추가, 삭제, 이동, 문서 간 복사

| 형식 | 읽기 | 수정 | 생성 |
|------|------|------|------|
| Word (.docx) | ✅ | ✅ | ✅ |
| Excel (.xlsx) | ✅ | ✅ | ✅ |
| PowerPoint (.pptx) | ✅ | ✅ | ✅ |

**Word** — [단락](https://github.com/iOfficeAI/OfficeCLI/wiki/word-paragraph), [런](https://github.com/iOfficeAI/OfficeCLI/wiki/word-run), [표](https://github.com/iOfficeAI/OfficeCLI/wiki/word-table), [스타일](https://github.com/iOfficeAI/OfficeCLI/wiki/word-style), [머리글/바닥글](https://github.com/iOfficeAI/OfficeCLI/wiki/word-header-footer), [이미지](https://github.com/iOfficeAI/OfficeCLI/wiki/word-picture), [수식](https://github.com/iOfficeAI/OfficeCLI/wiki/word-equation), [메모](https://github.com/iOfficeAI/OfficeCLI/wiki/word-comment), [각주](https://github.com/iOfficeAI/OfficeCLI/wiki/word-footnote), [워터마크](https://github.com/iOfficeAI/OfficeCLI/wiki/word-watermark), [북마크](https://github.com/iOfficeAI/OfficeCLI/wiki/word-bookmark), [목차](https://github.com/iOfficeAI/OfficeCLI/wiki/word-toc), [차트](https://github.com/iOfficeAI/OfficeCLI/wiki/word-chart), [하이퍼링크](https://github.com/iOfficeAI/OfficeCLI/wiki/word-hyperlink), [섹션](https://github.com/iOfficeAI/OfficeCLI/wiki/word-section), [양식 필드](https://github.com/iOfficeAI/OfficeCLI/wiki/word-formfield), [콘텐츠 컨트롤 (SDT)](https://github.com/iOfficeAI/OfficeCLI/wiki/word-sdt), [필드](https://github.com/iOfficeAI/OfficeCLI/wiki/word-field), [문서 속성](https://github.com/iOfficeAI/OfficeCLI/wiki/word-document)

**Excel** — [셀](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-cell), 수식(150개 이상의 내장 함수 자동 계산), [시트](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-sheet), [테이블](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-table), [조건부 서식](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-conditionalformatting), [차트](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-chart), [피벗 테이블](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-pivottable), [이름 범위](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-namedrange), [데이터 유효성 검사](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-validation), [이미지](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-picture), [스파크라인](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-sparkline), [메모](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-comment), [자동 필터](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-autofilter), [도형](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-shape), CSV/TSV 가져오기, `$Sheet:A1` 셀 주소 지정

**PowerPoint** — [슬라이드](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-slide), [도형](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-shape), [이미지](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-picture), [표](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-table), [차트](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-chart), [애니메이션](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-slide), [모프 전환](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-morph-check), [3D 모델 (.glb)](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-3dmodel), [슬라이드 줌](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-zoom), [수식](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-equation), [테마](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-theme), [연결선](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-connector), [비디오/오디오](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-video), [그룹](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-group), [노트](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-notes), [플레이스홀더](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-placeholder)

## 사용 사례

**개발자용:**
- 데이터베이스나 API에서 보고서 자동 생성
- 문서 일괄 처리(일괄 검색/교체, 스타일 업데이트)
- CI/CD 환경에서 문서 파이프라인 구축(테스트 결과에서 문서 생성)
- Docker/컨테이너 환경에서의 헤드리스 Office 자동화

**AI 에이전트용:**
- 사용자 프롬프트에서 프레젠테이션 생성(위 예시 참조)
- 문서에서 구조화된 데이터를 JSON으로 추출
- 납품 전 문서 품질 검증

**팀용:**
- 문서 템플릿을 복제하고 데이터 입력
- CI/CD 파이프라인에서 자동 문서 검증

## 설치

단일 자체 완결형 바이너리로 제공. .NET 런타임 내장 -- 설치할 것도, 관리할 런타임도 없습니다.

**원라인 설치:**

```bash
# macOS / Linux
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash

# Windows (PowerShell)
irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
```

**또는 수동 다운로드** [GitHub Releases](https://github.com/iOfficeAI/OfficeCLI/releases):

| 플랫폼 | 바이너리 |
|--------|---------|
| macOS Apple Silicon | `officecli-mac-arm64` |
| macOS Intel | `officecli-mac-x64` |
| Linux x64 | `officecli-linux-x64` |
| Linux ARM64 | `officecli-linux-arm64` |
| Windows x64 | `officecli-win-x64.exe` |
| Windows ARM64 | `officecli-win-arm64.exe` |

설치 확인: `officecli --version`

**또는 다운로드한 바이너리에서 셀프 설치:**

```bash
officecli install
```

업데이트는 백그라운드에서 자동 확인됩니다. `officecli config autoUpdate false`로 비활성화하거나 `OFFICECLI_SKIP_UPDATE=1`로 단일 실행 시 건너뛸 수 있습니다. 설정은 `~/.officecli/config.json`에 있습니다.

## 주요 기능

### 라이브 미리보기

`watch`는 로컬 HTTP 서버를 시작하여 PowerPoint 파일의 라이브 HTML 미리보기를 제공합니다. 수정할 때마다 브라우저가 자동 새로고침 — AI 에이전트와의 반복 디자인에 최적입니다.

```bash
officecli watch deck.pptx
# http://localhost:26315 열기 — set/add/remove 시마다 자동 새로고침
```

도형, 차트, 수식, 3D 모델(Three.js), 모프 전환, 줌 내비게이션, 모든 도형 효과를 렌더링합니다.

### 레지던트 모드와 배치

다단계 워크플로우에서 레지던트 모드는 문서를 메모리에 유지합니다. 배치 모드는 한 번의 open/save 사이클에서 여러 작업을 실행합니다.

```bash
# 레지던트 모드 — 명명된 파이프로 거의 제로 지연
officecli open report.docx
officecli set report.docx /body/p[1]/r[1] --prop bold=true
officecli set report.docx /body/p[2]/r[1] --prop color=FF0000
officecli close report.docx

# 배치 모드 — 원자적 다중 명령 실행
echo '[{"command":"set","path":"/slide[1]/shape[1]","props":{"text":"Hello"}},
      {"command":"set","path":"/slide[1]/shape[2]","props":{"fill":"FF0000"}}]' \
  | officecli batch deck.pptx --json
```

### 3계층 아키텍처

간단하게 시작하고, 필요할 때만 깊이 들어가세요.

| 레이어 | 용도 | 명령어 |
|--------|------|--------|
| **L1: 읽기** | 콘텐츠의 시맨틱 뷰 | `view` (text, annotated, outline, stats, issues, html) |
| **L2: DOM** | 구조화된 요소 작업 | `get`, `query`, `set`, `add`, `remove`, `move` |
| **L3: 원시 XML** | XPath 직접 접근 — 범용 폴백 | `raw`, `raw-set`, `add-part`, `validate` |

```bash
# L1 — 고수준 뷰
officecli view report.docx annotated
officecli view budget.xlsx text --cols A,B,C --max-lines 50

# L2 — 요소 수준 작업
officecli query report.docx "run:contains(TODO)"
officecli add budget.xlsx / --type sheet --prop name="Q2 Report"
officecli move report.docx /body/p[5] --to /body --index 1

# L3 — L2로 부족할 때 원시 XML
officecli raw deck.pptx /slide[1]
officecli raw-set report.docx document \
  --xpath "//w:p[1]" --action append \
  --xml '<w:r><w:t>Injected text</w:t></w:r>'
```

## AI 통합

### MCP 서버

내장 [MCP](https://modelcontextprotocol.io) 서버 — 명령어 하나로 등록:

```bash
officecli mcp claude       # Claude Code
officecli mcp cursor       # Cursor
officecli mcp vscode       # VS Code / Copilot
officecli mcp lmstudio     # LM Studio
officecli mcp list         # 등록 상태 확인
```

JSON-RPC로 모든 문서 작업을 제공 — 셸 접근 불필요.

### 직접 CLI 통합

2단계로 OfficeCLI를 모든 AI 에이전트에 통합:

1. **바이너리 설치** -- 명령어 하나 ([설치](#설치) 참조)
2. **완료.** OfficeCLI가 AI 도구(Claude Code, GitHub Copilot, Codex)를 자동 감지하고, 알려진 설정 디렉토리를 확인하여 스킬 파일을 설치합니다. 에이전트는 즉시 Office 문서를 생성, 읽기, 수정할 수 있습니다.

<details>
<summary><strong>수동 설정 (선택사항)</strong></summary>

자동 설치가 환경을 지원하지 않는 경우, 스킬 파일을 수동으로 설치할 수 있습니다:

**SKILL.md를 에이전트에 직접 제공:**

```bash
curl -fsSL https://officecli.ai/SKILL.md
```

**Claude Code 로컬 스킬로 설치:**

```bash
curl -fsSL https://officecli.ai/SKILL.md -o ~/.claude/skills/officecli.md
```

**기타 에이전트:** `SKILL.md`(239줄, 약 8K 토큰)의 내용을 에이전트의 시스템 프롬프트 또는 도구 설명에 포함하세요.

</details>

**모든 언어에서 호출:**

```python
# Python
import subprocess, json
def cli(*args): return subprocess.check_output(["officecli", *args], text=True)
cli("create", "deck.pptx")
cli("set", "deck.pptx", "/slide[1]/shape[1]", "--prop", "text=Hello")
```

```js
// JavaScript
const { execFileSync } = require('child_process')
const cli = (...args) => execFileSync('officecli', args, { encoding: 'utf8' })
cli('set', 'deck.pptx', '/slide[1]/shape[1]', '--prop', 'text=Hello')
```

모든 명령어가 `--json`으로 구조화된 출력을 지원합니다. 경로 기반 주소 지정으로 에이전트가 XML 네임스페이스를 이해할 필요가 없습니다.

### 에이전트가 OfficeCLI를 선호하는 이유

- **결정론적 JSON 출력** -- 모든 명령어가 `--json`을 지원하며, 일관된 스키마의 구조화된 데이터를 반환. 정규식 파싱 불필요.
- **경로 기반 주소 지정** -- 모든 요소가 안정적인 경로를 가짐(`/slide[1]/shape[2]`). XML 네임스페이스를 이해하지 않고도 문서 탐색 가능. 참고: 경로는 OfficeCLI 고유 구문(1부터 시작하는 인덱스, 요소 로컬 이름)을 사용하며, XPath가 아닙니다.
- **단계적 복잡성** -- L1(읽기)에서 시작, L2(수정)로 확대, 필요할 때만 L3(원시 XML)로 폴백. 토큰 소비 최소화.
- **자가 치유 워크플로우** -- `validate`, `view issues`, 도움말 시스템으로 에이전트가 사람의 개입 없이 문제를 감지하고 자체 수정 가능.
- **내장 도움말** -- 속성 이름이나 값 형식이 불확실할 때 `officecli <format> set <element>`을 실행하여 확인. 추측 불필요.
- **자동 설치** -- 스킬 파일 수동 설정 불필요. OfficeCLI가 AI 도구를 자동 감지하고 설정 완료.

### 내장 도움말

속성 이름을 모를 때, 계층형 도움말로 확인:

```bash
officecli pptx set              # 모든 설정 가능한 요소와 속성
officecli pptx set shape        # 특정 요소 유형의 세부사항
officecli pptx set shape.fill   # 단일 속성 형식과 예시
officecli docx query            # 셀렉터 참조: 속성, :contains, :has() 등
```

`pptx`를 `docx`나 `xlsx`로 대체 가능. 동사는 `view`, `get`, `query`, `set`, `add`, `raw`.

`officecli --help`로 전체 개요 확인.

### JSON 출력 스키마

모든 명령어가 `--json`을 지원합니다. 일반적인 응답 형식:

**단일 요소** (`get --json`):

```json
{"tag": "shape", "path": "/slide[1]/shape[1]", "attributes": {"name": "TextBox 1", "text": "Hello"}}
```

**요소 목록** (`query --json`):

```json
[
  {"tag": "paragraph", "path": "/body/p[1]", "attributes": {"style": "Heading1", "text": "Title"}},
  {"tag": "paragraph", "path": "/body/p[5]", "attributes": {"style": "Heading1", "text": "Summary"}}
]
```

**오류**는 구조화된 오류 객체를 반환합니다. 오류 코드, 수정 제안, 사용 가능한 값을 포함:

```json
{
  "success": false,
  "error": {
    "error": "Slide 50 not found (total: 8)",
    "code": "not_found",
    "suggestion": "Valid Slide index range: 1-8"
  }
}
```

오류 코드: `not_found`, `invalid_value`, `unsupported_property`, `invalid_path`, `unsupported_type`, `missing_property`, `file_not_found`, `file_locked`, `invalid_selector`. 속성 이름은 자동 교정 지원 -- 속성 이름 오타 시 가장 근접한 매칭을 제안합니다.

**오류 복구** -- 에이전트가 사용 가능한 요소를 확인하여 자체 수정:

```bash
# 에이전트가 잘못된 경로 시도
officecli get report.docx /body/p[99] --json
# 반환: {"success": false, "error": {"error": "...", "code": "not_found", "suggestion": "..."}}

# 에이전트가 사용 가능한 요소를 확인하여 자체 수정
officecli get report.docx /body --depth 1 --json
# 사용 가능한 하위 요소 목록 반환, 에이전트가 올바른 경로 선택
```

**변경 확인** (`set`, `add`, `remove`, `move`, `create`에서 `--json` 사용 시):

```json
{"success": true, "path": "/slide[1]/shape[1]"}
```

`officecli --help`로 종료 코드와 오류 형식의 전체 설명 확인.

## 비교

| | OfficeCLI | Microsoft Office | LibreOffice | python-docx / openpyxl |
|---|---|---|---|---|
| 오픈소스 & 무료 | ✓ (Apache 2.0) | ✗ (유료 라이선스) | ✓ | ✓ |
| AI 네이티브 CLI + JSON | ✓ | ✗ | ✗ | ✗ |
| 제로 설치 (단일 바이너리) | ✓ | ✗ | ✗ | ✗ (Python + pip 필요) |
| 모든 언어에서 호출 | ✓ (CLI) | ✗ (COM/Add-in) | ✗ (UNO API) | Python만 |
| 경로 기반 요소 접근 | ✓ | ✗ | ✗ | ✗ |
| 원시 XML 폴백 | ✓ | ✗ | ✗ | 부분 지원 |
| 라이브 미리보기 | ✓ | ✓ | ✗ | ✗ |
| 헤드리스 / CI | ✓ | ✗ | 부분 지원 | ✓ |
| 크로스 플랫폼 | ✓ | Windows/Mac | ✓ | ✓ |
| Word + Excel + PowerPoint | ✓ | ✓ | ✓ | 여러 라이브러리 필요 |

## 업데이트 및 설정

```bash
officecli config autoUpdate false              # 자동 업데이트 확인 비활성화
OFFICECLI_SKIP_UPDATE=1 officecli ...          # 단일 실행 시 확인 건너뛰기 (CI)
```

## 명령어 참조

| 명령어 | 설명 |
|--------|------|
| [`create`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-create) | 빈 .docx, .xlsx, .pptx 생성 (확장자로 유형 결정) |
| [`view`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-view) | 콘텐츠 보기 (모드: `outline`, `text`, `annotated`, `stats`, `issues`, `html`) |
| [`get`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-get) | 요소와 하위 요소 가져오기 (`--depth N`, `--json`) |
| [`query`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-query) | CSS 스타일 쿼리 (`[attr=value]`, `:contains()`, `:has()` 등) |
| [`set`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-set) | 요소 속성 수정 |
| [`add`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-add) | 요소 추가 (또는 `--from <path>`로 복제) |
| [`remove`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-remove) | 요소 삭제 |
| [`move`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-move) | 요소 이동 (`--to <parent> --index N`) |
| [`swap`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-swap) | 두 요소 교체 |
| [`validate`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-validate) | OpenXML 스키마 검증 |
| [`batch`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-batch) | 한 번의 open/save 사이클에서 여러 작업 실행 (stdin, `--input`, 또는 `--commands`) |
| [`merge`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-merge) | 템플릿 병합 — `{{key}}` 플레이스홀더를 JSON 데이터로 교체 |
| [`watch`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-watch) | 브라우저에서 라이브 HTML 미리보기, 자동 새로고침 |
| [`mcp`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-mcp) | AI 도구 통합용 MCP 서버 시작 |
| [`raw`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-raw) | 문서 파트의 원시 XML 보기 |
| [`raw-set`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-raw) | XPath로 원시 XML 수정 |
| `add-part` | 새 문서 파트 추가 (머리글, 차트 등) |
| [`open`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-open) | 레지던트 모드 시작 (문서를 메모리에 유지) |
| `close` | 저장하고 레지던트 모드 종료 |
| [`install`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-install) | 바이너리 + 스킬 + MCP 설치 (`all`, `claude`, `cursor` 등) |
| `config` | 설정 가져오기 또는 변경 |
| `<format> <command>` | [내장 도움말](https://github.com/iOfficeAI/OfficeCLI/wiki/command-reference) (예: `officecli pptx set shape`) |

## 엔드투엔드 워크플로우 예시

전형적인 에이전트 자가 치유 워크플로우: 프레젠테이션 생성, 콘텐츠 입력, 검증, 문제 수정 -- 모두 사람의 개입 없이.

```bash
# 1. 생성
officecli create report.pptx

# 2. 콘텐츠 추가
officecli add report.pptx / --type slide --prop title="Q4 Results"
officecli add report.pptx /slide[1] --type shape \
  --prop text="Revenue: $4.2M" --prop x=2cm --prop y=5cm --prop size=28
officecli add report.pptx / --type slide --prop title="Details"
officecli add report.pptx /slide[2] --type shape \
  --prop text="Growth driven by new markets" --prop x=2cm --prop y=5cm

# 3. 검증
officecli view report.pptx outline
officecli validate report.pptx

# 4. 문제 수정
officecli view report.pptx issues --json
# 출력에 따라 문제 수정:
officecli set report.pptx /slide[1]/shape[1] --prop font=Arial
```

### 템플릿 병합

문서 내 `{{key}}` 플레이스홀더를 JSON 데이터로 교체 -- 단락, 표 셀, 도형, 머리글, 바닥글, 차트 제목 등 모든 텍스트 콘텐츠 지원.

```bash
# 인라인 JSON 데이터
officecli merge template.docx output.docx '{"name":"Alice","dept":"Sales","date":"2026-03-30"}'

# JSON 파일에서 읽기
officecli merge template.pptx report.pptx data.json

# Excel 템플릿
officecli merge budget-template.xlsx q4-budget.xlsx '{"quarter":"Q4","year":"2026"}'
```

### 단위와 색상

모든 치수 및 색상 속성은 유연한 입력 형식을 지원:

| 유형 | 지원 형식 | 예시 |
|------|----------|------|
| **치수** | cm, in, pt, px 또는 원시 EMU | `2cm`, `1in`, `72pt`, `96px`, `914400` |
| **색상** | 16진수, 색상 이름, RGB, 테마 색상 | `#FF0000`, `FF0000`, `red`, `rgb(255,0,0)`, `accent1` |
| **글꼴 크기** | 숫자만 또는 pt 접미사 | `14`, `14pt`, `10.5pt` |
| **간격** | pt, cm, in 또는 배율 | `12pt`, `0.5cm`, `1.5x`, `150%` |

## 자주 사용하는 패턴

```bash
# Word 문서의 모든 Heading1 텍스트 교체
officecli query report.docx "paragraph[style=Heading1]" --json | ...
officecli set report.docx /body/p[1]/r[1] --prop text="New Title"

# 모든 슬라이드 콘텐츠를 JSON으로 내보내기
officecli get deck.pptx / --depth 2 --json

# Excel 셀 일괄 업데이트
officecli batch budget.xlsx --input updates.json --json

# CSV 데이터를 Excel 시트로 가져오기
officecli add budget.xlsx / --type sheet --prop name="Q1 Data" --prop csv=sales.csv

# 템플릿 병합으로 보고서 일괄 생성
officecli merge invoice-template.docx invoice-001.docx '{"client":"Acme","total":"$5,200"}'

# 납품 전 문서 품질 확인
officecli validate report.docx && officecli view report.docx issues --json
```

## 문서

[Wiki](https://github.com/iOfficeAI/OfficeCLI/wiki)에서 모든 명령어, 요소 유형, 속성의 상세 가이드를 확인하세요:

- **형식별:** [Word](https://github.com/iOfficeAI/OfficeCLI/wiki/word-reference) | [Excel](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-reference) | [PowerPoint](https://github.com/iOfficeAI/OfficeCLI/wiki/powerpoint-reference)
- **워크플로우:** [엔드투엔드 예시](https://github.com/iOfficeAI/OfficeCLI/wiki/workflows) -- Word 보고서, Excel 대시보드, PPT 프레젠테이션, 일괄 수정, 레지던트 모드
- **문제 해결:** [자주 발생하는 오류와 해결책](https://github.com/iOfficeAI/OfficeCLI/wiki/troubleshooting)
- **AI 에이전트 가이드:** [Wiki 내비게이션 결정 트리](https://github.com/iOfficeAI/OfficeCLI/wiki/agent-guide)

## 소스에서 빌드

컴파일에는 [.NET 10 SDK](https://dotnet.microsoft.com/download)가 필요합니다. 출력은 자체 완결형 네이티브 바이너리 -- .NET이 내장되어 있어 실행 시 설치 불필요.

```bash
./build.sh
```

## 라이선스

[Apache License 2.0](LICENSE)

버그 리포트와 기여는 [GitHub Issues](https://github.com/iOfficeAI/OfficeCLI/issues)로 환영합니다.

---

OfficeCLI가 유용하다면 [GitHub에서 스타를 눌러주세요](https://github.com/iOfficeAI/OfficeCLI) — 더 많은 사람들이 프로젝트를 발견하는 데 도움이 됩니다.

[OfficeCLI.AI](https://OfficeCLI.AI) | [GitHub](https://github.com/iOfficeAI/OfficeCLI)
