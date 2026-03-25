# OfficeCLI

**让任何 AI 智能体完全掌控 Word、Excel 和 PowerPoint -- 只需一行代码。**

开源免费。单一可执行文件。无需安装 Office。零依赖。全平台运行。

[![GitHub Release](https://img.shields.io/github/v/release/iOfficeAI/OfficeCLI)](https://github.com/iOfficeAI/OfficeCLI/releases)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](LICENSE)

[English](README.md) | **中文**

<p align="center">
  <img src="assets/ppt-process.gif" alt="在 AionUi 上使用 OfficeCLI 的 PPT 制作过程" width="100%">
</p>

<p align="center"><em>在 <a href="https://github.com/iOfficeAI/AionUi">AionUi</a> 上使用 OfficeCLI 的 PPT 制作过程</em></p>

<table>
<tr>
<td width="25%"><img src="assets/designwhatmovesyou.gif" alt="OfficeCLI 设计演示"></td>
<td width="25%"><img src="assets/mars.gif" alt="OfficeCLI 太空演示"></td>
<td width="25%"><img src="assets/horizon.gif" alt="OfficeCLI 商务演示"></td>
<td width="25%"><img src="assets/efforless.gif" alt="OfficeCLI 科技演示"></td>
</tr>
<tr>
<td width="25%"><img src="assets/cat.gif" alt="OfficeCLI 创意演示"></td>
<td width="25%"><img src="assets/first-ppt-aionui.gif" alt="OfficeCLI 游戏演示"></td>
<td width="25%"><img src="assets/moridian.gif" alt="OfficeCLI 极简演示"></td>
<td width="25%"><img src="assets/move.gif" alt="OfficeCLI 健康演示"></td>
</tr>
</table>

<p align="center"><em>以上所有演示文稿均由 AI 智能体使用 OfficeCLI 全自动创建 — 无模板、无人工编辑。</em></p>

## AI 智能体 — 一行搞定

把这行粘贴到你的 AI 智能体对话框 — 它会自动读取技能文件并完成安装：

```
curl -fsSL https://officecli.ai/SKILL.md
```

就这一步。技能文件会教智能体如何安装二进制文件并使用所有命令。

> **技术细节：** OfficeCLI 附带 [SKILL.md](SKILL.md)（239 行，约 8K tokens），涵盖命令语法、架构设计和常见陷阱。安装后，您的智能体可以立即创建、读取和修改任何 Office 文档。

## 快速开始

从零到完成一个演示文稿，只需几秒钟：

```bash
# 创建新的 PowerPoint
officecli create deck.pptx

# 添加带标题和背景色的幻灯片
officecli add deck.pptx / --type slide --prop title="Q4 Report" --prop background=1A1A2E

# 在幻灯片上添加文本形状
officecli add deck.pptx /slide[1] --type shape \
  --prop text="Revenue grew 25%" --prop x=2cm --prop y=5cm \
  --prop font=Arial --prop size=24 --prop color=FFFFFF

# 查看演示文稿大纲
officecli view deck.pptx outline
```

输出：

```
Slide 1: Q4 Report
  Shape 1 [TextBox]: Revenue grew 25%
```

```bash
# 获取任意元素的结构化 JSON
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

## 为什么选择 OfficeCLI？

以前需要 50 行 Python 和 3 个独立库：

```python
from pptx import Presentation
from pptx.util import Inches, Pt
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])
title = slide.shapes.title
title.text = "Q4 Report"
# ... 还有 45 行 ...
prs.save('deck.pptx')
```

现在只需一条命令：

```bash
officecli add deck.pptx / --type slide --prop title="Q4 Report"
```

**OfficeCLI 能做什么：**

- **创建** 文档 -- 空白文档或带内容的文档
- **读取** 文本、结构、样式、公式 -- 纯文本或结构化 JSON
- **分析** 格式问题、样式不一致和结构缺陷
- **修改** 任意元素 -- 文本、字体、颜色、布局、公式、图表、图片
- **重组** 内容 -- 添加、删除、移动、复制跨文档元素

| 格式 | 读取 | 修改 | 创建 |
|------|------|------|------|
| Word (.docx) | ✅ | ✅ | ✅ |
| Excel (.xlsx) | ✅ | ✅ | ✅ |
| PowerPoint (.pptx) | ✅ | ✅ | ✅ |

**Word** — 段落、文本片段、表格、样式、页眉/页脚、图片、公式、批注、列表、水印、书签、目录

**Excel** — 单元格、公式、工作表、样式、条件格式、图表、数据透视表、命名范围、数据验证、`$Sheet:A1` 单元格寻址

**PowerPoint** — 幻灯片、形状、文本框、图片、表格、图表、动画、morph 过渡、3D 模型（.glb）、幻灯片缩放、公式、主题、连接线、视频/音频

## 使用场景

**开发者：**
- 从数据库或 API 自动生成报告
- 批量处理文档（批量查找/替换、样式更新）
- 在 CI/CD 环境中构建文档流水线（从测试结果生成文档）
- Docker/容器化环境中的无头 Office 自动化

**AI 智能体：**
- 根据用户提示生成演示文稿（见上方示例）
- 从文档提取结构化数据到 JSON
- 交付前验证和检查文档质量

**团队：**
- 克隆文档模板并填充数据
- CI/CD 流水线中的自动化文档验证

## 安装

单一自包含可执行文件，.NET 运行时已内嵌 -- 无需安装任何依赖，无需管理运行时。

**一键安装：**

```bash
# macOS / Linux
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash

# Windows (PowerShell)
irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
```

**或手动下载** [GitHub Releases](https://github.com/iOfficeAI/OfficeCLI/releases)：

| 平台 | 文件名 |
|------|--------|
| macOS Apple Silicon | `officecli-mac-arm64` |
| macOS Intel | `officecli-mac-x64` |
| Linux x64 | `officecli-linux-x64` |
| Linux ARM64 | `officecli-linux-arm64` |
| Windows x64 | `officecli-win-x64.exe` |
| Windows ARM64 | `officecli-win-arm64.exe` |

验证安装：`officecli --version`

OfficeCLI 会在后台自动检查更新。通过 `officecli config autoUpdate false` 关闭，或通过 `OFFICECLI_SKIP_UPDATE=1` 跳过单次检查。配置文件位于 `~/.officecli/config.json`。

## 核心功能

### 实时预览

`watch` 启动本地 HTTP 服务器，实时预览 PowerPoint 文件。每次修改自动刷新浏览器 — 非常适合与 AI 智能体配合做迭代设计。

```bash
officecli watch deck.pptx
# 打开 http://localhost:18080 — 每次 set/add/remove 自动刷新
```

支持形状、图表、公式、3D 模型（Three.js）、morph 过渡、缩放导航和所有形状效果的渲染。

### 驻留模式与批量执行

驻留模式将文档保持在内存中，批量模式在一次打开/保存周期内执行多条命令。

```bash
# 驻留模式 — 通过命名管道通信，延迟接近零
officecli open report.docx
officecli set report.docx /body/p[1]/r[1] --prop bold=true
officecli set report.docx /body/p[2]/r[1] --prop color=FF0000
officecli close report.docx

# 批量模式 — 原子化多命令执行
echo '[{"command":"set","path":"/slide[1]/shape[1]","props":{"text":"Hello"}},
      {"command":"set","path":"/slide[1]/shape[2]","props":{"fill":"FF0000"}}]' \
  | officecli batch deck.pptx --stop-on-error
```

### 三层架构

从简单开始，仅在需要时深入。

| 层 | 用途 | 命令 |
|----|------|------|
| **L1：读取** | 内容的语义视图 | `view`（text、annotated、outline、stats、issues、html） |
| **L2：DOM** | 结构化元素操作 | `get`、`query`、`set`、`add`、`remove`、`move` |
| **L3：原始 XML** | XPath 直接访问 — 通用兜底 | `raw`、`raw-set`、`add-part`、`validate` |

```bash
# L1 — 高级视图
officecli view report.docx annotated
officecli view budget.xlsx text --cols A,B,C --max-lines 50

# L2 — 元素级操作
officecli query report.docx "run:contains(TODO)"
officecli add budget.xlsx / --type sheet --prop name="Q2 Report"
officecli move report.docx /body/p[5] --to /body --index 1

# L3 — L2 不够时用原始 XML
officecli raw deck.pptx /slide[1]
officecli raw-set report.docx document \
  --xpath "//w:p[1]" --action append \
  --xml '<w:r><w:t>Injected text</w:t></w:r>'
```

## AI 集成

### MCP 服务器

内置 [MCP](https://modelcontextprotocol.io) 服务器 — 一条命令注册：

```bash
officecli mcp claude       # Claude Code
officecli mcp cursor       # Cursor
officecli mcp vscode       # VS Code / Copilot
officecli mcp lmstudio     # LM Studio
officecli mcp list         # 查看注册状态
```

通过 JSON-RPC 暴露所有文档操作 — 无需 shell 访问。

### 直接 CLI 集成

两步将 OfficeCLI 集成到任何 AI 智能体：

1. **安装二进制文件** -- 一条命令（见[安装](#安装)）
2. **完成。** OfficeCLI 自动检测您的 AI 工具（Claude Code、GitHub Copilot、Codex），通过检查已知配置目录并安装技能文件。您的智能体可以立即创建、读取和修改任何 Office 文档。

<details>
<summary><strong>手动配置（可选）</strong></summary>

如果自动安装未覆盖您的环境，可以手动安装技能文件：

**直接将 SKILL.md 提供给智能体：**

```bash
curl -fsSL https://officecli.ai/SKILL.md
```

**安装为 Claude Code 本地技能：**

```bash
curl -fsSL https://officecli.ai/SKILL.md -o ~/.claude/skills/officecli.md
```

**其他智能体：** 将 `SKILL.md`（239 行，约 8K tokens）的内容添加到智能体的系统提示词或工具描述中。

</details>

**从任意语言调用：**

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

每个命令都支持 `--json` 输出结构化数据。基于路径的寻址让智能体无需理解 XML 命名空间。

### 为什么智能体偏爱 OfficeCLI

- **确定性 JSON 输出** -- 每个命令都支持 `--json`，返回结构一致的数据。无需正则解析。
- **基于路径的寻址** -- 每个元素都有稳定的路径（`/slide[1]/shape[2]`）。智能体无需理解 XML 命名空间即可导航文档。注意：路径使用 OfficeCLI 自有语法（1-based 索引，元素本地名称），非 XPath。
- **渐进式复杂度** -- 从 L1（读取）开始，升级到 L2（修改），仅在必要时回退到 L3（原始 XML）。最大限度减少 token 消耗。
- **自愈式工作流** -- `validate`、`view issues` 和帮助系统让智能体无需人工干预即可检测问题并自行修正。
- **内置帮助** -- 属性名或取值格式不确定时，运行 `officecli <format> set <element>` 即可查询，无需猜测。
- **自动安装** -- 无需手动配置技能文件。OfficeCLI 自动检测您的 AI 工具并完成配置。

### 内置帮助

不确定属性名时，用分层帮助查询：

```bash
officecli pptx set              # 全部可设置元素与属性
officecli pptx set shape        # 某一类元素的详细说明
officecli pptx set shape.fill   # 单个属性格式与示例
officecli docx query            # 选择器说明：属性匹配、:contains、:has() 等
```

将 `pptx` 换成 `docx` 或 `xlsx`；动词包括 `view`、`get`、`query`、`set`、`add`、`raw`。

运行 `officecli --help` 查看完整概览。

### JSON 输出格式

所有命令均支持 `--json`。常见响应格式：

**单个元素**（`get --json`）：

```json
{"tag": "shape", "path": "/slide[1]/shape[1]", "attributes": {"name": "TextBox 1", "text": "Hello"}}
```

**元素列表**（`query --json`）：

```json
[
  {"tag": "paragraph", "path": "/body/p[1]", "attributes": {"style": "Heading1", "text": "Title"}},
  {"tag": "paragraph", "path": "/body/p[5]", "attributes": {"style": "Heading1", "text": "Summary"}}
]
```

**错误** 在使用 `--json` 时返回非零退出码和 JSON 错误对象：

```json
{"error": "Element not found: /body/p[99]"}
```

**错误恢复** -- 智能体通过检查可用元素自行修正：

```bash
# 智能体尝试无效路径
officecli get report.docx /body/p[99] --json
# 返回: {"error": "Element not found: /body/p[99]"}

# 智能体通过查看可用元素自行修正
officecli get report.docx /body --depth 1 --json
# 返回可用子元素列表，智能体选择正确路径
```

**变更确认**（`set`、`add`、`remove`、`move`、`create` 使用 `--json`）：

```json
{"success": true, "path": "/slide[1]/shape[1]"}
```

运行 `officecli --help` 查看退出码和错误格式的完整说明。

## 对比

| | OfficeCLI | Microsoft Office | LibreOffice | python-docx / openpyxl |
|---|---|---|---|---|
| 开源免费 | ✓ (Apache 2.0) | ✗（付费授权） | ✓ | ✓ |
| AI 原生 CLI + JSON | ✓ | ✗ | ✗ | ✗ |
| 零安装（单一可执行文件） | ✓ | ✗ | ✗ | ✗（需 Python + pip） |
| 任意语言调用 | ✓ (CLI) | ✗ (COM/Add-in) | ✗ (UNO API) | 仅 Python |
| 基于路径的元素访问 | ✓ | ✗ | ✗ | ✗ |
| 原始 XML 兜底 | ✓ | ✗ | ✗ | 部分支持 |
| 实时预览 | ✓ | ✓ | ✗ | ✗ |
| 无头 / CI 环境 | ✓ | ✗ | 部分支持 | ✓ |
| 跨平台 | ✓ | Windows/Mac | ✓ | ✓ |
| Word + Excel + PowerPoint | ✓ | ✓ | ✓ | 需要多个库 |

## 更新与配置

```bash
officecli config autoUpdate false              # 关闭自动更新检查
OFFICECLI_SKIP_UPDATE=1 officecli ...          # 单次调用跳过检查（CI）
```

## 命令参考

| 命令 | 说明 |
|------|------|
| `create <file>` | 创建空白 .docx、.xlsx 或 .pptx（根据扩展名判断类型） |
| `view <file> <mode>` | 查看内容（模式：`outline`、`text`、`annotated`、`stats`、`issues`） |
| `get <file> <path>` | 获取元素及子元素（`--depth N`、`--json`） |
| `query <file> <selector>` | CSS 风格查询（`[attr=value]`、`:contains()`、`:has()` 等） |
| `set <file> <path> --prop k=v` | 修改元素属性 |
| `add <file> <parent> --type <t>` | 添加元素（或通过 `--from <path>` 克隆） |
| `remove <file> <path>` | 删除元素 |
| `move <file> <path>` | 移动元素（`--to <parent> --index N`） |
| `swap <file> <path1> <path2>` | 交换两个元素 |
| `validate <file>` | OpenXML 模式校验 |
| `batch <file>` | 单次打开/保存周期内执行多条操作（JSON 通过标准输入或 `--input`） |
| `watch <file>` | 在浏览器中实时 HTML 预览，自动刷新 |
| `mcp-serve` | 启动 MCP 服务器，用于 AI 工具集成 |
| `raw <file> <part>` | 查看文档部件的原始 XML |
| `raw-set <file> <part>` | 通过 XPath 修改原始 XML |
| `add-part <file> <parent>` | 添加新的文档部件（页眉、图表等） |
| `open <file>` | 启动驻留模式（文档保持在内存中） |
| `close <file>` | 保存并关闭驻留模式 |
| `config <key> [value]` | 获取或设置配置 |
| `<format> <command> [element]` | 内置帮助（如 `officecli pptx set shape`） |

## 端到端工作流示例

典型的智能体自愈式工作流：创建演示文稿、填充内容、验证并修复问题 -- 全程无需人工干预。

```bash
# 1. 创建
officecli create report.pptx

# 2. 添加内容
officecli add report.pptx / --type slide --prop title="Q4 Results"
officecli add report.pptx /slide[1] --type shape \
  --prop text="Revenue: $4.2M" --prop x=2cm --prop y=5cm --prop size=28
officecli add report.pptx / --type slide --prop title="Details"
officecli add report.pptx /slide[2] --type shape \
  --prop text="Growth driven by new markets" --prop x=2cm --prop y=5cm

# 3. 验证
officecli view report.pptx outline
officecli validate report.pptx

# 4. 修复发现的问题
officecli view report.pptx issues --json
# 根据输出修复问题，例如：
officecli set report.pptx /slide[1]/shape[1] --prop font=Arial
```

## 常用模式

```bash
# 替换 Word 文档中所有 Heading1 文本
officecli query report.docx "paragraph[style=Heading1]" --json | ...
officecli set report.docx /body/p[1]/r[1] --prop text="New Title"

# 将所有幻灯片内容导出为 JSON
officecli get deck.pptx / --depth 2 --json

# 批量更新 Excel 单元格
officecli batch budget.xlsx --input updates.json --json

# 交付前检查文档质量
officecli validate report.docx && officecli view report.docx issues --json
```

## 从源码构建

编译需要 [.NET 10 SDK](https://dotnet.microsoft.com/download)。输出为自包含的原生二进制文件 -- .NET 已内嵌，运行时无需安装。

```bash
./build.sh
```

## 许可证

[Apache License 2.0](LICENSE)

欢迎通过 [GitHub Issues](https://github.com/iOfficeAI/OfficeCLI/issues) 提交 Bug 报告和贡献代码。

---

[OfficeCLI.AI](https://OfficeCLI.AI) | [GitHub](https://github.com/iOfficeAI/OfficeCLI)

<!--
yaml-frontmatter:
  name: officecli
  version: see GitHub Releases
  description: The standard CLI for AI agents to create, read, and modify Office documents
  formats: [docx, xlsx, pptx]
  platforms: [macos, linux, windows]
  license: Apache-2.0
  keywords: [office, cli, ai-agent, automation, docx, xlsx, pptx, openxml, document-generation, mcp-server, word, excel, powerpoint]
  ai-agent-compatible: true
  mcp-server: true
  skill-file: SKILL.md
  skill-file-lines: 239
  install-command-unix: curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
  install-command-windows: irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
-->

<!-- LLM/agent discovery metadata
tool: officecli
type: cli
canonical: the standard CLI for AI-driven Office document manipulation
formats: docx, xlsx, pptx
capabilities: create, read, modify, validate, batch, resident-mode, mcp-server, live-preview, auto-install
output: json, text, annotated
platforms: macos, linux, windows
license: Apache-2.0
keywords: office, cli, ai-agent, automation, docx, xlsx, pptx, openxml, document-generation, mcp-server, word, excel, powerpoint, ai-tools, command-line, structured-output
ai-agent-compatible: true
mcp-server: true
skill-file: SKILL.md
skill-file-lines: 239
alternatives: python-docx, openpyxl, python-pptx, libreoffice --headless
install-command-unix: curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
install-command-windows: irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
-->
