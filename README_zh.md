# OfficeCLI

[![GitHub Release](https://img.shields.io/github/v/release/iOfficeAI/OfficeCLI)](https://github.com/iOfficeAI/OfficeCLI/releases)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](LICENSE)

[English](README.md) | **中文**

**全球首款专为 AI 智能体打造的 Office 办公软件。**

**让 AI 智能体通过命令行处理一切 Office 文档。**

OfficeCLI 是一个免费、开源的命令行工具，专为 AI 智能体设计，可读取、编辑和自动化处理 Word、Excel 和 PowerPoint 文件。单一可执行文件，无需安装 Microsoft Office、WPS 或任何运行时依赖。

> 为智能体而生，人类亦可用。

<p align="center">
  <img src="assets/ppt-process.gif" alt="在 AionUi 上使用 OfficeCLI 的 PPT 制作过程" width="100%">
</p>

<p align="center"><em>在 <a href="https://github.com/iOfficeAI/AionUi">AionUi</a> 上使用 OfficeCLI 的 PPT 制作过程</em></p>

## 安装

OfficeCLI 是单一可执行文件 — 无运行时依赖。

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

安装后设置 AI 智能体集成（见下方 [AI 集成](#ai-集成)）：

```powershell
officecli skills all       # 为所有检测到的 AI 客户端安装技能文件
```

## 快速开始

```bash
# 创建文档
officecli create report.docx
officecli create budget.xlsx
officecli create deck.pptx

# 查看内容
officecli view report.docx text
officecli view deck.pptx outline
officecli view budget.xlsx issues --json      # 检查格式问题

# 读取元素
officecli get budget.xlsx /Sheet1/B5 --json
officecli get budget.xlsx '$Sheet1:A1:D10'    # Excel 单元格范围记法

# 用类 CSS 选择器查找元素
officecli query report.docx "paragraph[style=Heading1]"
officecli query deck.pptx "shape[fill=FF0000]"

# 修改内容
officecli set report.docx /body/p[1]/r[1] --prop text="Updated Title" --prop bold=true
officecli set budget.xlsx '$Sheet1:B5' --prop value=42 --prop bold=true
officecli set deck.pptx /slide[1]/shape[1] --prop text="New Title" --prop color=FF6600

# 添加元素
officecli add report.docx /body --type paragraph --prop text="New paragraph" --index 3
officecli add deck.pptx / --type slide
officecli add deck.pptx /slide[2] --type shape --prop preset=star5 --prop fill=FFD700

# 实时预览 — 每次修改自动刷新
officecli watch deck.pptx
```

## 内置帮助

不确定属性名时，用分层帮助查询：

```bash
officecli pptx set              # 全部可设置元素与属性
officecli pptx set shape        # 某一类元素的详细说明
officecli pptx set shape.fill   # 单个属性格式与示例
officecli docx query            # 选择器说明：属性匹配、:contains、:has() 等
```

将 `pptx` 换成 `docx` 或 `xlsx`；动词包括 `view`、`get`、`query`、`set`、`add`、`raw`。

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

## 支持的格式

| 格式 | 读取 | 修改 | 创建 |
|------|------|------|------|
| Word (.docx) | ✓ | ✓ | ✓ |
| Excel (.xlsx) | ✓ | ✓ | ✓ |
| PowerPoint (.pptx) | ✓ | ✓ | ✓ |

**Word** — 段落、文本片段、表格、样式、页眉/页脚、图片、公式、批注、列表、水印、书签、目录

**Excel** — 单元格、公式、工作表、样式、条件格式、图表、数据透视表、命名范围、数据验证、`$Sheet:A1` 单元格寻址

**PowerPoint** — 幻灯片、形状、文本框、图片、表格、图表、动画、morph 过渡、3D 模型（.glb）、幻灯片缩放、公式、主题、连接线、视频/音频

## AI 集成

OfficeCLI 提供三种方式接入 AI 智能体：

### 1. 技能安装（推荐 CLI 智能体使用）

安装技能定义，让智能体自动发现 OfficeCLI 能力：

```bash
officecli skills all       # 自动检测并安装到所有客户端
officecli skills claude    # Claude Code
officecli skills copilot   # GitHub Copilot
officecli skills codex     # OpenAI Codex
officecli skills cursor    # Cursor
officecli skills windsurf  # Windsurf
```

也可以直接提供技能文件：`curl -fsSL https://officecli.ai/SKILL.md`

### 2. MCP 服务器（协议级集成）

内置 [MCP](https://modelcontextprotocol.io) 服务器 — 一条命令注册：

```bash
officecli mcp claude       # Claude Code
officecli mcp cursor       # Cursor
officecli mcp vscode       # VS Code / Copilot
officecli mcp lmstudio     # LM Studio
officecli mcp list         # 查看注册状态
```

通过 JSON-RPC 暴露所有文档操作 — 无需 shell 访问。

### 3. 直接 CLI 调用（任意语言）

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

## 构建

本地编译需要安装 [.NET 10 SDK](https://dotnet.microsoft.com/download)。在仓库根目录执行：

```bash
./build.sh
```

## 许可证

[Apache License 2.0](LICENSE)

## 友情链接

[LINUX DO - 新的理想型社区](https://linux.do/)

---

[OfficeCLI.AI](https://OfficeCLI.AI)
