# OfficeCLI

[![GitHub Release](https://img.shields.io/github/v/release/iOfficeAI/OfficeCLI)](https://github.com/iOfficeAI/OfficeCLI/releases)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](LICENSE)

[English](README.md) | **中文**

**全球首款专为 AI 智能体打造的 Office 办公软件。**

**让 AI 智能体通过命令行处理一切 Office 文档。**

OfficeCLI 是一个免费、开源的命令行工具，专为 AI 智能体设计，可读取、编辑和自动化处理 Word、Excel 和 PowerPoint 文件。单一可执行文件，无需安装 Microsoft Office、WPS 或任何运行时依赖。

> 为智能体而生，人类亦可用。

<video src="assets/ppt-processs.mp4" poster="assets/ppt-process.png" autoplay loop muted playsinline width="100%"></video>

<p align="center"><em>在 <a href="https://github.com/iOfficeAI/AionUi">AionUi</a> 上使用 OfficeCLI 的 PPT 制作过程</em></p>

## AI 智能体接入

OfficeCLI 附带 [SKILL.md](SKILL.md)，用于指导 AI 智能体高效使用本工具。

首先让你的智能体读取此文件：

```bash
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/SKILL.md
```

如果你的智能体支持本地技能安装，建议安装到本地：

**Claude Code：**

```bash
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/SKILL.md -o ~/.claude/skills/officecli.md
```

**其他智能体：**

将 `SKILL.md` 的内容添加到智能体的系统提示词或工具描述中。

然后安装 CLI 二进制文件：

```bash
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
```

Windows (PowerShell)：

```powershell
irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
```

## 为什么选择 OfficeCLI？

AI 智能体擅长处理文本，但 Office 文档是 XML 的二进制封装。OfficeCLI 弥合了这一鸿沟，让智能体能够：

- **创建** 文档 — 空白文档或带内容的文档
- **读取** 文本、结构、样式、公式 — 纯文本或结构化 JSON
- **分析** 格式问题、样式不一致和结构问题
- **修改** 任意元素 — 文本、字体、颜色、布局、公式、图表、图片
- **重组** 内容 — 添加、删除、移动、复制跨文档元素

全部通过简单的 CLI 命令完成，支持结构化 JSON 输出，无需安装 Office。

## 安装

OfficeCLI 是单一可执行文件 — 无运行时依赖。一条命令即可安装：

**macOS / Linux：**

```bash
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
```

**Windows (PowerShell)：**

```powershell
irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
```

也可以从 [GitHub Releases](https://github.com/iOfficeAI/OfficeCLI/releases) 手动下载。

## 快速开始

```bash
# 创建空白文档
officecli create report.docx
officecli create budget.xlsx
officecli create deck.pptx

# 查看文档内容
officecli view report.docx text

# 检查格式问题
officecli view report.docx issues --json

# 读取指定单元格
officecli get budget.xlsx /Sheet1/B5 --json

# 修改内容
officecli set report.docx /body/p[1]/r[1] --prop text="Updated Title" --prop bold=true

# 驻留模式批量编辑（文档保持在内存中）
officecli open presentation.pptx
officecli set presentation.pptx /slide[1]/shape[1] --prop text="New Title"
officecli set presentation.pptx /slide[2]/shape[3] --prop text="New Subtitle"
officecli close presentation.pptx
```

## 内置帮助

属性名、取值格式不确定时，请用分层帮助查询，不要凭感觉写。将下文中的 `pptx` 换成 `docx` 或 `xlsx`；动词包括 `view`、`get`、`query`、`set`、`add`、`raw`。

```bash
officecli pptx set              # 全部可设置元素与属性
officecli pptx set shape        # 某一类元素的详细说明
officecli pptx set shape.fill   # 单个属性格式与示例
```

运行 `officecli --help` 可看到相同说明。若构建时包含 `wiki/` 目录，部分帮助内容可能嵌入二进制。

## 三层架构

OfficeCLI 采用渐进式复杂度模型 — 从简单开始，仅在需要时深入。

### L1：读取与检查

文档内容的高级语义视图。

```bash
# Word — 带行号的纯文本
officecli view report.docx text

# Word — 带格式标注的文本
officecli view report.docx annotated

# Excel — 按列筛选查看
officecli view budget.xlsx text --cols A,B,C --max-lines 50

# Excel — 检测公式和样式问题
officecli view budget.xlsx issues --json

# PowerPoint — 所有幻灯片大纲
officecli view deck.pptx outline

# PowerPoint — 字体和样式统计
officecli view deck.pptx stats
```

### L2：DOM 操作

通过结构化元素路径和属性修改文档。

```bash
# Word — 查询标题并设置格式（类 CSS 选择器；完整语法见帮助）
officecli query report.docx "paragraph[style=Heading1]"
officecli docx query            # 选择器说明：属性匹配、:contains、:has() 等
officecli set report.docx /body/p[1]/r[1] --prop bold=true --prop color=FF0000

# Word — 添加段落、删除段落
officecli add report.docx /body --type paragraph --prop text="New paragraph" --index 3
officecli remove report.docx /body/p[5]

# Excel — 读取和修改单元格
officecli get budget.xlsx /Sheet1/B5 --json
officecli set budget.xlsx /Sheet1/A1 --prop formula="=SUM(A2:A10)" --prop numFmt="0.00%"

# Excel — 新建工作表、添加行
officecli add budget.xlsx / --type sheet --prop name="Q2 Report"
officecli add budget.xlsx /Sheet1 --type row --prop values="Name,Amount,Date"

# PowerPoint — 修改幻灯片内容
officecli set deck.pptx /slide[1]/shape[1] --prop text="New Title"
officecli set deck.pptx /slide[2]/shape[3] --prop fontSize=24 --prop bold=true

# PowerPoint — 添加幻灯片、从其他幻灯片复制形状
officecli add deck.pptx / --type slide
officecli add deck.pptx /slide[3] --from /slide[1]/shape[2]

# 移动元素
officecli move report.docx /body/p[5] --to /body --index 1
```

### L3：原始 XML

通过 XPath 直接访问 XML — 任何 OpenXML 操作的通用兜底方案。

```bash
# Word — 查看和修改原始 XML
officecli raw report.docx document
officecli raw-set report.docx document \
  --xpath "//w:p[1]" \
  --action append \
  --xml '<w:r><w:t>Injected text</w:t></w:r>'

# Word — 添加页眉
officecli add-part report.docx /body --type header

# Excel — 查看工作表原始 XML
officecli raw budget.xlsx /Sheet1

# Excel — 向工作表添加图表
officecli add-part budget.xlsx /Sheet1 --type chart

# PowerPoint — 查看幻灯片原始 XML
officecli raw deck.pptx /slide[1]

# 验证任意文档
officecli validate report.docx
officecli validate budget.xlsx
```

## 支持的格式

| 格式 | 读取 | 修改 | 创建 |
|------|------|------|------|
| Word (.docx) | ✓ | ✓ | ✓ |
| Excel (.xlsx) | ✓ | ✓ | ✓ |
| PowerPoint (.pptx) | ✓ | ✓ | ✓ |

### Word — 段落、文本片段、表格、样式、页眉/页脚、图片、公式、批注、列表

### Excel — 单元格、公式、工作表、样式（字体、填充、边框、数字格式）、条件格式、图表

### PowerPoint — 幻灯片、形状、文本框、图片、动画、公式、主题、对齐与形状效果

## 驻留模式

对于多步骤工作流，驻留模式将文档保持在后台进程中，避免每次命令都重新加载文件。

```bash
officecli open report.docx        # 启动驻留进程
officecli view report.docx text   # 即时响应 — 无需重新加载
officecli set report.docx ...     # 即时响应 — 无需重新加载
officecli close report.docx       # 保存并退出
```

通过命名管道通信，命令间延迟接近零。

## 批量模式（batch）

在**一次**打开/保存周期内执行多条命令（若文档已由 `open` 驻留，则会转发到驻留进程）。通过标准输入或 `--input` 传入 JSON 数组。

```bash
echo '[{"command":"view","mode":"outline"},{"command":"get","path":"/slide[1]","depth":1}]' \
  | officecli batch deck.pptx
```

使用 `--stop-on-error` 可在首次失败时中止。每条命令的 `command` 支持 `get`、`query`、`set`、`add`、`remove`、`move`、`view`、`raw`、`raw-set`、`validate` 等（完整字段见源码中的 `BatchItem`）。

## 更新与配置

CLI 可能在后台**非阻塞**地检查更新并与 GitHub 最新 release 对齐。配置位于 `~/.officecli/config.json`。

- **关闭自动更新检查：** `officecli config autoUpdate false`（查看当前值：`officecli config autoUpdate`）。
- **单次调用跳过后台检查（如 CI）：** `OFFICECLI_SKIP_UPDATE=1 officecli ...`

仍可通过 `install.sh` / `install.ps1` 安装或升级二进制文件。

## Python 使用示例

```python
import subprocess, json

def cli(*args): return subprocess.check_output(["officecli", *args], text=True)
def cli_json(*args): return json.loads(cli(*args, "--json"))

cli("create", "deck.pptx")
cli("set", "deck.pptx", "/slide[1]/shape[1]", "--prop", "text=Hello")
shapes = cli_json("query", "deck.pptx", "shape")
```

## JavaScript 使用示例

```js
const { execFileSync } = require('child_process')

const cli = (...args) => execFileSync('officecli', args, { encoding: 'utf8' })
const cliJson = (...args) => JSON.parse(cli(...args, '--json'))

cli('create', 'deck.pptx')
cli('set', 'deck.pptx', '/slide[1]/shape[1]', '--prop', 'text=Hello')
const shapes = cliJson('query', 'deck.pptx', 'shape')
```

## AI 智能体集成

### 为什么 AI 智能体应该使用 OfficeCLI？

**确定性 JSON 输出** — 每个命令都支持 `--json`，返回结构一致的数据。无需正则解析。

**完善的验证与诊断** — `validate`、`view issues`、`raw-set` 等命令帮助智能体检测问题，并在修改后验证文档正确性。

**基于路径的寻址** — 每个文档中的每个元素都有稳定的路径。智能体无需理解 XML 命名空间即可导航文档。

**渐进式复杂度** — 智能体从 L1（读取）开始，升级到 L2（修改），仅在必要时回退到 L3（原始 XML）。这在保持所有操作可能的同时，最大限度减少 token 消耗。

## 对比

OfficeCLI 与其他 AI 智能体处理 Office 文档的方案相比如何？

| | OfficeCLI | Microsoft Office | LibreOffice | python-docx / openpyxl |
|---|---|---|---|---|
| 开源免费 | ✓ (Apache 2.0) | ✗（付费授权） | ✓ | ✓ |
| AI 友好的 CLI | ✓ | ✗ | 部分支持 | ✗ |
| 结构化 JSON 输出 | ✓ | ✗ | ✗ | ✗ |
| 零安装（单一可执行文件） | ✓ | ✗ | ✗ | ✗（需要 Python + pip） |
| 任意语言调用 | ✓ (CLI) | ✗ (COM/Add-in) | ✗ (UNO API) | ✗（仅 Python） |
| 基于路径的元素访问 | ✓ | ✗ | ✗ | ✗ |
| 原始 XML 兜底 | ✓ | ✗ | ✗ | 部分支持 |
| 驻留模式（内存常驻） | ✓ | ✗ | ✗ | ✗ |
| 支持无头/CI 环境 | ✓ | ✗ | 部分支持 | ✓ |
| 跨平台 | ✓ | ✗（Windows/Mac） | ✓ | ✓ |
| Word + Excel + PowerPoint | ✓ | ✓ | ✓ | 需要多个库 |
| 读取 + 写入 + 创建 | ✓ | ✓ | ✓ | ✓ |

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
