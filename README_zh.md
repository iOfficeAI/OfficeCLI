# OfficeCLI

> **OfficeCLI 是全球首个、也是最好的专为 AI 智能体设计的 Office 套件。**

**让任何 AI 智能体完全掌控 Word、Excel 和 PowerPoint——只需一行代码。**

开源免费。单一可执行文件。无需安装 Office。零依赖。全平台运行。

**OfficeCLI 的内置 HTML 渲染引擎，高度还原文档原貌 —— 这正是让 AI 拥有"眼睛"的关键。** 它把 `.docx` / `.xlsx` / `.pptx` 渲染为 HTML 或 PNG，闭合"渲染 → 看 → 改"的循环。

[![GitHub Release](https://img.shields.io/github/v/release/iOfficeAI/OfficeCLI)](https://github.com/iOfficeAI/OfficeCLI/releases)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](LICENSE)

[English](README.md) | **中文** | [日本語](README_ja.md) | [한국어](README_ko.md)

<p align="center">
  <strong>🌐 官方网站:</strong> <a href="https://officecli.ai" target="_blank">officecli.ai</a> &nbsp;|&nbsp; <strong>💬 社区:</strong> <a href="https://discord.gg/2QAwJn7Egx" target="_blank">Discord</a>
</p>

<p align="center">
  <img src="assets/ppt-process.webp" alt="在 AionUi 上使用 OfficeCLI 的 PPT 制作过程" width="100%">
</p>

<p align="center"><em>在 <a href="https://github.com/iOfficeAI/AionUi">AionUi</a> 上使用 OfficeCLI 的 PPT 制作过程</em></p>

<p align="center"><strong>PowerPoint 演示文稿</strong></p>

<table>
<tr>
<td width="33%"><img src="assets/designwhatmovesyou.gif" alt="OfficeCLI 设计演示 (PowerPoint)"></td>
<td width="33%"><img src="assets/horizon.gif" alt="OfficeCLI 商务演示 (PowerPoint)"></td>
<td width="33%"><img src="assets/efforless.gif" alt="OfficeCLI 科技演示 (PowerPoint)"></td>
</tr>
<tr>
<td width="33%"><img src="assets/blackhole.gif" alt="OfficeCLI 太空演示 (PowerPoint)"></td>
<td width="33%"><img src="assets/first-ppt-aionui.gif" alt="OfficeCLI 游戏演示 (PowerPoint)"></td>
<td width="33%"><img src="assets/shiba.gif" alt="OfficeCLI 创意演示 (PowerPoint)"></td>
</tr>
</table>

<p align="center">—</p>
<p align="center"><strong>Word 文档</strong></p>

<table>
<tr>
<td width="33%"><img src="assets/showcase/word1.gif" alt="OfficeCLI 学术论文 (Word)"></td>
<td width="33%"><img src="assets/showcase/word2.gif" alt="OfficeCLI 项目建议书 (Word)"></td>
<td width="33%"><img src="assets/showcase/word3.gif" alt="OfficeCLI 年度报告 (Word)"></td>
</tr>
</table>

<p align="center">—</p>
<p align="center"><strong>Excel 电子表格</strong></p>

<table>
<tr>
<td width="33%"><img src="assets/showcase/excel1.gif" alt="OfficeCLI 预算跟踪 (Excel)"></td>
<td width="33%"><img src="assets/showcase/excel2.gif" alt="OfficeCLI 成绩管理 (Excel)"></td>
<td width="33%"><img src="assets/showcase/excel3.gif" alt="OfficeCLI 销售仪表盘 (Excel)"></td>
</tr>
</table>

<p align="center"><em>以上所有文档均由 AI 智能体使用 OfficeCLI 全自动创建 — 无模板、无人工编辑。</em></p>

## AI 智能体 — 一行搞定

把这行粘贴到你的 AI 智能体对话框 — 它会自动读取技能文件并完成安装：

```
curl -fsSL https://officecli.ai/SKILL.md
```

就这一步。技能文件会教智能体如何安装二进制文件并使用所有命令。

## 普通用户

**方式 A — 图形界面：** 安装 [**AionUi**](https://github.com/iOfficeAI/AionUi) — 一款桌面应用，用自然语言就能创建和编辑 Office 文档，底层由 OfficeCLI 驱动。只需描述你想要什么，AionUi 帮你搞定。

**方式 B — 命令行：** 从 [GitHub Releases](https://github.com/iOfficeAI/OfficeCLI/releases) 下载对应平台的二进制文件，然后运行：

```bash
officecli install
```

该命令会将二进制文件复制到 PATH，并自动将 **officecli 技能文件**安装到检测到的所有 AI 编程助手 — Claude Code、Cursor、Windsurf、GitHub Copilot 等。您的智能体可以立即创建、读取和编辑 Office 文档，无需额外配置。

## 开发者 — 30 秒亲眼看到效果

```bash
# 1. 安装（macOS / Linux）— 也可以：brew install officecli / npm install -g @officecli/officecli
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
# Windows (PowerShell): irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex

# 2. 创建一个空白 PowerPoint
officecli create deck.pptx

# 3. 启动实时预览 — 浏览器自动打开 http://localhost:26315
officecli watch deck.pptx

# 4. 打开另一个终端，添加一页幻灯片 — 浏览器即时刷新
officecli add deck.pptx / --type slide --prop title="Hello, World!"
```

就这么简单。你执行的每一条 `add`、`set`、`remove` 命令都会实时刷新预览。继续尝试吧 — 浏览器就是你的实时反馈窗口。

## 快速开始

```bash
# 创建演示文稿并添加内容
officecli create deck.pptx
officecli add deck.pptx / --type slide --prop title="Q4 Report" --prop background=1A1A2E
officecli add deck.pptx '/slide[1]' --type shape \
  --prop text="Revenue grew 25%" --prop x=2cm --prop y=5cm \
  --prop font=Arial --prop size=24 --prop color=FFFFFF

# 查看大纲
officecli view deck.pptx outline
# → Slide 1: Q4 Report
# →   Shape 1 [TextBox]: Revenue grew 25%

# 查看 HTML — 在浏览器中打开渲染预览，无需启动服务器
officecli view deck.pptx html

# 获取任意元素的结构化 JSON
officecli get deck.pptx '/slide[1]/shape[1]' --json
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

**Word** — 完整的 [i18n 与 RTL 支持](https://github.com/iOfficeAI/OfficeCLI/wiki/i18n)（按脚本字体槽位、按脚本 BCP-47 语言标签 `lang.latin/ea/cs`、复杂脚本粗体/斜体/字号、`direction=rtl` 在段落/文本片段/节/表格/样式/页眉/页脚/docDefaults 间级联、`rtlGutter` + `pgBorders` 简写、印地语/阿拉伯语/泰语/中日韩本地化页码）、[段落](https://github.com/iOfficeAI/OfficeCLI/wiki/word-paragraph)、[文本片段](https://github.com/iOfficeAI/OfficeCLI/wiki/word-run)、[表格](https://github.com/iOfficeAI/OfficeCLI/wiki/word-table)、[样式](https://github.com/iOfficeAI/OfficeCLI/wiki/word-style)、[页眉/页脚](https://github.com/iOfficeAI/OfficeCLI/wiki/word-header-footer)、[图片](https://github.com/iOfficeAI/OfficeCLI/wiki/word-picture)（PNG/JPG/GIF/SVG）、[公式](https://github.com/iOfficeAI/OfficeCLI/wiki/word-equation)、[批注](https://github.com/iOfficeAI/OfficeCLI/wiki/word-comment)、[脚注](https://github.com/iOfficeAI/OfficeCLI/wiki/word-footnote)、[水印](https://github.com/iOfficeAI/OfficeCLI/wiki/word-watermark)、[书签](https://github.com/iOfficeAI/OfficeCLI/wiki/word-bookmark)、[目录](https://github.com/iOfficeAI/OfficeCLI/wiki/word-toc)、[图表](https://github.com/iOfficeAI/OfficeCLI/wiki/word-chart)、[超链接](https://github.com/iOfficeAI/OfficeCLI/wiki/word-hyperlink)、[节](https://github.com/iOfficeAI/OfficeCLI/wiki/word-section)、[表单域](https://github.com/iOfficeAI/OfficeCLI/wiki/word-formfield)、[内容控件 (SDT)](https://github.com/iOfficeAI/OfficeCLI/wiki/word-sdt)、[域](https://github.com/iOfficeAI/OfficeCLI/wiki/word-field)（22 种零参数 + MERGEFIELD / REF / PAGEREF / SEQ / STYLEREF / DOCPROPERTY / IF）、[OLE 对象](https://github.com/iOfficeAI/OfficeCLI/wiki/word-ole)、[文档属性](https://github.com/iOfficeAI/OfficeCLI/wiki/word-document)

**Excel** — [单元格](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-cell)（添加时支持音标/振假名）、公式（内置 350+ 函数自动求值，可溢出的动态数组自动加 `_xlfn.` 前缀，含财务/债券与统计函数族）、[工作表](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-sheet)（visible/hidden/veryHidden、打印边距、printTitleRows/Cols、RTL `sheetView`、级联感知的工作表重命名）、[表格](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-table)、[排序](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-sort)（工作表/区域、多键、附属感知）、[条件格式](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-conditionalformatting)、[图表](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-chart)（含箱线图、[帕累托图](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-chart-add) 自动排序 + 累计百分比、对数轴）、[数据透视表](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-pivottable)（多字段、日期分组、showDataAs、排序、总计、分类汇总、紧凑/大纲/表格布局、重复项目标签、空白行、计算字段）、[切片器](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-slicer)、[命名范围](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-namedrange)、[数据验证](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-validation)、[图片](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-picture)（PNG/JPG/GIF/SVG，双重表示回退）、[迷你图](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-sparkline)、[批注](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-comment)（RTL）、[自动筛选](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-autofilter)、[形状](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-shape)、[OLE 对象](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-ole)、CSV/TSV 导入、`$Sheet:A1` 单元格寻址

**PowerPoint** — [幻灯片](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-slide)（页眉/页脚/日期/页码切换、隐藏）、[形状](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-shape)（图案填充、模糊效果、超链接提示 + 跳转幻灯片链接）、[图片](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-picture)（PNG/JPG/GIF/SVG，填充模式：stretch/contain/cover/tile，亮度/对比度/发光/阴影）、[表格](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-table)、[图表](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-chart)、[动画](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-slide)、[morph 过渡](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-morph-check)、[3D 模型（.glb）](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-3dmodel)、[幻灯片缩放](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-zoom)、[公式](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-equation)、[主题](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-theme)、[连接线](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-connector)、[视频/音频](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-video)、[组合](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-group)、[备注](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-notes)（RTL、lang）、[批注](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-comment)（RTL）、[OLE 对象](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-ole)、[占位符](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-placeholder)（按 phType 添加/设置）

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

**或通过包管理器安装：**

```bash
# Homebrew（macOS / Linux）
brew install officecli

# Scoop（Windows）
scoop install officecli

# npm（全平台 — 安装时自动拉取对应平台的原生二进制）
npm install -g @officecli/officecli
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

**或从已下载的二进制文件自安装（直接运行 `officecli` 也会触发安装）：**

```bash
officecli install    # 显式安装
officecli            # 直接运行也会触发安装
```

OfficeCLI 会在后台自动检查更新。通过 `officecli config autoUpdate false` 关闭，或通过 `OFFICECLI_SKIP_UPDATE=1` 跳过单次检查。配置文件位于 `~/.officecli/config.json`。

## 核心功能

### 内置引擎与生成原语

OfficeCLI 是自包含的。下列能力全部内置在二进制中——**无需 Office**。

#### 渲染引擎 —— 高保真、内置

OfficeCLI 的基石：一个从零实现、高保真的 HTML 渲染引擎，让 AI 智能体能"看见"渲染后的文档，而不是凭 DOM 瞎猜。它覆盖形状、图表（趋势线、误差线、瀑布、K 线、sparkline）、公式（OMML → LaTeX，KaTeX 渲染）、通过 Three.js 渲染的 3D `.glb` 模型、morph 过渡、幻灯片缩放、形状效果。按页 PNG 截图是把渲染出的 HTML 通过无头浏览器截出来的。三种模式：

- **`view html`** —— 独立 HTML 文件，资源内联。任何浏览器打开即可看。
- **`view screenshot`** —— 按页 PNG，供多模态智能体读图检查。
- **`watch`** —— 本地 HTTP 服务 + 自动刷新预览；每次 `add` / `set` / `remove` 立即更新浏览器。Excel watch 还支持单元格内联编辑、图表拖动定位。

```bash
officecli view deck.pptx html -o /tmp/deck.html
officecli view deck.pptx screenshot -o /tmp/deck.png # 多页用 --page 1-N
officecli watch deck.pptx                            # http://localhost:26315
```

> 没有可视化，生成 PPT 的智能体就是在盲跑——它能读 DOM，但分辨不出标题溢出、两个形状重叠。因为渲染引擎内置在二进制里，"渲染 → 看 → 改"循环在 CI、Docker、无显示器的服务器——只要二进制能跑的地方都能用。

#### 公式与透视引擎

350+ Excel 函数写入即自动求值——写 `=SUM(A1:A2)`，`get` 单元格，值已经在那。不需要回到 Office 重算。覆盖可溢出的动态数组（`FILTER` / `SORT` / `UNIQUE` / `SEQUENCE` / `LET` / `LAMBDA` / `MAP`，`_xlfn.` 自动加前缀）、`VLOOKUP` / `XLOOKUP` / `INDEX` / `MATCH`、财务与债券函数（`XIRR` / `PRICE` / `YIELD` / `DURATION` / `COUPNUM`）、统计分布/检验/回归（`NORM.DIST` / `T.TEST` / `LINEST`）、日期与文本函数等。

外加从源数据范围一条命令生成原生 OOXML 数据透视表——多字段行/列/筛选器、10 种聚合方式、`showDataAs` 多种模式、日期分组、计算字段、Top-N、布局选项。透视表缓存和定义都写入 OOXML，Excel 打开即看到聚合后的结果：

```bash
officecli add sales.xlsx '/Sheet1' --type pivottable \
  --prop source='Data!A1:E10000' --prop rows='Region,Category' \
  --prop cols=Quarter --prop values='Revenue:sum,Units:avg' \
  --prop showDataAs=percentOfTotal
```

#### 模板合并 —— 设计一次，填充 N 次

`merge` 把任意 `.docx` / `.xlsx` / `.pptx` 中的 `{{key}}` 占位符替换为 JSON 数据——段落、表格单元格、形状、页眉页脚、图表标题都支持。智能体一次性设计版式（昂贵），生产代码填充 N 次（廉价、确定、零 token 成本）。避免了"每份报告都从头重生成、产出 N 份版式不一致"的失败模式。

```bash
officecli merge invoice-template.docx out-001.docx --data '{"client":"Acme","total":"$5,200"}'
officecli merge q4-template.pptx q4-acme.pptx --data data.json
```

#### Dump 往返 —— 从现有文档学习

`dump` 把任意 `.docx`、`.pptx`、`.xlsx` —— 整个文档**或任意子树**（单段、单表、单页幻灯片、单个工作表、styles、numbering、theme、settings）——序列化为可重放的 batch JSON，`batch` 重放回去。给一份用户想模仿的范本，智能体读结构化规格而不是原始 OOXML XML，修改后重放。打通"我有一份现成模板"和"给我生成 100 份变体"之间的链路。

```bash
officecli dump existing.docx -o blueprint.json                  # 整个文档
officecli dump existing.docx /body/tbl[1] -o table.json         # 任意子树
officecli dump existing.xlsx /Sheet1 -o sheet.json              # 单个工作表
officecli batch new.docx --input blueprint.json
```

### 驻留模式与批量执行

驻留模式将文档保持在内存中，批量模式在一次打开/保存周期内执行多条命令。

```bash
# 驻留模式 — 通过命名管道通信，延迟接近零
officecli open report.docx
officecli set report.docx /body/p[1]/r[1] --prop bold=true
officecli set report.docx /body/p[2]/r[1] --prop color=FF0000
officecli close report.docx

# 批量模式 — 多命令执行，默认原子化：只要有一条失败，整批全部回滚
echo '[{"command":"set","path":"/slide[1]/shape[1]","props":{"text":"Hello"}},
      {"command":"set","path":"/slide[1]/shape[2]","props":{"fill":"FF0000"}}]' \
  | officecli batch deck.pptx --json

# 内联 batch，无需标准输入
officecli batch deck.pptx --commands '[{"op":"set","path":"/slide[1]/shape[1]","props":{"text":"Hi"}}]'

# 使用 --best-effort 保留已成功的部分，即使其他条目失败（v1.0.137 前的旧行为）
officecli batch deck.pptx --input updates.json --best-effort --json

# 遇到第一个失败就停止执行（默认原子模式下仍会整体回滚；需搭配 --best-effort 才会保留已成功部分）
officecli batch deck.pptx --input updates.json --stop-on-error --json
```

> **要用其他工具读这个文件?先落盘。**
> OfficeCLI 自己的读(`get`/`query`/`view`)永远能看到最新改动,所以在 officecli 内部无需操心保存。但常驻进程会延迟写盘,因此**在非-officecli 程序读取该文件之前**——python-docx/openpyxl、Microsoft Word、渲染器、交付/上传——先落盘:
> ```bash
> officecli set report.docx /body/p[1] --prop bold=true
> officecli save report.docx           # 落盘, 保留常驻进程(或 `close` = 落盘 + 释放)
> python my_reader.py report.docx      # 此时才能看到改动
> ```
> 常驻进程闲置约 10s 后也会自动落盘一次。完整落盘模型(auto-save / auto-close / save / close、环境变量调节):[wiki → open / close](https://github.com/iOfficeAI/OfficeCLI/wiki/command-open#when-the-file-on-disk-is-refreshed)。

### 三层架构

从简单开始，仅在需要时深入。

| 层 | 用途 | 命令 |
|----|------|------|
| **L1：读取** | 内容的语义视图 | `view`（text、annotated、outline、stats、issues、html、svg、screenshot） |
| **L2：DOM** | 结构化元素操作 | `get`、`query`、`set`、`add`、`remove`、`move`、`swap` |
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
officecli raw deck.pptx '/slide[1]'
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

**其他智能体：** 将 `SKILL.md` 的内容添加到智能体的系统提示词或工具描述中。

</details>

### 智能体为什么在 OfficeCLI 上如鱼得水

- **确定性 JSON 输出** —— 每条命令都支持 `--json`，schema 一致。无需正则解析、无需抓 stdout。
- **基于路径的寻址** —— 每个元素都有稳定路径（`/slide[1]/shape[2]`）。智能体无需理解 XML 命名空间即可导航文档。（OfficeCLI 自己的语法：1-based 索引、元素本地名——不是 XPath。）
- **渐进式复杂度（L1 → L2 → L3）** —— 智能体从只读视图入手，升级到 DOM 操作，仅在必要时降到 raw XML。最大限度节省 token。
- **自愈式工作流** —— `validate`、`view issues`、以及结构化错误码（`not_found`、`invalid_value`、`unsupported_property`）会返回 suggestion 和有效范围。智能体无需人工介入即可自纠错。
- **内置 agent 友好渲染引擎** —— `view html` / `view screenshot` / `watch` 原生输出 HTML 和 PNG。无需 Office。智能体能"看见"自己的产出，并在 CI / Docker / 无头环境里修复排版问题。
- **内置公式与透视引擎** —— 350+ Excel 函数写入即自动求值（含可溢出动态数组、财务/债券与统计函数族）；从源数据范围一条命令生成原生 OOXML 数据透视表。智能体立刻读到计算值和聚合结果，不需要回到 Office 重算。
- **模板合并** —— 智能体一次性设计版式，下游代码把 `{{key}}` 占位符填充 N 次。避免每份报告都烧 token 重生成。
- **Dump 往返** —— `dump` 把任意 `.docx`、`.pptx`、`.xlsx` 转成可重放的 batch JSON。智能体通过读结构化规格学习人类范本，而不是从原始 OOXML XML 反推。
- **内置帮助** —— 属性名或取值格式不确定时，智能体跑 `officecli <format> set <element>`，不靠猜。
- **自动安装** —— OfficeCLI 自动识别您的 AI 工具（Claude Code、Cursor、VS Code…）并完成配置。无需手动放 skill 文件。

### 内置帮助

不确定属性名时，用分层帮助查询：

```bash
officecli help pptx set              # 全部可设置元素与属性
officecli help pptx set shape        # 某一类元素的详细说明
officecli help docx query            # 选择器说明：属性匹配、:contains、:has() 等
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

**错误** 返回结构化错误对象，包含错误码、建议修正和可用值：

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

错误码：`not_found`、`invalid_value`、`unsupported_property`、`invalid_path`、`unsupported_type`、`missing_property`、`file_not_found`、`file_locked`、`invalid_selector`。属性名支持自动纠错 -- 拼错属性名时会返回最接近的匹配建议。

**错误恢复** -- 智能体通过检查可用元素自行修正：

```bash
# 智能体尝试无效路径
officecli get report.docx /body/p[99] --json
# 返回: {"success": false, "error": {"error": "...", "code": "not_found", "suggestion": "..."}}

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
| 内置 agent 友好渲染引擎 | ✓ | ✗ | ✗ | ✗ |
| 无头 HTML/PNG 输出 | ✓ | ✗ | 部分支持 | ✗ |
| 跨格式模板合并（`{{key}}`）| ✓ | ✗ | ✗ | ✗ |
| Dump → batch JSON 往返 | ✓ | ✗ | ✗ | ✗ |
| 实时预览（编辑后自动刷新） | ✓ | ✗ | ✗ | ✗ |
| 无头 / CI 环境 | ✓ | ✗ | 部分支持 | ✓ |
| 跨平台 | ✓ | Windows/Mac | ✓ | ✓ |
| Word + Excel + PowerPoint | ✓ | ✓ | ✓ | 需要多个库 |

## 命令参考

| 命令 | 说明 |
|------|------|
| [`create`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-create) | 创建空白 .docx、.xlsx 或 .pptx（根据扩展名判断类型） |
| [`view`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-view) | 查看内容（模式：`outline`、`text`、`annotated`、`stats`、`issues`、`html`） |
| [`get`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-get) | 获取元素及子元素（`--depth N`、`--json`） |
| [`query`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-query) | CSS 风格查询（`[attr=value]`、`:contains()`、`:has()` 等） |
| [`set`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-set) | 修改元素属性 |
| [`add`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-add) | 添加元素（或通过 `--from <path>` 克隆） |
| [`remove`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-remove) | 删除元素 |
| [`move`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-move) | 移动元素（`--to <parent>`、`--index N`、`--after <path>`、`--before <path>`） |
| [`swap`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-swap) | 交换两个元素 |
| [`validate`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-validate) | OpenXML 模式校验 |
| [`batch`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-batch) | 单次打开/保存周期内执行多条操作（stdin、`--input` 或 `--commands`；默认原子化——只要有一条失败整批回滚，`--best-effort` 保留已成功部分，`--stop-on-error` 提前中止） |
| [`merge`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-merge) | 模板合并 — 用 JSON 数据替换 `{{key}}` 占位符 |
| [`watch`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-watch) | 在浏览器中实时 HTML 预览，自动刷新 |
| [`mcp`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-mcp) | 启动 MCP 服务器，用于 AI 工具集成 |
| [`raw`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-raw) | 查看文档部件的原始 XML |
| [`raw-set`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-raw) | 通过 XPath 修改原始 XML |
| `add-part` | 添加新的文档部件（页眉、图表等） |
| [`open`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-open) | 启动驻留模式（文档保持在内存中） |
| `close` | 保存并关闭驻留模式 |
| [`install`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-install) | 安装二进制文件 + 技能文件 + MCP（`all`、`claude`、`cursor` 等） |
| `config` | 获取或设置配置 |
| `help <format> <command>` | [内置帮助](https://github.com/iOfficeAI/OfficeCLI/wiki/command-reference)（如 `officecli help pptx set shape`） |

## 端到端工作流示例

典型的智能体自愈式工作流：创建演示文稿、填充内容、验证并修复问题 -- 全程无需人工干预。

```bash
# 1. 创建
officecli create report.pptx

# 2. 添加内容
officecli add report.pptx / --type slide --prop title="Q4 Results"
officecli add report.pptx '/slide[1]' --type shape \
  --prop text="Revenue: $4.2M" --prop x=2cm --prop y=5cm --prop size=28
officecli add report.pptx / --type slide --prop title="Details"
officecli add report.pptx '/slide[2]' --type shape \
  --prop text="Growth driven by new markets" --prop x=2cm --prop y=5cm

# 3. 验证
officecli view report.pptx outline
officecli validate report.pptx

# 4. 修复发现的问题
officecli view report.pptx issues --json
# 根据输出修复问题，例如：
officecli set report.pptx '/slide[1]/shape[1]' --prop font=Arial
```

### 单位与颜色

所有尺寸和颜色属性均接受灵活的输入格式：

| 类型 | 支持的格式 | 示例 |
|------|-----------|------|
| **尺寸** | cm、in、pt、px 或原始 EMU | `2cm`、`1in`、`72pt`、`96px`、`914400` |
| **颜色** | 十六进制、命名色、RGB、主题色 | `#FF0000`、`FF0000`、`red`、`rgb(255,0,0)`、`accent1` |
| **字号** | 纯数字或带 pt 后缀 | `14`、`14pt`、`10.5pt` |
| **间距** | pt、cm、in 或倍数 | `12pt`、`0.5cm`、`1.5x`、`150%` |

## 常用模式

```bash
# 替换 Word 文档中所有 Heading1 文本
officecli query report.docx "paragraph[style=Heading1]" --json | ...
officecli set report.docx /body/p[1]/r[1] --prop text="New Title"

# 将所有幻灯片内容导出为 JSON
officecli get deck.pptx / --depth 2 --json

# 批量更新 Excel 单元格
officecli batch budget.xlsx --input updates.json --json

# 导入 CSV 数据到 Excel 工作表
officecli add budget.xlsx / --type sheet --prop name="Q1 Data"
officecli import budget.xlsx "/Q1 Data" sales.csv --header

# 模板合并批量生成报告
officecli merge invoice-template.docx invoice-001.docx --data '{"client":"Acme","total":"$5,200"}'

# 交付前检查文档质量
officecli validate report.docx && officecli view report.docx issues --json
```

**Python 调用** —— 包装一次，每次调用都返回解析好的 JSON：

```python
import json, subprocess

def cli(*args):
    return json.loads(subprocess.check_output(["officecli", *args, "--json"], text=True))

cli("create", "deck.pptx")
cli("add", "deck.pptx", "/", "--type", "slide", "--prop", "title=Q4 报告")
slide = cli("get", "deck.pptx", "/slide[1]")
print(slide["attributes"]["text"])
```

## 文档

[Wiki](https://github.com/iOfficeAI/OfficeCLI/wiki) 提供了每个命令、元素类型和属性的详细指南：

- **按格式查看：**[Word](https://github.com/iOfficeAI/OfficeCLI/wiki/word-reference) | [Excel](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-reference) | [PowerPoint](https://github.com/iOfficeAI/OfficeCLI/wiki/powerpoint-reference)
- **工作流：**[端到端示例](https://github.com/iOfficeAI/OfficeCLI/wiki/workflows) -- Word 报告、Excel 数据表、PPT 演示、批量修改、驻留模式
- **故障排除：**[常见错误与解决方案](https://github.com/iOfficeAI/OfficeCLI/wiki/troubleshooting)
- **AI 智能体指南：**[Wiki 导航决策树](https://github.com/iOfficeAI/OfficeCLI/wiki/agent-guide)

## 从源码构建

编译需要 [.NET 10 SDK](https://dotnet.microsoft.com/download)。输出为自包含的原生二进制文件 -- .NET 已内嵌，运行时无需安装。

```bash
./build.sh
```

## 许可证

[Apache License 2.0](LICENSE)

欢迎通过 [GitHub Issues](https://github.com/iOfficeAI/OfficeCLI/issues) 提交 Bug 报告和贡献代码。

---

如果觉得 OfficeCLI 好用，请在 [GitHub 上点个 Star](https://github.com/iOfficeAI/OfficeCLI) — 帮助更多人发现这个项目。

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
alternatives: python-docx, openpyxl, python-pptx, libreoffice --headless
install-command-unix: curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
install-command-windows: irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
-->
