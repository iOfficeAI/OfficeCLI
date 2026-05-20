# OfficeCLI

> **OfficeCLI は世界初にして最高の、AI エージェント向けに設計された Office スイートです。**

**あらゆる AI エージェントに Word、Excel、PowerPoint の完全な制御権を — たった一行のコードで。**

オープンソース。単一バイナリ。Office のインストール不要。依存関係ゼロ。全プラットフォーム対応。

**エージェントフレンドリーなレンダリングエンジンを内蔵** — エージェントは自分が作ったものを "見る" ことができ、Office 不要。`.docx` / `.xlsx` / `.pptx` を HTML または PNG にレンダリングし、"レンダリング → 見る → 修正" のループはバイナリが動くあらゆる場所で完結します。

[![GitHub Release](https://img.shields.io/github/v/release/iOfficeAI/OfficeCLI)](https://github.com/iOfficeAI/OfficeCLI/releases)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](LICENSE)

[English](README.md) | [中文](README_zh.md) | **日本語** | [한국어](README_ko.md)

<p align="center">
  <strong>💬 コミュニティ:</strong> <a href="https://discord.gg/2QAwJn7Egx" target="_blank">Discord</a>
</p>

<p align="center">
  <img src="assets/ppt-process.webp" alt="AionUi で OfficeCLI を使った PPT 作成プロセス" width="100%">
</p>

<p align="center"><em><a href="https://github.com/iOfficeAI/AionUi">AionUi</a> で OfficeCLI を使った PPT 作成プロセス</em></p>

<p align="center"><strong>PowerPoint プレゼンテーション</strong></p>

<table>
<tr>
<td width="33%"><img src="assets/designwhatmovesyou.gif" alt="OfficeCLI デザインプレゼン (PowerPoint)"></td>
<td width="33%"><img src="assets/horizon.gif" alt="OfficeCLI ビジネスプレゼン (PowerPoint)"></td>
<td width="33%"><img src="assets/efforless.gif" alt="OfficeCLI テクノロジープレゼン (PowerPoint)"></td>
</tr>
<tr>
<td width="33%"><img src="assets/blackhole.gif" alt="OfficeCLI 宇宙プレゼン (PowerPoint)"></td>
<td width="33%"><img src="assets/first-ppt-aionui.gif" alt="OfficeCLI ゲームプレゼン (PowerPoint)"></td>
<td width="33%"><img src="assets/shiba.gif" alt="OfficeCLI クリエイティブプレゼン (PowerPoint)"></td>
</tr>
</table>

<p align="center">—</p>
<p align="center"><strong>Word 文書</strong></p>

<table>
<tr>
<td width="33%"><img src="assets/showcase/git/word-1-officecli.gif" alt="OfficeCLI 学術論文 (Word)"></td>
<td width="33%"><img src="assets/showcase/git/word-2-officecli.gif" alt="OfficeCLI プロジェクト提案書 (Word)"></td>
<td width="33%"><img src="assets/showcase/git/word-3-officecli.gif" alt="OfficeCLI 年次報告書 (Word)"></td>
</tr>
</table>

<p align="center">—</p>
<p align="center"><strong>Excel スプレッドシート</strong></p>

<table>
<tr>
<td width="33%"><img src="assets/showcase/git/excel-1.gif" alt="OfficeCLI 予算管理 (Excel)"></td>
<td width="33%"><img src="assets/showcase/git/excel-2.gif" alt="OfficeCLI 成績管理 (Excel)"></td>
<td width="33%"><img src="assets/showcase/git/excel-3.gif" alt="OfficeCLI 売上ダッシュボード (Excel)"></td>
</tr>
</table>

<p align="center"><em>上記の文書はすべて AI エージェントが OfficeCLI を使って全自動で作成 — テンプレートなし、手動編集なし。</em></p>

## AI エージェント向け — 一行で開始

これを AI エージェントのチャットに貼り付けるだけ — スキルファイルを自動で読み込み、インストールを完了します：

```
curl -fsSL https://officecli.ai/SKILL.md
```

これだけです。スキルファイルがエージェントにバイナリのインストール方法と全コマンドの使い方を教えます。

## 一般ユーザー向け

**オプション A — GUI：** [**AionUi**](https://github.com/iOfficeAI/AionUi) をインストール — 自然言語で Office 文書を作成・編集できるデスクトップアプリ。内部で OfficeCLI が動いています。やりたいことを説明するだけで、AionUi がすべて処理します。

**オプション B — CLI：** [GitHub Releases](https://github.com/iOfficeAI/OfficeCLI/releases) からお使いのプラットフォーム用バイナリをダウンロードして、以下を実行：

```bash
officecli install
```

バイナリを PATH にコピーし、検出されたすべての AI コーディングエージェント（Claude Code、Cursor、Windsurf、GitHub Copilot など）に **officecli スキル**を自動インストールします。エージェントはすぐに Office 文書の作成・読み取り・編集が可能になります。追加設定は不要です。

## 開発者向け — 30秒でライブ体験

```bash
# 1. インストール（macOS / Linux）
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
# Windows (PowerShell): irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex

# 2. 空の PowerPoint を作成
officecli create deck.pptx

# 3. ライブプレビューを開始 — ブラウザで http://localhost:26315 が開きます
officecli watch deck.pptx

# 4. 別のターミナルを開いてスライドを追加 — ブラウザが即座に更新されます
officecli add deck.pptx / --type slide --prop title="Hello, World!"
```

これだけです。`add`、`set`、`remove` コマンドを実行するたびに、プレビューがリアルタイムで更新されます。どんどん試してみてください — ブラウザがあなたのライブフィードバックループです。

## クイックスタート

```bash
# プレゼンテーションを作成してコンテンツを追加
officecli create deck.pptx
officecli add deck.pptx / --type slide --prop title="Q4 Report" --prop background=1A1A2E
officecli add deck.pptx '/slide[1]' --type shape \
  --prop text="Revenue grew 25%" --prop x=2cm --prop y=5cm \
  --prop font=Arial --prop size=24 --prop color=FFFFFF

# アウトラインを表示
officecli view deck.pptx outline
# → Slide 1: Q4 Report
# →   Shape 1 [TextBox]: Revenue grew 25%

# HTML で表示 — サーバー不要、ブラウザでレンダリングされたプレビューを開きます
officecli view deck.pptx html

# 任意の要素の構造化 JSON を取得
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

## なぜ OfficeCLI？

以前は 50行の Python と 3つのライブラリが必要でした：

```python
from pptx import Presentation
from pptx.util import Inches, Pt
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])
title = slide.shapes.title
title.text = "Q4 Report"
# ... さらに 45行 ...
prs.save('deck.pptx')
```

今はコマンド一つで：

```bash
officecli add deck.pptx / --type slide --prop title="Q4 Report"
```

**OfficeCLI でできること：**

- **作成** ドキュメント -- 空白またはコンテンツ付き
- **読み取り** テキスト、構造、スタイル、数式 -- プレーンテキストまたは構造化 JSON
- **分析** フォーマットの問題、スタイルの不整合、構造的な欠陥
- **修正** 任意の要素 -- テキスト、フォント、色、レイアウト、数式、チャート、画像
- **再構成** コンテンツ -- 要素の追加、削除、移動、文書間コピー

| フォーマット | 読み取り | 修正 | 作成 |
|-------------|---------|------|------|
| Word (.docx) | ✅ | ✅ | ✅ |
| Excel (.xlsx) | ✅ | ✅ | ✅ |
| PowerPoint (.pptx) | ✅ | ✅ | ✅ |

**Word** — 完全な [i18n & RTL サポート](https://github.com/iOfficeAI/OfficeCLI/wiki/i18n)（スクリプト別フォントスロット、スクリプト別 BCP-47 言語タグ `lang.latin/ea/cs`、複雑スクリプトの太字/斜体/サイズ、段落/ラン/セクション/表/スタイル/ヘッダー/フッター/docDefaults をカスケードする `direction=rtl`、`rtlGutter` + `pgBorders` ショートハンド、ヒンディー語/アラビア語/タイ語/CJK のロケール対応ページ番号）、[段落](https://github.com/iOfficeAI/OfficeCLI/wiki/word-paragraph)、[ラン](https://github.com/iOfficeAI/OfficeCLI/wiki/word-run)、[表](https://github.com/iOfficeAI/OfficeCLI/wiki/word-table)、[スタイル](https://github.com/iOfficeAI/OfficeCLI/wiki/word-style)、[ヘッダー/フッター](https://github.com/iOfficeAI/OfficeCLI/wiki/word-header-footer)、[画像](https://github.com/iOfficeAI/OfficeCLI/wiki/word-picture)（PNG/JPG/GIF/SVG）、[数式](https://github.com/iOfficeAI/OfficeCLI/wiki/word-equation)、[コメント](https://github.com/iOfficeAI/OfficeCLI/wiki/word-comment)、[脚注](https://github.com/iOfficeAI/OfficeCLI/wiki/word-footnote)、[透かし](https://github.com/iOfficeAI/OfficeCLI/wiki/word-watermark)、[ブックマーク](https://github.com/iOfficeAI/OfficeCLI/wiki/word-bookmark)、[目次](https://github.com/iOfficeAI/OfficeCLI/wiki/word-toc)、[チャート](https://github.com/iOfficeAI/OfficeCLI/wiki/word-chart)、[ハイパーリンク](https://github.com/iOfficeAI/OfficeCLI/wiki/word-hyperlink)、[セクション](https://github.com/iOfficeAI/OfficeCLI/wiki/word-section)、[フォームフィールド](https://github.com/iOfficeAI/OfficeCLI/wiki/word-formfield)、[コンテンツコントロール (SDT)](https://github.com/iOfficeAI/OfficeCLI/wiki/word-sdt)、[フィールド](https://github.com/iOfficeAI/OfficeCLI/wiki/word-field)（22 種類のゼロ引数 + MERGEFIELD / REF / PAGEREF / SEQ / STYLEREF / DOCPROPERTY / IF）、[OLE オブジェクト](https://github.com/iOfficeAI/OfficeCLI/wiki/word-ole)、[文書プロパティ](https://github.com/iOfficeAI/OfficeCLI/wiki/word-document)

**Excel** — [セル](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-cell)（追加時にふりがな対応）、数式（150以上の組み込み関数を自動計算、動的配列関数に `_xlfn.` 自動プレフィックス）、[シート](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-sheet)（visible/hidden/veryHidden、印刷余白、printTitleRows/Cols、RTL `sheetView`、カスケード対応のシート名変更）、[テーブル](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-table)、[ソート](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-sort)（シート/範囲、マルチキー、サイドカー対応）、[条件付き書式](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-conditionalformatting)、[チャート](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-chart)（箱ひげ図、[パレート図](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-chart-add) 自動ソート + 累積%、対数軸を含む）、[ピボットテーブル](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-pivottable)（マルチフィールド、日付グループ化、showDataAs、ソート、総計、小計、コンパクト/アウトライン/表形式レイアウト、項目ラベル繰り返し、空白行、計算フィールド）、[スライサー](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-slicer)、[名前付き範囲](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-namedrange)、[データ入力規則](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-validation)、[画像](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-picture)（PNG/JPG/GIF/SVG、デュアル表現フォールバック）、[スパークライン](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-sparkline)、[コメント](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-comment)（RTL）、[オートフィルター](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-autofilter)、[図形](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-shape)、[OLE オブジェクト](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-ole)、CSV/TSV インポート、`$Sheet:A1` セルアドレッシング

**PowerPoint** — [スライド](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-slide)（ヘッダー/フッター/日付/スライド番号トグル、非表示）、[図形](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-shape)（パターン塗りつぶし、ぼかし効果、ハイパーリンクツールチップ + スライドジャンプリンク）、[画像](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-picture)（PNG/JPG/GIF/SVG、塗りモード: stretch/contain/cover/tile、明るさ/コントラスト/光彩/影）、[表](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-table)、[チャート](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-chart)、[アニメーション](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-slide)、[モーフトランジション](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-morph-check)、[3D モデル (.glb)](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-3dmodel)、[スライドズーム](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-zoom)、[数式](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-equation)、[テーマ](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-theme)、[コネクタ](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-connector)、[ビデオ/オーディオ](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-video)、[グループ](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-group)、[ノート](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-notes)（RTL、lang）、[コメント](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-comment)（RTL）、[OLE オブジェクト](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-ole)、[プレースホルダー](https://github.com/iOfficeAI/OfficeCLI/wiki/ppt-placeholder)（phType で追加/設定）

## 使用シーン

**開発者向け：**
- データベースや API からのレポート自動生成
- 文書の一括処理（一括検索/置換、スタイル更新）
- CI/CD 環境でのドキュメントパイプライン構築（テスト結果からドキュメント生成）
- Docker/コンテナ環境でのヘッドレス Office 自動化

**AI エージェント向け：**
- ユーザーのプロンプトからプレゼンテーションを生成（上記の例を参照）
- ドキュメントから構造化データを JSON に抽出
- 納品前のドキュメント品質検証

**チーム向け：**
- ドキュメントテンプレートを複製してデータを入力
- CI/CD パイプラインでの自動ドキュメント検証

## インストール

単一の自己完結型バイナリとして配布。.NET ランタイムは内蔵 -- インストール不要、ランタイム管理不要。

**ワンライナーインストール：**

```bash
# macOS / Linux
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash

# Windows (PowerShell)
irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
```

**または手動ダウンロード** [GitHub Releases](https://github.com/iOfficeAI/OfficeCLI/releases)：

| プラットフォーム | バイナリ |
|----------------|---------|
| macOS Apple Silicon | `officecli-mac-arm64` |
| macOS Intel | `officecli-mac-x64` |
| Linux x64 | `officecli-linux-x64` |
| Linux ARM64 | `officecli-linux-arm64` |
| Windows x64 | `officecli-win-x64.exe` |
| Windows ARM64 | `officecli-win-arm64.exe` |

インストール確認：`officecli --version`

**またはダウンロード済みバイナリからセルフインストール（`officecli` を直接実行してもインストールがトリガーされます）：**

```bash
officecli install    # 明示的インストール
officecli            # 直接実行でもインストールがトリガー
```

更新はバックグラウンドで自動チェックされます。`officecli config autoUpdate false` で無効化、または `OFFICECLI_SKIP_UPDATE=1` で単回スキップ可能。設定は `~/.officecli/config.json` にあります。

## 主な機能

### 内蔵エンジンと生成プリミティブ

OfficeCLI は自己完結型です。以下の機能はすべてバイナリ内蔵 — **Office 不要**。

#### レンダリングエンジン

ゼロから実装したエージェントフレンドリーなレンダリングエンジンがバイナリ内に同梱され、シェイプ、チャート (トレンドライン、エラーバー、ウォーターフォール、ローソク足、スパークライン)、数式 (OMML → MathJax 互換)、Three.js による 3D `.glb` モデル、モーフトランジション、スライドズーム、シェイプエフェクトをカバー。ページごとの PNG スクリーンショットは、レンダリングされた HTML をヘッドレスブラウザに渡して生成されます。3 つのモード:

- **`view html`** — スタンドアロン HTML ファイル、アセットインライン。任意のブラウザで開けます。
- **`view screenshot`** — ページごとの PNG、マルチモーダルエージェント向け。
- **`watch`** — ローカル HTTP サーバー + 自動更新プレビュー。`add` / `set` / `remove` でブラウザが即座に更新。Excel watch はインラインセル編集とチャートのドラッグ再配置をサポート。

```bash
officecli view deck.pptx html -o /tmp/deck.html
officecli view deck.pptx screenshot -o /tmp/deck.png # 複数ページは --page 1-N
officecli watch deck.pptx                            # http://localhost:26315
```

> 可視化なしでは、スライドを生成するエージェントは盲目的に飛んでいるようなもの — DOM は読めても、タイトルがオーバーフローしているか、2 つのシェイプが重なっているかは判断できません。レンダリングがバイナリに内蔵されているため、"レンダリング → 見る → 修正" のループは CI、Docker、ディスプレイのないサーバー — バイナリが動くあらゆる場所で動作します。

#### 数式 & ピボットエンジン

150+ の Excel 関数が書き込み時に自動評価 — `=SUM(A1:A2)` を書いて、セルを `get` する、値はすでにそこに。Office で再計算するラウンドトリップは不要。動的配列関数 (`FILTER` / `UNIQUE` / `SORT` / `SEQUENCE`、`_xlfn.` 自動プレフィックス)、`VLOOKUP` / `INDEX` / `MATCH`、日付・テキスト関数など 140+ の関数をカバー。

加えて、ソース範囲から 1 コマンドでネイティブな OOXML ピボットテーブル — マルチフィールドの行/列/フィルター、10 種類の集計、`showDataAs` モード、日付グループ化、計算フィールド、Top-N、レイアウト。ピボットキャッシュ + 定義は OOXML に書き込まれ、Excel で開くと集計済みの状態で表示されます:

```bash
officecli add sales.xlsx '/Sheet1' --type pivottable \
  --prop source='Data!A1:E10000' --prop rows='Region,Category' \
  --prop cols=Quarter --prop values='Revenue:sum,Units:avg' \
  --prop showDataAs=percentOfTotal
```

#### テンプレートマージ — 一度設計、N 回入力

`merge` は任意の `.docx` / `.xlsx` / `.pptx` の `{{key}}` プレースホルダーを JSON データで置換 — 段落、表セル、シェイプ、ヘッダー/フッター、チャートタイトル全体で動作。エージェントが一度レイアウトを設計 (高コスト)、本番コードが N 回入力 (低コスト、決定論的、トークンコストゼロ)。エージェントが各レポートを毎回ゼロから再生成し、N 個の一貫性のないレイアウトを生み出す失敗モードを回避します。

```bash
officecli merge invoice-template.docx out-001.docx '{"client":"Acme","total":"$5,200"}'
officecli merge q4-template.pptx q4-acme.pptx data.json
```

#### Dump によるラウンドトリップ — 既存ドキュメントから学ぶ

`dump` は任意の `.docx` — ドキュメント全体**または任意のサブツリー**（単一の段落、表、styles、numbering、theme、settings）— を再生可能なバッチ JSON にシリアライズし、`batch` で再生。ユーザーが模倣したいサンプルから、エージェントは生の OOXML XML ではなく構造化された仕様を読み、変更して再生します。"既存テンプレートがある" と "100 個のバリエーションを生成して" を繋ぎます。

```bash
officecli dump existing.docx -o blueprint.json                  # ドキュメント全体
officecli dump existing.docx /body/tbl[1] -o table.json         # 任意のサブツリー
officecli batch new.docx --input blueprint.json
```

### レジデントモードとバッチ

複数ステップのワークフローでは、レジデントモードがドキュメントをメモリに保持。バッチモードは一度の open/save サイクルで複数操作を実行します。

```bash
# レジデントモード — 名前付きパイプ経由で遅延ほぼゼロ
officecli open report.docx
officecli set report.docx /body/p[1]/r[1] --prop bold=true
officecli set report.docx /body/p[2]/r[1] --prop color=FF0000
officecli close report.docx

# バッチモード — アトミックなマルチコマンド実行（デフォルトで最初のエラーで停止）
echo '[{"command":"set","path":"/slide[1]/shape[1]","props":{"text":"Hello"}},
      {"command":"set","path":"/slide[1]/shape[2]","props":{"fill":"FF0000"}}]' \
  | officecli batch deck.pptx --json

# インラインバッチ — stdin 不要
officecli batch deck.pptx --commands '[{"op":"set","path":"/slide[1]/shape[1]","props":{"text":"Hi"}}]'

# --force でエラーをスキップして続行
officecli batch deck.pptx --input updates.json --force --json
```

### 三層アーキテクチャ

シンプルに始めて、必要な時だけ深く。

| レイヤー | 用途 | コマンド |
|---------|------|---------|
| **L1：読み取り** | コンテンツのセマンティックビュー | `view`（text、annotated、outline、stats、issues、html、svg、screenshot） |
| **L2：DOM** | 構造化された要素操作 | `get`、`query`、`set`、`add`、`remove`、`move`、`swap` |
| **L3：生 XML** | XPath による直接アクセス — 万能フォールバック | `raw`、`raw-set`、`add-part`、`validate` |

```bash
# L1 — 高レベルビュー
officecli view report.docx annotated
officecli view budget.xlsx text --cols A,B,C --max-lines 50

# L2 — 要素レベルの操作
officecli query report.docx "run:contains(TODO)"
officecli add budget.xlsx / --type sheet --prop name="Q2 Report"
officecli move report.docx /body/p[5] --to /body --index 1

# L3 — L2 では足りない時に生 XML
officecli raw deck.pptx '/slide[1]'
officecli raw-set report.docx document \
  --xpath "//w:p[1]" --action append \
  --xml '<w:r><w:t>Injected text</w:t></w:r>'
```

## AI 統合

### MCP サーバー

組み込み [MCP](https://modelcontextprotocol.io) サーバー — コマンド一つで登録：

```bash
officecli mcp claude       # Claude Code
officecli mcp cursor       # Cursor
officecli mcp vscode       # VS Code / Copilot
officecli mcp lmstudio     # LM Studio
officecli mcp list         # 登録状態を確認
```

JSON-RPC で全ドキュメント操作を公開 — シェルアクセス不要。

### 直接 CLI 統合

2ステップで OfficeCLI を任意の AI エージェントに統合：

1. **バイナリをインストール** -- コマンド一つ（[インストール](#インストール)参照）
2. **完了。** OfficeCLI は AI ツール（Claude Code、GitHub Copilot、Codex）を自動検出し、既知の設定ディレクトリを確認してスキルファイルをインストールします。エージェントはすぐに Office 文書の作成・読み取り・変更が可能です。

<details>
<summary><strong>手動設定（オプション）</strong></summary>

自動インストールがお使いの環境に対応していない場合、手動でスキルファイルをインストールできます：

**SKILL.md を直接エージェントに読み込ませる：**

```bash
curl -fsSL https://officecli.ai/SKILL.md
```

**Claude Code のローカルスキルとしてインストール：**

```bash
curl -fsSL https://officecli.ai/SKILL.md -o ~/.claude/skills/officecli.md
```

**その他のエージェント：** `SKILL.md` の内容をエージェントのシステムプロンプトまたはツール説明に含めてください。

</details>

### エージェントが OfficeCLI で活躍する理由

- **決定論的 JSON 出力** — すべてのコマンドが `--json` をサポートし、スキーマは一貫。正規表現パース不要、stdout スクレイピング不要。
- **パスベースのアドレッシング** — すべての要素に安定したパス (`/slide[1]/shape[2]`)。エージェントは XML 名前空間を理解せずにドキュメントをナビゲート可能。(OfficeCLI 独自の構文: 1-based インデックス、要素ローカル名 — XPath ではない。)
- **段階的複雑度 (L1 → L2 → L3)** — エージェントは読み取り専用ビューから始め、DOM 操作にエスカレート、必要な時のみ raw XML にフォールバック。トークン消費を最小化。
- **自己修復ワークフロー** — `validate`、`view issues`、構造化エラーコード (`not_found`、`invalid_value`、`unsupported_property`) は suggestion と有効範囲を返します。エージェントは人間の介入なしに自己修正します。
- **内蔵エージェントフレンドリーレンダリングエンジン** — `view html` / `view screenshot` / `watch` がネイティブに HTML と PNG を出力。Office 不要。エージェントは CI / Docker / ヘッドレス環境でも自分の出力を "見て" レイアウトの問題を修正できます。
- **内蔵数式 & ピボットエンジン** — 150+ の Excel 関数が書き込み時に自動評価; ソース範囲から 1 コマンドでネイティブ OOXML ピボットテーブル。エージェントは Office で再計算せずに、計算値と集計結果を即座に読み取れます。
- **テンプレートマージ** — エージェントがレイアウトを一度設計し、下流コードが `{{key}}` プレースホルダーを N 回入力。各レポートを再生成してトークンを焼くことを避けます。
- **ラウンドトリップ Dump** — `dump` が任意の `.docx` を再生可能なバッチ JSON に変換。エージェントは生の OOXML XML ではなく構造化された仕様を読んで、人間が作成したサンプルから学習。
- **内蔵ヘルプ** — プロパティ名や値形式に迷ったら、エージェントは推測せず `officecli <format> set <element>` を実行。
- **自動インストール** — OfficeCLI は使っているツール (Claude Code、Cursor、VS Code…) を検出して自己構成します。手動の skill ファイルセットアップ不要。

### 組み込みヘルプ

プロパティ名がわからない時は、階層型ヘルプで確認：

```bash
officecli pptx set              # 全設定可能な要素とプロパティ
officecli pptx set shape        # 特定の要素タイプの詳細
officecli pptx set shape.fill   # 単一プロパティのフォーマットと例
officecli docx query            # セレクタリファレンス：属性、:contains、:has() など
```

`pptx` を `docx` や `xlsx` に置き換え可能。動詞は `view`、`get`、`query`、`set`、`add`、`raw`。

`officecli --help` で全体概要を確認。

### JSON 出力スキーマ

全コマンドが `--json` に対応。一般的なレスポンス形式：

**単一要素**（`get --json`）：

```json
{"tag": "shape", "path": "/slide[1]/shape[1]", "attributes": {"name": "TextBox 1", "text": "Hello"}}
```

**要素リスト**（`query --json`）：

```json
[
  {"tag": "paragraph", "path": "/body/p[1]", "attributes": {"style": "Heading1", "text": "Title"}},
  {"tag": "paragraph", "path": "/body/p[5]", "attributes": {"style": "Heading1", "text": "Summary"}}
]
```

**エラー** は構造化エラーオブジェクトを返却。エラーコード、修正提案、利用可能な値を含みます：

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

エラーコード：`not_found`、`invalid_value`、`unsupported_property`、`invalid_path`、`unsupported_type`、`missing_property`、`file_not_found`、`file_locked`、`invalid_selector`。プロパティ名は自動修正対応 -- プロパティ名のスペルミスは最も近い候補を提案します。

**エラー回復** -- エージェントは利用可能な要素を確認して自己修正：

```bash
# エージェントが無効なパスを試行
officecli get report.docx /body/p[99] --json
# 返却: {"success": false, "error": {"error": "...", "code": "not_found", "suggestion": "..."}}

# エージェントが利用可能な要素を確認して自己修正
officecli get report.docx /body --depth 1 --json
# 利用可能な子要素のリストを返却、エージェントが正しいパスを選択
```

**変更確認**（`set`、`add`、`remove`、`move`、`create` で `--json` 使用時）：

```json
{"success": true, "path": "/slide[1]/shape[1]"}
```

`officecli --help` で終了コードとエラー形式の完全な説明を確認。

## 比較

| | OfficeCLI | Microsoft Office | LibreOffice | python-docx / openpyxl |
|---|---|---|---|---|
| オープンソース＆無料 | ✓ (Apache 2.0) | ✗（有料ライセンス） | ✓ | ✓ |
| AI ネイティブ CLI + JSON | ✓ | ✗ | ✗ | ✗ |
| ゼロインストール（単一バイナリ） | ✓ | ✗ | ✗ | ✗（Python + pip 必要） |
| 任意の言語から呼び出し | ✓ (CLI) | ✗ (COM/Add-in) | ✗ (UNO API) | Python のみ |
| パスベースの要素アクセス | ✓ | ✗ | ✗ | ✗ |
| 生 XML フォールバック | ✓ | ✗ | ✗ | 部分対応 |
| 内蔵エージェントフレンドリーレンダリングエンジン | ✓ | ✗ | ✗ | ✗ |
| ヘッドレス HTML/PNG 出力 | ✓ | ✗ | 部分対応 | ✗ |
| クロスフォーマットテンプレートマージ (`{{key}}`) | ✓ | ✗ | ✗ | ✗ |
| Dump → batch JSON ラウンドトリップ | ✓ | ✗ | ✗ | ✗ |
| ライブプレビュー (編集後自動更新) | ✓ | ✗ | ✗ | ✗ |
| ヘッドレス / CI | ✓ | ✗ | 部分対応 | ✓ |
| クロスプラットフォーム | ✓ | Windows/Mac | ✓ | ✓ |
| Word + Excel + PowerPoint | ✓ | ✓ | ✓ | 複数ライブラリが必要 |

## コマンドリファレンス

| コマンド | 説明 |
|---------|------|
| [`create`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-create) | 空白の .docx、.xlsx、.pptx を作成（拡張子からタイプを判定） |
| [`view`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-view) | コンテンツを表示（モード：`outline`、`text`、`annotated`、`stats`、`issues`、`html`） |
| [`get`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-get) | 要素と子要素を取得（`--depth N`、`--json`） |
| [`query`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-query) | CSS スタイルのクエリ（`[attr=value]`、`:contains()`、`:has()` など） |
| [`set`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-set) | 要素のプロパティを変更 |
| [`add`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-add) | 要素を追加（または `--from <path>` でクローン） |
| [`remove`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-remove) | 要素を削除 |
| [`move`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-move) | 要素を移動（`--to <parent>`、`--index N`、`--after <path>`、`--before <path>`） |
| [`swap`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-swap) | 2つの要素を交換 |
| [`validate`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-validate) | OpenXML スキーマ検証 |
| [`batch`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-batch) | 一度の open/save サイクルで複数操作を実行（stdin、`--input`、または `--commands`；デフォルトで最初のエラーで停止、`--force` で続行） |
| [`merge`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-merge) | テンプレートマージ — `{{key}}` プレースホルダーを JSON データで置換 |
| [`watch`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-watch) | ブラウザでライブ HTML プレビュー、自動更新 |
| [`mcp`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-mcp) | AI ツール統合用の MCP サーバーを起動 |
| [`raw`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-raw) | ドキュメントパートの生 XML を表示 |
| [`raw-set`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-raw) | XPath で生 XML を変更 |
| `add-part` | 新しいドキュメントパート（ヘッダー、チャートなど）を追加 |
| [`open`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-open) | レジデントモードを開始（ドキュメントをメモリに保持） |
| `close` | 保存してレジデントモードを終了 |
| [`install`](https://github.com/iOfficeAI/OfficeCLI/wiki/command-install) | バイナリ + スキル + MCP をインストール（`all`、`claude`、`cursor` など） |
| `config` | 設定の取得または変更 |
| `<format> <command>` | [組み込みヘルプ](https://github.com/iOfficeAI/OfficeCLI/wiki/command-reference)（例：`officecli pptx set shape`） |

## エンドツーエンドワークフロー例

典型的なエージェント自己修復ワークフロー：プレゼンテーションの作成、コンテンツの入力、検証、問題の修正 -- すべて人間の介入なし。

```bash
# 1. 作成
officecli create report.pptx

# 2. コンテンツを追加
officecli add report.pptx / --type slide --prop title="Q4 Results"
officecli add report.pptx '/slide[1]' --type shape \
  --prop text="Revenue: $4.2M" --prop x=2cm --prop y=5cm --prop size=28
officecli add report.pptx / --type slide --prop title="Details"
officecli add report.pptx '/slide[2]' --type shape \
  --prop text="Growth driven by new markets" --prop x=2cm --prop y=5cm

# 3. 検証
officecli view report.pptx outline
officecli validate report.pptx

# 4. 問題の修正
officecli view report.pptx issues --json
# 出力に基づいて問題を修正：
officecli set report.pptx '/slide[1]/shape[1]' --prop font=Arial
```

### 単位と色

すべての寸法・色プロパティは柔軟な入力形式に対応：

| タイプ | 対応形式 | 例 |
|-------|---------|-----|
| **寸法** | cm、in、pt、px または生 EMU | `2cm`、`1in`、`72pt`、`96px`、`914400` |
| **色** | 16進数、色名、RGB、テーマ色 | `#FF0000`、`FF0000`、`red`、`rgb(255,0,0)`、`accent1` |
| **フォントサイズ** | 数値のみまたは pt 接尾辞付き | `14`、`14pt`、`10.5pt` |
| **間隔** | pt、cm、in または倍率 | `12pt`、`0.5cm`、`1.5x`、`150%` |

## よく使うパターン

```bash
# Word 文書の全 Heading1 テキストを置換
officecli query report.docx "paragraph[style=Heading1]" --json | ...
officecli set report.docx /body/p[1]/r[1] --prop text="New Title"

# 全スライドのコンテンツを JSON でエクスポート
officecli get deck.pptx / --depth 2 --json

# Excel セルを一括更新
officecli batch budget.xlsx --input updates.json --json

# CSV データを Excel シートにインポート
officecli add budget.xlsx / --type sheet --prop name="Q1 Data" --prop csv=sales.csv

# テンプレートマージでレポートを一括生成
officecli merge invoice-template.docx invoice-001.docx '{"client":"Acme","total":"$5,200"}'

# 納品前にドキュメント品質をチェック
officecli validate report.docx && officecli view report.docx issues --json
```

**Python から呼び出し** — 一度ラップすれば、すべての呼び出しでパース済み JSON が返ります：

```python
import json, subprocess

def cli(*args):
    return json.loads(subprocess.check_output(["officecli", *args, "--json"], text=True))

cli("create", "deck.pptx")
cli("add", "deck.pptx", "/", "--type", "slide", "--prop", "title=Q4 レポート")
slide = cli("get", "deck.pptx", "/slide[1]")
print(slide["attributes"]["text"])
```

## ドキュメント

[Wiki](https://github.com/iOfficeAI/OfficeCLI/wiki) に全コマンド、要素タイプ、プロパティの詳細ガイドがあります：

- **フォーマット別：**[Word](https://github.com/iOfficeAI/OfficeCLI/wiki/word-reference) | [Excel](https://github.com/iOfficeAI/OfficeCLI/wiki/excel-reference) | [PowerPoint](https://github.com/iOfficeAI/OfficeCLI/wiki/powerpoint-reference)
- **ワークフロー：**[エンドツーエンド例](https://github.com/iOfficeAI/OfficeCLI/wiki/workflows) -- Word レポート、Excel ダッシュボード、PPT プレゼン、一括変更、レジデントモード
- **トラブルシューティング：**[よくあるエラーと解決策](https://github.com/iOfficeAI/OfficeCLI/wiki/troubleshooting)
- **AI エージェントガイド：**[Wiki ナビゲーション決定木](https://github.com/iOfficeAI/OfficeCLI/wiki/agent-guide)

## ソースからビルド

コンパイルには [.NET 10 SDK](https://dotnet.microsoft.com/download) が必要です。出力は自己完結型のネイティブバイナリ -- .NET は内蔵されているため、実行時にはインストール不要です。

```bash
./build.sh
```

## ライセンス

[Apache License 2.0](LICENSE)

バグ報告やコントリビューションは [GitHub Issues](https://github.com/iOfficeAI/OfficeCLI/issues) まで。

---

OfficeCLI が役に立ったら、ぜひ [GitHub でスターを付けてください](https://github.com/iOfficeAI/OfficeCLI) — より多くの人にプロジェクトを届ける力になります。

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
