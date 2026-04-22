# kai-export-ppt-lite

Pure-Python HTML-to-PPTX export for editable slide decks in sandboxed environments.  
面向沙箱环境的纯 Python HTML 转 PPTX 导出器，目标是生成可编辑的 PowerPoint，而不是截图式 PPT。

Current release: `v1.2.0`

## 中文说明

### 这是什么

`kai-export-ppt-lite` 用来把 HTML 演示稿导出为可编辑的 `.pptx`。它不依赖 Playwright、Chrome 或 Node.js，适合浏览器运行时不可用、但仍需要可编辑 PPT 输出的环境。

它当前的核心定位是：

- 用 `BeautifulSoup` 和 `lxml` 解析 HTML/CSS
- 用 `python-pptx` 生成原生、可编辑的 PowerPoint 元素
- 在没有浏览器布局引擎的前提下，模拟文本、卡片、列表、表格、presentation rows、inline code/kbd/badge 等结构
- 配合回归测试和评估脚本，持续把 exporter 从单 deck 修补推进到可复用能力

### 当前状态

当前版本的验证基线是本地 `Blue Sky` 参考 deck：

- 源 HTML：`demo/blue-sky-zh.html`
- 本地 golden 参考：`demo/blue-sky-golden-native.pptx`
- 最新验证分数：`9.1/10`
- 结构化评估：
  - `overflow = 0`
  - `overlap = 0`
  - `card containment = 0`
  - `element gaps = 2`
  - `total actionable = 2`

还没有达到“每页都 >= 9.5”的最终目标，但这条分支已经明显从“按页打补丁”转成了“增强 exporter 通用能力”。

### 主要能力

- 可编辑文本导出：`h1-h6`、`p`、`li`、`span`、`a`
- 卡片、圆角矩形、边框、accent bar、pill-like inline overlay
- 支持 `code`、`kbd`、`badge`、`link`、grouped inline command row
- 支持区分真实数据表格与 presentation-like rows
- 无浏览器条件下的 flex/grid 近似布局
- 面向回归的测试体系与视觉/结构 eval

### 仓库内容

GitHub 仓库当前主要包含：

- `scripts/export-sandbox-pptx.py`
  主导出器
- `scripts/test-export.py`
  回归测试
- `scripts/rigorous-eval.py`
  多维评估脚本
- `scripts/compare-visual-comprehensive.py`
  和 golden deck 的视觉结构对比
- `tests/fixtures/export-corpus/`
  手写 fixture，用于保证通用导出能力
- `SKILL.md`
  Codex skill 入口说明

以下目录/文件是本地工作资产，不作为 GitHub 仓库交付内容：

- `demo/`
  本地输入 HTML、golden deck、输出 PPTX/预览图
- `docs/`
  本地链接到 `mydocs` 的设计与复盘文档
- `memory/`
  本地工作断点和 session 记录

### 安装

```bash
pip install beautifulsoup4 lxml python-pptx Pillow
```

### 用法

导出 HTML 到 PPTX：

```bash
python3 scripts/export-sandbox-pptx.py demo/blue-sky-zh.html demo/output.pptx
```

可选参数：

```bash
python3 scripts/export-sandbox-pptx.py <file.html> [demo/output.pptx] [--width 1440] [--height 810] [--no-chrome]
```

- `--width`
  幻灯片宽度，默认 `1440`
- `--height`
  幻灯片高度，默认 `810`
- `--no-chrome`
  跳过页码和导航点

### 验证流程

```bash
python3 -m py_compile scripts/export-sandbox-pptx.py scripts/test-export.py scripts/rigorous-eval.py
python3 scripts/test-export.py
python3 scripts/export-sandbox-pptx.py demo/blue-sky-zh.html demo/output.pptx
python3 scripts/rigorous-eval.py
```

### v1.2.0 更新重点

- README 升级为适合 GitHub 首页的中英双语说明
- 仓库发布面收敛：本地 `docs/`、根目录历史输出产物不再作为仓库内容
- `presentation_rows` 在 generic table renderer 中支持更合理的列宽分配
- shortcut-heavy 首列现在会获得额外 runway，减少快捷键导航行被挤压
- centered closing command row 的尾部链接色已与参考 deck 对齐
- divider 的 auto-margin 定位回到稳定居中逻辑

### 已知边界

- 还没有实现和 native golden 完全一致
- 当前低分主要集中在：
  - Slide 10 closing card 的几何宽度与 paragraph model
  - 若干页面的小标题/图标深墨色与 golden 黑色之间的细微差异
  - 部分 card 的高度和内部节奏仍有轻微漂移

## English

### What This Is

`kai-export-ppt-lite` converts HTML presentations into editable `.pptx` files without Playwright, Chrome, or Node.js. It is built for environments where a browser-based renderer cannot be installed, but editable PowerPoint output is still required.

The current exporter is centered on:

- HTML/CSS parsing via `BeautifulSoup` and `lxml`
- Native, editable PPTX generation via `python-pptx`
- Browser-free layout approximation for text, cards, lists, tables, presentation rows, and inline code/kbd/badge patterns
- Regression tests and evaluation tooling that push the exporter toward reusable capabilities instead of deck-specific patches

### Current Status

The current validated baseline is the local `Blue Sky` reference deck:

- Source HTML: `demo/blue-sky-zh.html`
- Local golden reference: `demo/blue-sky-golden-native.pptx`
- Latest verified score: `9.1/10`
- Structured eval:
  - `overflow = 0`
  - `overlap = 0`
  - `card containment = 0`
  - `element gaps = 2`
  - `total actionable = 2`

This branch has not yet reached the target of “every slide >= 9.5”, but it is now substantially more generalized than the earlier slide-by-slide patching phase.

### Core Capabilities

- Editable text export for `h1-h6`, `p`, `li`, `span`, and `a`
- Rounded cards, fills, borders, accent bars, and pill-like inline overlays
- First-class handling for `code`, `kbd`, `badge`, `link`, and grouped command rows
- Separation between real data tables and presentation-like rows
- Browser-free flex/grid approximation
- Regression tests and multi-dimensional eval tooling

### Repository Contents

The GitHub repository mainly ships:

- `scripts/export-sandbox-pptx.py`
  Main exporter
- `scripts/test-export.py`
  Regression suite
- `scripts/rigorous-eval.py`
  Multi-dimensional evaluation gate
- `scripts/compare-visual-comprehensive.py`
  Golden-vs-sandbox comparison tool
- `tests/fixtures/export-corpus/`
  Handwritten fixtures for generalized exporter coverage
- `SKILL.md`
  Codex skill entry for this repository

The following assets remain local-only and are not intended as GitHub repository content:

- `demo/`
  Local input decks, golden references, generated PPTX files, and previews
- `docs/`
  Local symlink to `mydocs` for design notes and optimization history
- `memory/`
  Local checkpoints and session logs

### Installation

```bash
pip install beautifulsoup4 lxml python-pptx Pillow
```

### Usage

Export HTML to PPTX:

```bash
python3 scripts/export-sandbox-pptx.py demo/blue-sky-zh.html demo/output.pptx
```

Optional flags:

```bash
python3 scripts/export-sandbox-pptx.py <file.html> [demo/output.pptx] [--width 1440] [--height 810] [--no-chrome]
```

- `--width`
  Slide width in pixels. Default: `1440`
- `--height`
  Slide height in pixels. Default: `810`
- `--no-chrome`
  Skip page counter and nav dots

### Validation Workflow

```bash
python3 -m py_compile scripts/export-sandbox-pptx.py scripts/test-export.py scripts/rigorous-eval.py
python3 scripts/test-export.py
python3 scripts/export-sandbox-pptx.py demo/blue-sky-zh.html demo/output.pptx
python3 scripts/rigorous-eval.py
```

### v1.2.0 Highlights

- README rewritten as a GitHub-facing bilingual document
- Release surface cleaned up so local `docs/` and historical root output artifacts are no longer treated as repository content
- `presentation_rows` now get stronger column-width handling even on the generic table renderer path
- Shortcut-heavy first columns receive extra runway, reducing cramped navigation rows
- Centered closing command rows now mute trailing link color to better match the reference deck
- Auto-margin divider positioning has been normalized back to stable centering behavior

### Known Gaps

- The exporter still does not fully match the native golden deck
- Remaining lower-score areas are concentrated in:
  - Slide 10 closing card geometry and paragraph model
  - minor heading/icon ink differences on a few slides
  - small card-height and internal-rhythm drift in selected layouts

See [RELEASE.md](./RELEASE.md) for release notes.
