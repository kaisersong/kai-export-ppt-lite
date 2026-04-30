# kai-export-ppt-lite

Pure-Python HTML-to-PPTX export for editable slide decks in sandboxed environments.  
面向沙箱环境的纯 Python HTML 转 PPTX 导出器，目标是生成可编辑的 PowerPoint，而不是截图式 PPT。

Current release: `v1.6.4`

## 中文说明

### 这是什么

`kai-export-ppt-lite` 用来把 HTML 演示稿导出为可编辑的 `.pptx`。它不依赖 Playwright、Chrome 或 Node.js，适合浏览器运行时不可用、但仍需要可编辑 PPT 输出的环境。

它当前的核心定位是：

- 用 `BeautifulSoup` 和 `lxml` 解析 HTML/CSS
- 用 `python-pptx` 生成原生、可编辑的 PowerPoint 元素
- 在没有浏览器布局引擎的前提下，模拟文本、卡片、列表、表格、presentation rows、inline code/kbd/badge 等结构
- 配合回归测试和评估脚本，持续把 exporter 从单 deck 修补推进到可复用能力

### 当前状态

当前版本的两个主要验证锚点是：

- `Blue Sky`
  - 源 HTML：`demo/blue-sky-zh.html`
  - 本地 golden 参考：`demo/blue-sky-golden-native.pptx`
  - 最新已验证分数：`9.1/10`
- `Chinese Chan`
  - 源 HTML：`demo/chinese-chan-zh.html`
  - 最新完整视觉对比分数：`9.68/10`
  - 当前结构化评估：
  - `overflow = 0`
  - `overlap = 0`
  - `card containment = 0`
  - `element gaps = 0`
  - `total actionable = 0`
- `Swiss Modern`
  - 约束路径：`canonical + constrained compatible`
  - 最新真实回归样本：`kingdee_from_original_v7.pptx`
  - 最新完整视觉对比分数：`9.36/10`
  - 已确认文件级文字语义：
  - `P3 body wrap = square`
  - `P5 title wrap = none`
- `Aurora Mesh`
  - 源 HTML：`demo/aurora-mesh-zh.html`
  - 当前分支最新完整视觉对比分数：`9.00/10`
  - 当前结构化评估：
  - `overflow = 0`
  - `overlap = 0`
  - 这轮优化重点：
  - 动态 mesh 背景降级为接近源稿氛围的纯色，而不是伪造大椭圆或退回纯黑
  - KPI 组件默认保持紧凑，只有源 CSS 明确给出 stretch 信号时才拉宽

整体上，这条分支已经明显从“按页打补丁”转成了“增强 exporter 通用能力 + producer-aware preset contract”。仍未达到“所有风格都稳定每页 >= 9.5”的最终目标，但 `Chinese Chan` 这类高度依赖字体与换行节奏的 preset 已经被真正拉上来。

最近新增并持续强化的，是一条 `slide-creator` contract-driven 路径，正在把 `data-story / enterprise-dark / chinese-chan / swiss-modern` 这类新风格从“启发式近似”推进到“基于 preset contract 的组件与排版导出”。其中 `Chinese Chan` 和 `Swiss Modern` 已经补上：

- typography contract
- authored line-break contract
- shared runtime chrome fallback
- centered command / seal fidelity gate
- role-aware Swiss layout builders
- title/body wrap guards against PowerPoint-only reflow

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
python3 scripts/export-sandbox-pptx.py <file.html> [demo/output.pptx] [--width 1440] [--height 900] [--with-chrome]
```

- `--width`
  幻灯片宽度，默认 `1440`
- `--height`
  幻灯片高度，默认 `900`
- `--with-chrome`
  额外添加 exporter 提供的页码和导航点

### 验证流程

```bash
python3 -m py_compile scripts/export-sandbox-pptx.py scripts/test-export.py scripts/rigorous-eval.py
python3 scripts/test-export.py
python3 scripts/export-sandbox-pptx.py demo/blue-sky-zh.html demo/output.pptx
python3 scripts/rigorous-eval.py
```

### v1.6.4 更新重点（patch）

- `build_grid_children` 新增 compact flex-row 识别（gap + 无 `space-between/around/evenly` justify-content）：子元素按 intrinsic 内容宽度紧凑排布，不再强制平分整行宽度。修复 Slide 1 / 8 的 `.hero-stats` KPI 横跨整行问题（P1 横跨 `8.7" → 1.8"`）
- `_build_swiss_column_content` 后置 stretch：左栏块级标题/段落强制取容器内宽，避免 `<br>` 显式换行被二次折成孤字（P2 标题 `2.08" → 4.03"`，由 spAutoFit 接管高度生成 2 行干净换行）
- 新增三个回归测试：`test_compact_flex_row_packs_stat_blocks_at_intrinsic_width`、`test_compact_flex_row_falls_back_to_even_split_when_oversized`、`test_stretch_column_block_text_to_inner_width_expands_narrow_heading`
- `Swiss Modern` canonical 整体保持 `9.06/10`（cross-renderer SSIM 已逼近 cap），结构对应 source CSS 的真实意图
- `Aurora Mesh` 视觉回归保持 `9.00/10`

### v1.6.3 更新重点（patch）

- `_build_swiss_column_content` 的右栏改成跟左栏一致的垂直居中（之前 canonical tier 硬编码 `right_y = 0`）
- `_build_swiss_title_grid` 的 `justify-content` 优先从 inner panel 取（如 `.flow-inner { justify-content: center }`）
- `Swiss Modern` canonical：Slide 05 `8.6 → 9.0`、Slide 06 `8.9 → 9.0`、整体 `9.00 → 9.06`
- `Aurora Mesh` 视觉回归保持 `9.00/10`

### v1.6.2 更新重点（patch）

- `display heading` 视觉补偿排除 `noto sans` 触发条件——当 Noto Sans SC 映射到 Helvetica Neue + Hiragino Sans GB 跨渲染器稳定字对时，1.30 倍 boost 会超调并把标题误折成两行
- 新增 `inner_panel` layout role 覆盖单内容面板（无 `.bg-num` 装饰）的 Swiss 幻灯片（Slide 5 `flow-inner`、Slide 6 `feat-inner`）
- `_FONT_MAP` 锁定 Archivo / Nunito / Noto Sans SC 全部映射到 Helvetica Neue + Hiragino Sans GB 跨渲染器稳定字对
- 对比脚本注入字体覆盖样式，使 SSIM 反映 layout 一致性而非跨渲染器字体差异
- `Swiss Modern` canonical 视觉快照：整体 `8.93 → 9.00`；Slide 03/04/05/07 各 `+0.1`
- `Aurora Mesh` 视觉回归保持 `9.00/10` 不变

### 已知边界

- 当前 SSIM 跨渲染器对比 gate 下，`≥ 9.5/slide` 不可达：单页 SSIM 实测上限 `~0.92`（最佳 `0.9227`），对应 score 约 `9.34`。要进一步提分需换 gate（structural eval 或共享 renderer），不是继续调 layout。

### v1.6.1 更新重点（patch）

- 修复 `_pack_direct_child_content` 不再让 absolute/fixed 子元素污染面板内容宽度
- 新增 `_build_absolute_decoration_strips`：在 Swiss `column_content` 面板内按授权 CSS 尺寸（`width / left / right / top / bottom`）渲染装饰条
- `Swiss Modern` canonical 视觉快照：Slide 02 `8.7 → 9.3`、整体 `8.85 → 8.93`
- `Aurora Mesh` 视觉回归保持 `9.00/10` 不变
- 回归套件全量通过 `python3 scripts/test-export.py`

### v1.6.0 更新重点

- 导出管线重构成显式多段合同（`analyze → profile → slide plan → geometry plan → render`）：
  - 新增 `analyze_source` 阶段，先产出 raw signal bundle，再进入任何渲染决策
  - 新增 profile 阶段，强约束 contract evidence 必须来自本地（preset attribution + tier precedence）
  - 新增 slide planning 层，per-slide plan state 隔离，规划副作用不再串到几何阶段
  - 新增 pptx geometry planning 阶段，独占布局决策并配套强化的 stage 合同
  - render 阶段成为 geometry plan 的纯消费者，不再二次重算布局
- vendored `slide-creator` presets 同步刷新；`slide root discovery` 提到共享 helper，覆盖 generic section deck
- `Aurora Mesh` 视觉对比 `9.00/10`：
  - 弃用接近黑色的背景退化，按授权 mesh 层推导出 atmospheric 单色近似
  - KPI 默认 compact，仅当源 CSS 显式声明 stretch 时才扩展
  - 保住 wrapper 居中和 install-card 结构，`overflow = 0`、`overlap = 0`
- 回归套件全量通过 `python3 scripts/test-export.py`

### 已知边界

- 还没有实现和 native golden 完全一致
- 当前 `Swiss Modern` 仍是 `canonical + constrained compatible`，不是任意 Swiss-like HTML 的通用保真引擎
- 当前低分主要集中在：
  - canonical Swiss 的 `title_grid / column_content` 光学节奏仍有漂移
  - 若干页面的标题尺度与版心还没完全贴齐 reference
  - 本机 office-render compare 仍不是理想的最终视觉判分器

## English

### What This Is

`kai-export-ppt-lite` converts HTML presentations into editable `.pptx` files without Playwright, Chrome, or Node.js. It is built for environments where a browser-based renderer cannot be installed, but editable PowerPoint output is still required.

The current exporter is centered on:

- HTML/CSS parsing via `BeautifulSoup` and `lxml`
- Native, editable PPTX generation via `python-pptx`
- Browser-free layout approximation for text, cards, lists, tables, presentation rows, and inline code/kbd/badge patterns
- Regression tests and evaluation tooling that push the exporter toward reusable capabilities instead of deck-specific patches

### Current Status

The current validation anchors are:

- `Blue Sky`
  - Source HTML: `demo/blue-sky-zh.html`
  - Local golden reference: `demo/blue-sky-golden-native.pptx`
  - Latest verified score: `9.1/10`
- `Chinese Chan`
  - Source HTML: `demo/chinese-chan-zh.html`
  - Latest completed visual compare: `9.68/10`
  - Current structured eval:
  - `overflow = 0`
  - `overlap = 0`
  - `card containment = 0`
  - `element gaps = 0`
  - `total actionable = 0`
- `Swiss Modern`
  - Constraint path: `canonical + constrained compatible`
  - Latest real regression artifact: `kingdee_from_original_v7.pptx`
  - Latest completed visual compare: `9.36/10`
  - Confirmed file-level text semantics:
  - `P3 body wrap = square`
  - `P5 title wrap = none`

This branch has not yet reached the final goal of “every style, every slide >= 9.5”, but it is now substantially more generalized than the earlier slide-by-slide patching phase. The exporter now enforces preset-aware typography and break fidelity for decks such as `Chinese Chan` and `Swiss Modern`, not just generic geometry.

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
python3 scripts/export-sandbox-pptx.py <file.html> [demo/output.pptx] [--width 1440] [--height 900] [--with-chrome]
```

- `--width`
  Slide width in pixels. Default: `1440`
- `--height`
  Slide height in pixels. Default: `900`
- `--with-chrome`
  Add exporter-provided page counter and nav dots

### Validation Workflow

```bash
python3 -m py_compile scripts/export-sandbox-pptx.py scripts/test-export.py scripts/rigorous-eval.py
python3 scripts/test-export.py
python3 scripts/export-sandbox-pptx.py demo/blue-sky-zh.html demo/output.pptx
python3 scripts/rigorous-eval.py
```

### v1.6.4 Highlights (patch)

- `build_grid_children` now detects compact flex-row intent (`gap` + non-distributing `justify-content`) and packs children at their intrinsic content width instead of forcing an even split across the row. Fixes the `.hero-stats` KPI block on Swiss Slides 1 / 8 (P1 horizontal span `8.7" → 1.8"`)
- `_build_swiss_column_content` post-pack stretches block-level headings / paragraphs in column panels to the full inner panel width, so authored `<br>` line breaks survive instead of getting wrapped into orphan half-rows (P2 title `2.08" → 4.03"`, height taken over by `spAutoFit` to produce two clean lines)
- Three new regression tests: `test_compact_flex_row_packs_stat_blocks_at_intrinsic_width`, `test_compact_flex_row_falls_back_to_even_split_when_oversized`, `test_stretch_column_block_text_to_inner_width_expands_narrow_heading`
- `Swiss Modern` canonical snapshot stays at `9.06/10` overall (cross-renderer SSIM is near the empirical cap), but the structure now matches the authored CSS intent
- `Aurora Mesh` visual compare unchanged at `9.00/10`

### v1.6.3 Highlights (patch)

- `_build_swiss_column_content` right column now mirrors the left column's vertical centering (previously hard-coded `right_y = 0` on canonical tier)
- `_build_swiss_title_grid` resolves `justify-content` from the inner panel first (e.g. `.flow-inner { justify-content: center }`) before falling back to slide root
- `Swiss Modern` canonical snapshot: Slide 05 `8.6 → 9.0`, Slide 06 `8.9 → 9.0`, overall `9.00 → 9.06`
- `Aurora Mesh` visual compare unchanged at `9.00/10`

### v1.6.2 Highlights (patch)

- Excluded `noto sans` from `_should_apply_display_heading_boost`'s trigger so the 1.30x optical boost no longer overshoots and force-wraps Swiss headings when CJK fallback resolves to Helvetica Neue + Hiragino Sans GB
- Added `inner_panel` layout role for single-panel Swiss slides without `.bg-num` decoration (Slide 5 `flow-inner`, Slide 6 `feat-inner`)
- Pinned `_FONT_MAP` for Archivo / Nunito / Noto Sans SC to the cross-renderer Helvetica Neue + Hiragino Sans GB pair
- Compare harness strips web-font dependencies before screenshotting so SSIM reflects layout fidelity rather than cross-renderer font-rasterization
- `Swiss Modern` canonical visual snapshot: overall `8.93 → 9.00`; Slide 03/04/05/07 each `+0.1`
- `Aurora Mesh` visual compare unchanged at `9.00/10`

### Known Gap

- `≥ 9.5/slide` is not reachable under the current SSIM cross-renderer gate. Single-page SSIM is empirically capped around `0.92` (best `0.9227`, score `9.34`). Reaching `9.5+` would require switching to a structural eval gate or a shared-renderer compare loop, not more layout tuning.

### v1.6.1 Highlights (patch)

- Fixed `_pack_direct_child_content` so absolute/fixed children no longer pollute the panel's measured content width
- Added `_build_absolute_decoration_strips` to harvest Swiss `column_content` decoration strips with their authored CSS dimensions (`width / left / right / top / bottom`)
- `Swiss Modern` canonical visual snapshot: Slide 02 `8.7 → 9.3`, overall `8.85 → 8.93`
- `Aurora Mesh` visual compare unchanged at `9.00/10`
- Full regression suite passes with `python3 scripts/test-export.py`

### v1.6.0 Highlights

- Restructured the export pipeline into explicit, contract-bound stages (`analyze → profile → slide plan → geometry plan → render`):
  - new `analyze_source` stage emits raw signal bundles before any rendering decision is made
  - new profile stage tightens style-profile semantics (preset attribution + tier precedence) and requires local contract evidence
  - new `slide planning` layer isolates per-slide plan state so planning side effects do not leak into geometry decisions
  - new `pptx geometry planning` stage owns layout decisions and ships a strengthened stage contract
  - `render` is now a pure consumer of geometry plans and no longer recomputes layout
- Refreshed vendored `slide-creator` presets; lifted `slide root discovery` into a reusable helper that handles generic section decks
- `Aurora Mesh` visual compare snapshot at `9.00/10`:
  - replaced the previous near-black fallback with an atmospheric solid-color approximation derived from the authored aurora mesh layers
  - kept Aurora KPI tracks compact by default, but still honor explicit authored stretch signals
  - preserved Aurora wrapper-centered layout and install-card structure with `overflow = 0` and `overlap = 0`
- Full regression suite passes with `python3 scripts/test-export.py`

### Known Gaps

- The exporter still does not fully match the native golden deck
- `Swiss Modern` is still a `canonical + constrained compatible` path, not a generic Swiss-like HTML fidelity engine
- Remaining lower-score areas are concentrated in:
  - canonical Swiss optical rhythm on some title/split pages
  - minor title scale / page-center drift on selected layouts
  - local office-render compare limitations on this machine

See [RELEASE.md](./RELEASE.md) for release notes.
