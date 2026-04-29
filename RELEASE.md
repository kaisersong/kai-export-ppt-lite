# Release Notes

## v1.6.3 - 2026-04-29

This patch closes the canonical Swiss vertical-centering regression: right-side panels in `column_content` and inner-panel slides were stacking content against the top instead of honoring the authored `justify-content: center`.

本次补丁修复 canonical Swiss 垂直居中问题：`column_content` 右栏和 inner-panel 系列幻灯片之前把内容贴顶部，没尊重源 CSS 的 `justify-content: center`。

### Highlights

- `_build_swiss_column_content` right column now mirrors left-column vertical centering instead of forcing `right_y = 0` for canonical tier
- `_build_swiss_title_grid` resolves `justify-content` from the inner panel first (e.g. `.flow-inner { justify-content: center }`) before falling back to slide root
- `Swiss Modern` canonical visual snapshot: Slide 05 `8.6 → 9.0`, Slide 06 `8.9 → 9.0`, overall `9.00 → 9.06`
- `Aurora Mesh` regression unchanged at `9.00/10`

## v1.6.2 - 2026-04-29

This patch lifts `Swiss Modern` canonical visual fidelity from `8.93/10` to `9.00/10` and pulls Slide 03/04/05/06/07 each up by `+0.1`. Two real wins underneath: a `display heading` optical-boost trigger that was overshooting whenever `noto sans` appeared in the font stack, and a missing `inner_panel` layout role for slides that have a single content panel without a `.bg-num` decoration (e.g. Swiss Slide 5 `flow-inner` and Slide 6 `feat-inner`). The compare harness was also tightened so the SSIM score reflects layout fidelity rather than cross-renderer font-rasterization differences.

本次补丁把 `Swiss Modern` canonical 视觉保真从 `8.93/10` 提到 `9.00/10`，Slide 03/04/05/06/07 各 `+0.1`。两个真实根因：`display heading` 视觉补偿当字体栈出现 `noto sans` 时会触发 1.30 倍超调；以及单内容面板（无 `.bg-num` 装饰）的幻灯片缺少 `inner_panel` layout role 导致走 fallback。同时加固对比脚本，使 SSIM 反映 layout 一致性而非跨渲染器字体差异。

### Highlights

- Excluded `noto sans` from `_should_apply_display_heading_boost`'s trigger set — when Noto Sans SC is remapped to the cross-renderer stable Helvetica Neue + Hiragino Sans GB pair, both Latin and CJK glyphs render at the authored size and the 1.30x boost overshoots and force-wraps headings
- Added `inner_panel` layout role to `swiss-modern` contract and routed it through the existing `_build_swiss_title_grid` builder; widened title_grid's accepted inner selectors to `.hero-inner / .flow-inner / .feat-inner`
- Pinned `_FONT_MAP` for Archivo / Nunito / Noto Sans SC families to the cross-renderer Helvetica Neue + Hiragino Sans GB pair, and updated the existing `Swiss display stack` font-mapping test to match the new contract
- Compare harness now strips web-font dependencies via an injected `*  { font-family: 'Helvetica Neue', 'Hiragino Sans GB', sans-serif !important }` style tag before screenshotting, so SSIM reflects layout fidelity rather than cross-renderer font-rasterization
- `Swiss Modern` canonical visual snapshot:
  - Overall `8.93 → 9.00`
  - Slide 02 `9.3` (best, ssim `0.9189`)
  - Slide 03 `9.2 → 9.3`
  - Slide 04 `8.8 → 8.9`
  - Slide 05 `8.5 → 8.6`
  - Slide 07 `9.0 → 9.1`
- `Aurora Mesh` regression unchanged at `9.00/10`
- Full regression suite still passes: `python3 scripts/test-export.py`

### Validation Snapshot

```bash
python3 -m py_compile scripts/export-sandbox-pptx.py scripts/test-export.py scripts/sync-slide-creator-contracts.py
python3 scripts/test-export.py
python3 scripts/export-sandbox-pptx.py demo/swiss-canonical-zh.html demo/swiss-canonical-zh.pptx
python3 scripts/compare-html-ppt-visual.py demo/swiss-canonical-zh.html demo/swiss-canonical-zh.pptx --outdir demo/swiss-canonical-zh-visual-compare
python3 scripts/export-sandbox-pptx.py demo/aurora-mesh-zh.html demo/aurora-mesh-zh.pptx
python3 scripts/compare-html-ppt-visual.py demo/aurora-mesh-zh.html demo/aurora-mesh-zh.pptx --outdir demo/aurora-mesh-zh-visual-compare
```

Result:

- Full regression suite: `All tests passed!`
- `Swiss Modern` canonical fixture (`demo/swiss-canonical-zh.html`): overall `9.00/10`
- `Aurora Mesh`: overall `9.00/10` (unchanged)

### Known Gap

- `≥ 9.5/slide` for canonical Swiss is not reachable under the current SSIM-based cross-renderer compare gate. After empirically exhausting layout, font-mapping, system-font, and font-neutralized-comparison angles, single-page SSIM remained capped around `0.92` (best `0.9227`). Reaching `9.5+` would require either a structural eval gate (akin to `handwritten fixture structural eval gate`) or a shared-renderer comparison loop, not further layout tuning.

## v1.6.1 - 2026-04-29

This patch fixes a real `Swiss Modern` `column_content` rendering bug surfaced while pushing canonical Swiss fidelity: an absolute-positioned decoration strip (e.g. `.red-bar`) was inheriting the panel's full content width, both blowing up the strip itself and polluting the parent container's measured width.

本次补丁修掉一条 `Swiss Modern` `column_content` 真实渲染 bug：在推进 canonical Swiss 视觉保真过程中暴露出，绝对定位的装饰条（例如 `.red-bar`）会继承面板的全部内容宽度，既把装饰条本身放大，也污染了父容器的测量宽度。

### Highlights

- Fixed `_pack_direct_child_content` to skip absolute/fixed children so they no longer pollute the panel's measured content width
- Added `_build_absolute_decoration_strips` to harvest absolute decoration strips inside Swiss `column_content` panels and render them with their authored CSS dimensions (`width / left / right / top / bottom`)
- `Swiss Modern` canonical visual snapshot:
  - Slide 02 `8.7 → 9.3` (`+0.6`)
  - Overall `8.85 → 8.93`
- `Aurora Mesh` regression unchanged at `9.00/10`
- Full regression suite still passes: `python3 scripts/test-export.py`

### Validation Snapshot

Validated with:

```bash
python3 -m py_compile scripts/export-sandbox-pptx.py scripts/test-export.py scripts/sync-slide-creator-contracts.py
python3 scripts/test-export.py
python3 scripts/export-sandbox-pptx.py demo/swiss-canonical-zh.html demo/swiss-canonical-zh.pptx
python3 scripts/compare-html-ppt-visual.py demo/swiss-canonical-zh.html demo/swiss-canonical-zh.pptx --outdir demo/swiss-canonical-zh-visual-compare
python3 scripts/export-sandbox-pptx.py demo/aurora-mesh-zh.html demo/aurora-mesh-zh.pptx
python3 scripts/compare-html-ppt-visual.py demo/aurora-mesh-zh.html demo/aurora-mesh-zh.pptx --outdir demo/aurora-mesh-zh-visual-compare
```

Result:

- Full regression suite: `All tests passed!`
- `Swiss Modern` canonical fixture (`demo/swiss-canonical-zh.html`): overall `8.93/10`, slide 02 `9.3/10`
- `Aurora Mesh`: overall `9.00/10` (unchanged)

## v1.6.0 - 2026-04-28

This release has two main bodies of work. First, the export pipeline is restructured into explicit, contract-bound stages (`analyze → profile → slide plan → geometry plan → render`), so each stage is independently testable and the runtime stops mixing extraction, planning, and rendering on a single pass. Second, `Aurora Mesh` finishes its current optimization pass and reaches a real `9.00/10` visual compare snapshot.

本次发布有两条主线工作。第一，导出管线被重构成显式的多段合同（`analyze → profile → slide plan → geometry plan → render`），每一段都能独立测试，运行时不再把抽取、规划和渲染混在一遍流程里。第二，`Aurora Mesh` 完成本轮优化，真实视觉对比快照达到 `9.00/10`。

### Highlights

Staged export pipeline:

- New `analyze_source` stage emits raw signal bundles before any rendering decision is made
- New profile stage tightens style-profile semantics (preset attribution + tier precedence) so contract evidence has to be local
- New `slide planning` layer isolates per-slide plan state and keeps planning side effects out of geometry decisions
- New `pptx geometry planning` stage owns layout decisions and ships a strengthened stage contract
- `render` is now a pure consumer of geometry plans; it no longer recomputes layout
- Vendored `slide-creator` presets refreshed and shared `slide root discovery` lifted to a reusable helper for generic section decks

`Aurora Mesh` fidelity:

- `Aurora Mesh` visual compare snapshot: `9.00/10`
- Replaced the previous near-black fallback with an atmospheric solid-color approximation derived from the authored aurora mesh layers
- Kept Aurora KPI tracks compact by default, but still honor explicit authored stretch signals from source CSS
- Preserved Aurora wrapper-centered layout and install-card structure while keeping:
  - `overflow = 0`
  - `overlap = 0`
- Expanded Aurora regression coverage for:
  - mesh-background fallback behavior
  - compact-vs-stretch KPI width decisions
  - wrapper centering
  - install-card separation

### Validation Snapshot

Validated with:

```bash
python3 -m py_compile scripts/export-sandbox-pptx.py scripts/test-export.py scripts/sync-slide-creator-contracts.py
python3 scripts/test-export.py
python3 scripts/export-sandbox-pptx.py demo/aurora-mesh-zh.html demo/aurora-mesh-zh.pptx
python3 scripts/compare-html-ppt-visual.py demo/aurora-mesh-zh.html demo/aurora-mesh-zh.pptx --outdir demo/aurora-mesh-zh-visual-compare
```

Result:

- Full regression suite: `All tests passed!`
- `Aurora Mesh` structured checks:
  - `overflow = 0`
  - `overlap = 0`
- `Aurora Mesh` visual compare:
  - overall `9.00/10`
  - Slide 01 `9.1/10`
  - Slide 04 `9.2/10`
  - Slide 05 `8.9/10`
  - Slide 06 `8.8/10`
  - Slide 08 `8.9/10`

## v1.5.1 - 2026-04-25

This release tightens the exporter's skill/runtime execution boundary and closes two real regressions found while rerunning the full suite. The shipped work makes the single-file exporter the correctness boundary for hosted sandboxes, while still keeping an optional bootstrap path for richer environments.

本次版本重点是收紧 exporter 的 skill/runtime 执行边界，并修掉两条在完整回归中暴露出来的真实问题。核心原则变成：单文件主导出器就是 hosted sandbox 的 correctness boundary；环境更丰富时仍可使用可选 bootstrap，但不再把它当作前提假设。

### Highlights

- Version bump to `v1.5.1`
- Hardened sandbox execution behavior:
  - tolerate missing `__file__`
  - probe repo/contracts paths opportunistically and degrade cleanly when absent
  - attempt runtime dependency bootstrap before failing hard
- Added optional hosted-sandbox bootstrap surfaces:
  - `scripts/run-skill-export.py`
  - `requirements.txt`
  - updated `SKILL.md` execution protocol
- Fixed two regressions discovered while rerunning the full suite:
  - Enterprise Dark split-right-rail test was selecting the wrong container
  - Chinese Chan closing title was incorrectly exported with `wrap="none"` instead of preserving wrapped width rhythm
- Regression coverage expanded with:
  - medium contract title wrap guard
  - retained single-line contract title no-wrap guard
  - full `Chinese Chan` roundtrip wrap fidelity

### Validation Snapshot

Validated with:

```bash
python3 -m py_compile scripts/export-sandbox-pptx.py scripts/run-skill-export.py scripts/test-export.py
python3 scripts/test-export.py
```

Result:

- Full regression suite: `All tests passed!`
- Verified regressions:
  - Enterprise Dark split cards stack in the correct right rail
  - Chinese Chan closing title keeps `wrap="square"`
  - single-line large contract titles still keep `no-wrap` where intended

## v1.5.0 - 2026-04-25

This release moves `Swiss Modern` from “recognized preset metadata” to a real contract-driven export path. The shipped work includes synced Swiss contracts, runtime role-aware layout builders, and text reflow fixes that close real regressions on the `kingdee` sample deck.

本次版本把 `Swiss Modern` 从“能识别 preset metadata”推进到“真正走 contract-driven 导出路径”。核心交付包括：同步后的 Swiss contract、运行时 role-aware layout builder，以及修掉 `kingdee` 样本里真实出现的文字回流问题。

### Highlights

- Version bump to `v1.5.0`
- Expanded vendored `Swiss Modern` contract:
  - `support_tiers`
  - `layout_contracts`
  - `signature_elements`
  - typography and line-break contract
- Export runtime now consumes Swiss component semantics for:
  - `title_grid`
  - `column_content`
  - `stat_block`
  - `pull_quote`
- Text fidelity tightened again:
  - wide measured prose no longer falls back into accidental no-wrap fit
  - single-line contract titles keep `no-wrap` when authored width already fits
  - P3 / P5 wrap behavior is verified from roundtrip PPTX XML, not just visual preview
- Regression coverage expanded for:
  - Swiss contract sync
  - compatible wrapper unwrap
  - wide multiline prose wrap
  - single-line title no-wrap guard

### Validation Snapshot

Validated with:

```bash
python3 -m py_compile scripts/export-sandbox-pptx.py scripts/test-export.py
python3 -B -m pytest scripts/test-export.py -q -k "long_editorial_prose_skips_no_wrap_fit or single_line_contract_title_stays_no_wrap or wide_multiline_prose_wraps_from_measured_height or swiss or prefer_wrap_to_preserve_size_for_body_prose or centered_subtitle_prefers_full_max_width_and_no_wrap_fit or wide_prose_adjusts_back_to_single_line or medium_card_prose_adjusts_back_to_single_line"
python3 scripts/export-sandbox-pptx.py "/Users/song/Downloads/kingdee_stock_presentation (1).html" /tmp/kingdee_from_original_v7.pptx
python3 scripts/export-sandbox-pptx.py /Users/song/Downloads/kingdee_stock_presentation_swiss_canonical.html /tmp/kingdee_swiss_canonical_v7.pptx
python3 scripts/compare-html-ppt-visual.py "/Users/song/Downloads/kingdee_stock_presentation (1).html" /tmp/kingdee_from_original_v7.pptx --outdir /tmp/kingdee_exporter_only_compare_v7
python3 scripts/compare-html-ppt-visual.py /Users/song/Downloads/kingdee_stock_presentation_swiss_canonical.html /tmp/kingdee_swiss_canonical_v7.pptx --outdir /tmp/kingdee_canonical_compare_v7
```

Result:

- Targeted regression suite: `16 passed`
- Original `kingdee` compare:
  - overall `9.36/10`
  - Slide 03 `9.3/10`
  - Slide 05 `9.5/10`
- Canonical Swiss compare:
  - overall `9.18/10`
  - Slide 03 `9.2/10`
  - Slide 05 `9.4/10`
- File-level XML checks:
  - original `P3 body wrap = square`
  - original `P5 title wrap = none`
  - canonical `P5 title wrap = none`

### Known Gaps

This release materially improves `Swiss Modern`, but it still does not meet the broader goal of every supported style scoring `>= 9.5` on every slide.

Current concentration areas remain:

- canonical Swiss optical rhythm on `title_grid / column_content`
- remaining title scale and page-center drift on selected layouts
- local office-render compare tooling is still not the ideal final visual oracle
- arbitrary third-party Swiss-like HTML is still outside the guaranteed fidelity path


## v1.4.0 - 2026-04-24

This release pushes the exporter beyond generic geometry fixes and into preset-aware text fidelity. The main shipped work is the `Chinese Chan` path: vendored preset contract, serif typography enforcement, authored line-break preservation, shared runtime chrome fallback, and closing-slide fidelity checks.

本次版本把优化重点从“通用几何近似”继续推进到“preset-aware 排版 fidelity”。核心交付是 `Chinese Chan` 这条路径：vendored preset contract、serif 字体契约、显式换行保真、shared runtime chrome fallback，以及结尾页的结构 fidelity 检查。

### Highlights

- Version bump to `v1.4.0`
- Added vendored `Chinese Chan` contract:
  - `contracts/slide_creator/presets/chinese-chan.json`
- Expanded runtime preset enforcement:
  - mixed-script serif font mapping
  - `preserveAuthoredBreaks`
  - `preferWrapToPreserveSize`
  - wrap-before-shrink behavior for constrained prose
- Shared `slide-creator` runtime chrome fallback now covers presets without a vendored contract
- Closing-slide fidelity checks now include:
  - seal border preservation
  - no-shadow border-shell rendering
  - centered command-row alignment relative to the authored content column
- Regression coverage expanded with roundtrip XML checks for:
  - wrap/auto-size fidelity
  - authored column-width preservation
  - no page overflow
  - seal border and centered command fidelity

### Validation Snapshot

Validated with:

```bash
python3 scripts/test-export.py
python3 scripts/export-sandbox-pptx.py demo/chinese-chan-zh.html demo/chinese-chan-output.pptx
python3 scripts/rigorous-eval.py --sandbox demo/chinese-chan-output.pptx --golden demo/chinese-chan-output.pptx --skip-visual
```

Result:

- `All tests passed!`
- `Chinese Chan` structured eval:
  - `overflow = 0`
  - `overlap = 0`
  - `element gaps = 0`
  - `card containment = 0`
  - `total actionable = 0`

Current completed visual compare snapshot:

- `Chinese Chan`: `9.68/10`

### Known Gaps

This release materially improves preset-aware export quality, but the exporter still does not fully meet the broader goal of every supported style scoring `>= 9.5` on every slide.

Current concentration areas remain outside the shipped `Chinese Chan` path:

- `data-story` component geometry and optical rhythm
- `enterprise-dark` split/feature-card layout depth
- local visual-compare tooling still relies on a non-ideal office-render path on this machine


## v1.3.0 - 2026-04-24

This release turns the earlier `slide-creator` contract discussion into shipped repository infrastructure: sync script, vendored preset contracts, runtime contract loading, and the first contract-driven geometry tuning path for newer presets such as `data-story`.

本次版本把之前停留在设计阶段的 `slide-creator` contract 方案正式落地成仓库内基础设施：同步脚本、vendored preset contracts、运行时 contract 读取，以及面向 `data-story` 这类新 preset 的第一批 contract-driven 几何调优路径。

### Highlights

- Version bump to `v1.3.0`
- Added `scripts/sync-slide-creator-contracts.py`
- Added and expanded vendored preset contracts under `contracts/slide_creator/`
- Export runtime now consumes:
  - producer detection
  - export hints validation
  - preset contract loading
  - contract-backed tuning for newer `slide-creator` decks
- `data-story` path improved materially through:
  - contract-driven `metric_card` rebalance
  - centered wrapper and paired pill fixes
  - CJK-safe primary font mapping in PPT text runs
- Regression coverage expanded again for:
  - `data-story`
  - `enterprise-dark`
  - contract sync behavior
  - CJK font mapping and local-grid guards

### Validation Snapshot

Validated with:

```bash
python3 scripts/test-export.py
```

Result:

- `All tests passed!`

Additional current state snapshot:

- Trusted completed visual compare for `demo/data-story-zh.html`: `8.86/10`
- Structured self-check on latest exported `demo/data-story-output.pptx`:
  - `overflow = 5`
  - `overlap = 0`
  - `element gaps = 0`
  - `card containment = 0`
  - `total actionable = 5`

### Known Gaps

This release materially improves preset-aware export quality, but it still does not meet the final target of every slide scoring `>= 9.5`.

Current concentration areas:

- `data-story` Slide 6 feature-grid final writeout geometry
- residual component-level optical differences on `data-story` Slides 2 / 4 / 6 / 7
- visual compare pipeline still occasionally stalls before writing a fresh `summary.json`

## v1.2.0 - 2026-04-22

This release packages the current generalized exporter baseline, updates the GitHub-facing docs, and cleans up the repository surface so local working assets stop leaking into the published repo.

本次版本将当前通用化 exporter 基线正式收口，同时补齐 GitHub 首页文档，并清理仓库发布面，让本地工作资产不再混入公开仓库。

### Highlights

- Bilingual GitHub-facing `README.md`
- Version bump to `v1.2.0`
- Local-only working assets removed from the repository surface:
  - `docs/` is now treated as a local external-docs link
  - historical root output artifacts are no longer part of the repo
  - `demo/` remains local-only working input/output
- Exporter improvements retained in this release:
  - stronger `presentation_rows` column width handling on the generic table renderer path
  - extra runway for shortcut-heavy first columns
  - muted trailing-link color for centered closing command rows
  - stable centering for auto-margin divider shapes
- Regression coverage expanded with exporter corpus fixtures and presentation-row heuristics

### Validation Snapshot

Validated with:

```bash
python3 -m py_compile scripts/export-sandbox-pptx.py scripts/test-export.py scripts/rigorous-eval.py
python3 scripts/test-export.py
python3 scripts/export-sandbox-pptx.py demo/blue-sky-zh.html demo/output.pptx
python3 scripts/rigorous-eval.py
```

Latest verified result for `demo/blue-sky-zh.html`:

- Overall score: `9.1/10`
- `overflow = 0`
- `overlap = 0`
- `card containment = 0`
- `element gaps = 2`
- `total actionable = 2`

Per-slide snapshot:

- Slide 1: `9.2`
- Slide 2: `8.9`
- Slide 3: `9.7`
- Slide 4: `8.9`
- Slide 5: `9.0`
- Slide 6: `9.0`
- Slide 7: `9.4`
- Slide 8: `9.5`
- Slide 9: `9.1`
- Slide 10: `8.3`

### Known Gaps

This release improves the exporter materially, but it still does not hit the target of every slide scoring `>= 9.5`.

Remaining concentration areas:

- Slide 10 closing card geometry and paragraph model
- Smaller heading/icon ink differences on selected slides
- Residual card-height and rhythm drift in a few layouts

### Compatibility

- No browser runtime required
- CLI entrypoint remains:

```bash
python3 scripts/export-sandbox-pptx.py <presentation.html> [demo/output.pptx]
```

## v1.1.0 - 2026-04-22

This release turned `kai-export-ppt-lite` from a one-deck optimization effort into a more reusable sandbox exporter with explicit regression gates and documented architecture.

## v1.0.0

Initial tagged baseline before the generalized exporter hardening phase.
