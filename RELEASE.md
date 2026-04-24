# Release Notes

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
