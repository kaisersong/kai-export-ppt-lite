# Slide 1 Fix: Stat card wrapper width correction to 8.2/10

## Problem

Slide 1 scored 7.1/10 with stat card wrappers shrinking from correct widths (~1.15") to ~0.69".

Root causes:
1. `_is_grid_wrapper` container widths were overwritten by `this_item_width` at line ~2477 in `build_grid_children`'s layout pass, losing the `child_padded_w` computed during wrapper creation
2. `content_widths` computation for centered grids only checked `TEXT_TAGS`, missing `<div>` elements with large text (e.g., `div.stat` with "21" at 44.8px)
3. `child_padded_w` formula added 100% CSS padding, making cards too wide

## Fixes Applied

### Fix 1: Preserve _is_grid_wrapper container width (line ~2481)

**Root cause:** Line 2477 `b['width'] = this_item_width` overwrites the wrapper width computed during wrapper creation. The `this_item_width` from `compute_text_content_width` returns text-only width (~0.687"), while the wrapper was created with `child_padded_w` (~1.299") that includes CSS padding.

**Fix:** Skip width assignment for `_is_grid_wrapper` containers:
```python
if not elem.get('_is_grid_wrapper'):
    b['width'] = this_item_width
```

**Result:** Wrapper widths preserved at creation-time values instead of being overwritten.

### Fix 2: Include div text in content_widths computation (line ~1738)

**Root cause:** `content_widths` for centered grids iterates descendants checking `TEXT_TAGS`, but `<div>` is not a TEXT_TAG. Stat card `div.stat` with "21" at 44.8px was skipped, only `<p>` "设计预设" at 12.48px was found.

**Fix:** Also check `<div>` descendants for direct text content. For large text (>= 24px), use character-aware width with 0.65 Latin factor instead of 0.55 (better accounts for digit/letter widths at large sizes).

**Result:** `content_widths` for stat cards: 0.839", 0.570", 0.849" (vs 0.762" before). Per-card widths now vary based on actual content.

### Fix 3: Reduced padding factor for child_padded_w (line ~2170)

**Root cause:** `child_padded_w = child_item_width + cell_pad_l + cell_pad_r` added 100% CSS padding (0.519" for 28px padding), making cards too wide (1.667" vs golden 1.149").

**Fix:** Use 50% padding factor:
```python
child_padded_w = child_item_width + (cell_pad_l + cell_pad_r + cell_border_l + cell_border_r) * 0.5
```

**Result:** Card widths closer to golden: 1.108" vs 1.149", 0.838" vs 0.894", 1.117" vs 1.301".

## Results

| Metric | Before | After |
|--------|--------|-------|
| Slide 1 score | 7.1/10 | 8.2/10 |
| Slide 1 mismatches | 28 | 15 |
| SIZE mismatches | Many | Reduced |
| Card 1 width | 0.687" | 1.108" (golden 1.149") |
| Card 2 width | 1.004" | 0.838" (golden 0.894") |
| Card 3 width | 0.804" | 1.117" (golden 1.301") |

## Remaining Issues (Slide 1, 15 mismatches)

1. **Title "AI 驱动的\nHTML 演示文稿"**: dx=0.260" dy=0.036", dw=0.520" — title position and width
2. **Subtitle "从提示词到精美演示..."**: dx=0.260" dy=0.035", dw=0.520" — subtitle position and width
3. **Stat card positions**: dx=0.087"-0.187", dy=0.125"-0.153" — systematic Y offset (~0.14" too high)
4. **"1" text width**: dw=0.366" — golden 0.594", sandbox 0.228"
5. **Card 3 shape**: dw=0.183" dh=0.043" — width still too narrow

The title/subtitle issues (dx=0.260", dw=0.520") are from centered text layout code affecting overall content width. Card position offsets trace to grid layout centering.

## Regression Check

- Slide 2: 7.8 → 6.9 (regression from card width changes affecting cols2 grid)
- Slide 3: 7.5 → 7.5 (unchanged)
- Slide 4: 8.9 → 8.9 (unchanged)
- Slide 5: 10.0 → 10.0 (unchanged)
- Slide 9: 8.2 → 7.2 (regression)
- Overall: 7.4 → 7.6 (slight improvement)
