# Slide 4 Fix v2: subtitle height + width + _sync_paired_elements fix

## Problem

Slide 4 scored 8.9/10 with 45 property mismatches (down from 72 at session start).

## Fixes Applied

### Fix 1: _sync_paired_elements default width guard

**Root cause:** `_sync_paired_elements` synced text width to shapes that still had default 12.33" width (not computed from content). The subtitle "每种预设都有完整的..." had `_pair_with` and its paired shape had 12.33" width, overwriting the layout-computed 8.704".

**Fix:** Added `shape_near_default` check (line ~3010):
```python
shape_near_default = abs(sb.get('width', 0) - 12.33) < 0.1
if shape_near_default:
    continue  # Don't sync — shape width is default, not computed
```

**Result:** Subtitle width fixed from 12.330" → 8.704" (golden 8.702").

### Fix 2: Subtitle height for padded elements

**Root cause:** The subtitle `.info` div has paddingTop=9px, paddingBottom=9px. Layout code adds this padding to bounds height (0.199" → 0.366"), but the export code created a textbox at that height with TF margins, then PPTX's TEXT_TO_FIT_SHAPE auto-size shrunk it.

**Golden:** auto_size=NONE, word_wrap=True, shape height=0.499"

**Fix:** Three-part fix:
1. In `build_text_element`, no change (padding added by layout code is correct)
2. In export code, compute `export_h = b['height'] + pad_h_in * 0.8` for padded single-line elements
3. Set bodyPr `anchor='t'` and `wrap='square'` via lxml to preserve height on save
4. Clear TF margins since export_h already includes padding
5. Force shape height after all TF margins are set

**Result:** Subtitle height fixed from 0.366" → 0.499" (matches golden exactly).

## Results

| Metric | Before | After |
|--------|--------|-------|
| Property mismatches | 72 | 45 |
| SIZE mismatches | 72 → 0 (card text) | 2 remaining (shapes) |
| Subtitle width | 12.330" | 8.704" (golden 8.702") |
| Subtitle height | 0.366" | 0.499" ✓ |
| Slide score | 8.3/10 | 8.9/10 |

## Remaining Issues (45 position mismatches)

All remaining mismatches are POSITION-related:

1. **Title "按内容类型自动匹配"**: dx=0.271" dy=0.236" — title position mismatch
2. **Card headers** ("深色", "浅色", "专业", "v1.5 新增"): systematic dy=+0.290" to +0.309" (too low)
3. **Card body text** ("Dark Botanical", etc.): systematic dy=-0.137" to -0.360" (too high)
4. **2 shape SIZE mismatches**: dw=0.316" dh=0.261"-0.305"

These position issues trace to the grid card layout code (`build_grid_children`) where:
- Card container Y positions are ~0.30" too low
- Card content Y positions within containers are ~0.14-0.36" too high

Fixing these requires changes to the grid layout engine which affects all slides.

## Regression Check

- Slide 1: 7.1 (was 7.4)
- Slide 2: 7.8 (was 7.6, improved)
- Slide 3: 7.5 (was 7.7)
- Slide 5: 10.0 (unchanged)
- Slide 9: 8.2 (was 8.0, improved)
- Overall: 7.7/10 (was 7.4)
