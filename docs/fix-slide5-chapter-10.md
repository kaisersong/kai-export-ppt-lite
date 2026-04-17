# Slide 5 (Chapter Page) Optimization — 10.0/10

## Target
Optimize chapter slide (slide 5: "01 工程化能力") to 9.5/10 score.

## Baseline
- Score: ~8.8/10
- Issues: Y offset +0.14" on all elements, heading height too tall (0.65" vs golden 0.46"), extra shapes from background-clip:text, divider not centered, width formula inflated, marginBottom not applied

## Fixes Applied

### 1. Negative marginBottom handling
**Problem:** `.chapter-num` has `margin-bottom: -16px` creating a -0.148" gap between chapter number and heading. The layout engine was adding base gap (0.13") + marginBottom, resulting in all content being shifted.

**Root cause:** `parse_px('-16px')` returned 16.0 (positive) instead of -16.0 because the regex `([\d.]+)` didn't capture the minus sign.

**Fix:** Changed regex in `parse_px` from `([\d.]+)` to `(-?[\d.]+)`.

### 2. Gap logic for negative marginBottom
**Problem:** Even after parse_px fix, the layout was adding `0.13 + marginBottom` for all elements.

**Root cause:** Golden uses base gap (0.13") for normal elements, but negative marginBottom replaces the base gap entirely (not added to it).

**Fix:** When marginBottom < 0, use only the marginBottom value (no base gap added):
```python
if mb < 0:
    current_y += b['height'] + mb  # negative margin replaces base gap
else:
    current_y += b['height'] + 0.13
```

### 3. marginBottom not stored in element styles
**Problem:** `build_text_element` and `build_shape_element` didn't include `marginBottom` in the styles dict, so the layout pass couldn't find it.

**Fix:** Added `'marginBottom': style.get('marginBottom', '')` to both functions' output.

### 4. Margin shorthand not expanded
**Problem:** `.divider` has `margin: 8px 0 14px` from CSS and `margin:14px auto` from inline style, but these weren't expanded to individual margin properties.

**Fix:** Added `_expand_margin()` function (similar to `_expand_padding`) to expand margin shorthand to marginTop/MarginBottom/etc.

### 5. Font size multiplier for width calculation
**Problem:** "01" chapter number width was 1.793" vs golden 1.646" (dw=+0.147").

**Root cause:** Latin character width multiplier was 0.55x font size, but golden uses 0.5x.

**Fix:** Changed all `font_size_px * 0.55` to `font_size_px * 0.5` in width calculations (5 occurrences).

### 6. Gradient with hex colors
**Problem:** Divider had `background: linear-gradient(90deg, #2563eb, #0ea5e9)` but `gradient_to_solid` only parsed rgba() format, falling through to `shape.fill.background()`.

**Fix:** Added hex color fallback in `gradient_to_solid` to extract first hex color from gradient.

## Results

### Before → After
| Metric | Before | After | Golden |
|--------|--------|-------|--------|
| Score | ~8.8/10 | **10.0/10** | - |
| "01" Y | 2.786" | 2.925" | 2.926" |
| "01" width | 1.793" | 1.630" | 1.646" |
| Heading Y | 4.545" | 4.406" | 4.407" |
| Divider Y | 5.138" | 4.999" | 4.999" |
| Subtitle Y | 5.295" | 5.156" | 5.156" |
| Mean X drift | -0.030" | -0.002" | - |
| Mean Y drift | +0.139" | -0.001" | - |
| Max \|dx\| | 0.073" | 0.010" | - |
| Max \|dy\| | 0.139" | 0.001" | - |

### Element count match
- Golden: 1 decoration + 4 text = 5 elements (excl. nav dots/counter)
- Sandbox: 1 decoration + 4 text = 5 elements

### All gaps match golden
- "01" → heading: -0.148" (negative margin)
- heading → divider: +0.130" (base gap)
- divider → subtitle: +0.130" (base gap)
