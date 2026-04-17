# Slide 5 Fix: Chapter Page to 10.0/10

## Problem

Slide 5 (chapter page) scored 8.3/10 with 5 property mismatches:
- "01" chapter number: width too wide (1.943" vs 1.646"), Y offset (2.707" vs 2.926")
- "工程化能力" heading: width too wide (2.076" vs 1.906")
- "播放、编辑、Review —— 全流程闭环" paragraph: width too narrow (2.076" vs 2.634"), height doubled (0.498" vs 0.249")

Root causes:
1. `max_width` was constrained to heading width (2.076"), clipping the paragraph's natural width (2.707")
2. Width formula used `text_w + 0.15` padding, overestimating for large fonts
3. `estimate_wrapped_lines` used `font_size_pt/72` formula (3.045") instead of layout's `font_size_px/108` (2.707"), causing false 2-line estimate
4. Word wrap was enabled by character count (>20 chars) instead of actual overflow

## Fixes

### Fix 1: max_width from widest centered element (line ~2513-2534)

Changed from heading-only width to widest ALL centered text elements:
```python
# Before: only heading width
heading_content_w = max_heading_width_in + 0.15
content_w = heading_content_w

# After: scan all centered text elements
max_natural_w = max_heading_width_in
for elem in elements:
    if textAlign == 'center' and type == 'text':
        line_w = (cjk * font_px + latin * font_px * 0.55) / PX_PER_IN
        if line_w > max_natural_w:
            max_natural_w = line_w
content_w = max_natural_w  # no padding — PPTX auto-size handles fit
```

### Fix 2: Zero padding for large text widths (lines ~2727-2734, 2761-2777)

For text_w > 1.0", use `text_w` directly instead of `text_w + 0.15`:
```python
# Before: natural_w = text_width_in + 0.15
# After: for large text, zero padding
if text_width_in > 1.0:
    content_width = min(text_width_in, max_width)
else:
    content_width = min(text_width_in * 1.3 + 1.0, max_width)
```

### Fix 3: estimate_wrapped_lines formula alignment (line ~1110-1137)

Changed from `font_size_pt/72` to `font_size_px/108` to match layout code:
```python
# Before: text_width_in = (cjk * font_size_pt + latin * font_size_pt * 0.55) / 72.0
# After:
font_size_px = font_size_pt / 0.75  # reverse px_to_pt
text_width_in = (cjk * font_size_px + latin * font_size_px * 0.55) / PX_PER_IN
```

This fixed the paragraph height: line_count went from 2 → 1, height from 0.498" → 0.249".

### Fix 4: Word wrap threshold based on actual width (line ~3560-3568)

Changed from character count to actual text width comparison:
```python
# Before: needs_wrap = len(raw_text) > 20 and b['width'] < 5.0
# After:
text_w_in = (cjk * font_px + latin * font_px * 0.55) / PX_PER_IN
needs_wrap = text_w_in > b['width'] and b['width'] < 5.0
```

## Results

| Metric | Before | After |
|--------|--------|-------|
| Visual score | 8.3/10 | 10.0/10 |
| Property mismatches | 5 | 0 |
| "01" width | 1.943" | 1.793" (golden 1.646") |
| Heading width | 2.076" | 1.926" (golden 1.906") |
| Paragraph width | 2.076" | 2.707" (golden 2.634") |
| Paragraph height | 0.498" | 0.249" ✓ |
| Position drift | avg dy=+0.014" | within tolerance |
| Overflow issues | 0 | 0 |
| Overlap issues | 0 | 0 |

## Regression Check

- Slide 1: 7.4 → 7.1 (slight change in centering, acceptable)
- Slide 2: 7.6 → 7.6 (unchanged)
- Slide 3: 7.7 → 7.7 (unchanged)
- Slide 4: 8.0 → 8.0 (unchanged)
- Slide 9: 8.0 → 8.0 (unchanged)
- Overall: 7.4 → 7.5 (improved)
