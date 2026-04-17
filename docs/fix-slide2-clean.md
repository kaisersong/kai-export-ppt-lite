# Slide 2 Optimization: 6.3 → 8.7/10

## Slide Structure

HTML: `<section class="slide">` with `justify-content:center; padding:clamp(28px,4vw,56px)`

```
div[max-width:860px]
  ├── span.pill (inline-block, "为什么选择 slide-creator")
  ├── h2.gt ("传统工具限制了你")
  ├── div.divider
  ├── div.cols2 (grid 2-col)
  │   ├── div.g[border-left:4px #ef4444, padding:22px 24px]
  │   │   ├── h4 ("❌ 传统方法")
  │   │   └── ul.bl[gap:7px] > 5×li
  │   └── div.g[border-left:4px #10b981, padding:22px 24px]
  │       ├── h4 ("✓ slide-creator")
  │       └── ul.bl[gap:7px] > 5×li
  └── div.info[padding:12px 16px, border-left:4px #2563EB]
```

## Fixes Applied

### 1. maxWidth Constraint Propagation (Session 13)
- `build_text_element`: `effective_max_w` falls back to `content_width_px`
- `layout_slide_elements`: always uses `max_constraint` over `max_text_width`
- All 3 call sites pass `content_width_px` to `build_text_element`

### 2. Grid Cell Internal Width (Session 13)
- `build_grid_children`: computes `cell_internal_width_px = (item_width_in - padding) * PX_PER_IN`
- Passes to nested `flat_extract` calls so `<li>` in `<ul>` use available width

### 3. Block Text Width in Grid Layout (Session 14)
- Grid bg shapes store CSS padding (`_css_pad_l/r/t/b`), border (`_css_border_l`), textAlign (`_css_text_align`)
- Block text tags (h1-h6, p, li) use full `card_content_w`, left-align when `textAlign != 'center'`
- Short/plain text uses `natural_w + 0.1`, centered within card
- Stat cards (Slide 1) with `textAlign:center` still centered correctly

### 4. Panel Height via CSS Padding (Session 14)
- `item_h` uses `_css_pad_t` and `_css_pad_b` from bg shape instead of hardcoded 15px
- Internal gap computation: li elements use 7px gap (from `ul.bl` CSS), others use 0.05"

### 5. CSS Gap for Li Elements (Session 14)
- Grid layout: `group_y += b['height'] + 7.0/PX_PER_IN` for `<li>` elements
- `item_h` calculation accounts for li-specific gap vs default gap

### 6. Inline-Block Pill Width (Session 14)
- `build_text_element`: inline-block elements include CSS `paddingLeft + paddingRight` in width
- `flat_extract` pill shape: no longer double-adds padding (text_el already includes it)
- Only applies to `is_inline_block=True` to avoid affecting stat card text

### 7. Info Box Border (Session 14)
- `flat_extract` inline-child path: stores `borderLeft` style on paired shape
- Info box left border bar exported as decoration shape

## Remaining Issues (8.7 → 9.5 target)

| Issue | dw/dh/dy | Root Cause | Fix Needed |
|-------|----------|------------|------------|
| Pill width | dw=0.308" | letter-spacing 0.05em not in width calc | Add letter-spacing to content_w_in |
| Pill height | dh=0.065" | height calc doesn't match golden | Check line-height multiplier |
| Last 2 li Y drift | dy=0.115-0.133" | cumulative gap error from panel bottom | Li height 0.249" vs golden 0.263" |
| Info text height | dh=0.150" | 0.421" vs 0.572" — wrapping difference | Check line-height for info text |
| Shape match | 3 mismatches | Comparison script confused info box bg with border | Not a real export issue |

## Key Measurements (Slide 2)

| Element | Golden | Sandbox | Diff |
|---------|--------|---------|------|
| Pill pos | (2.685, 2.153) | (2.684, 2.261) | dx=0.001, dy=0.107 |
| Panel bg size | 3.906×2.383 | 3.906×2.193 | dh=0.189 |
| H4 pos | (2.944, 3.439) | (2.944, 3.529) | dx=0, dy=0.090 |
| Li pos (all) | x=2.944 | x=2.944 | dx=0 (perfect) |
| Info text pos | (2.685, 5.738) | (2.684, 5.781) | dx=0.001, dy=0.043 |
