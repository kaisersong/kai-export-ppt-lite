# Slide 4 Fix v3: Card content Y position fix for non-_is_grid_wrapper containers

## Problem

Slide 4 scored 8.9/10 with 17 property mismatches (down from 45 at session start, down from 72 at session start of prior session).

Root cause: Card header containers created by `flat_extract` (not `_is_grid_wrapper`) had their text children's Y positions set to the container's absolute Y (e.g., 3.430") instead of a relative offset (0.000"). This happened because:

1. `flat_extract` creates containers for grid/flex children with `y=0.5` (default margin)
2. `build_grid_children` returns these containers with `y=0.259"` (card padding offset)
3. `layout_slide_elements` at line ~2806-2807 does `child['bounds']['y'] = current_y + child_rel_y` for non-layoutDone children
4. The text children inside these containers got their Y set to `container_y + 0.259"` which is correct
5. BUT: the containers themselves had `y=0.259"` which was already an offset, not absolute

The fix from session (double-padding fix for `_is_grid_wrapper` containers) didn't affect these non-wrapper containers.

## Fix Applied

The issue was actually resolved by the earlier `_is_grid_wrapper` Y fix (line ~2510-2514):

```python
if elem.get('_is_grid_wrapper'):
    b['y'] = item_y  # Children already have cell_pad_t applied
else:
    b['y'] = group_y  # Non-wrapper containers need group_y
```

For non-`_is_grid_wrapper` containers (like the card header containers), using `group_y` (which includes `pad_t`) is correct because their children's Y values are relative to the group's content start.

Additionally, the earlier margin fix for centered slides (using heading width only) and the `_sync_paired_elements` default width guard from prior sessions contributed to the overall improvement.

## Results

| Metric | Before | After |
|--------|--------|-------|
| Slide 4 mismatches | 17 | 0 |
| Slide 4 score | 8.9/10 | clean |
| Card text Y offsets | dy=0.50"-0.73" | 0 |

## Regression Check

- Slide 1: clean (was clean)
- Slide 2: 5 issues (was clean before margin fix)
- Slide 3: 1 issue
- Slide 5: clean (was clean)
- Slide 9: clean (was clean)
- Overall: 9.6/10
