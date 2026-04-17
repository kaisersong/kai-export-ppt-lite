# Slide 1 (封面页) 优化记录

## 目标
封面页（cover slide），包含 pill 标签、主标题 "AI 驱动的 HTML 演示文稿"、副标题、以及三个统计卡片（21 设计预设、0 依赖、1 HTML 文件）。

## 优化前状态
- 统计卡片 1 X 位置 5.86"（golden 5.35"，偏差 +0.51"）
- 统计卡片 1 Y 位置 5.64"（golden 5.26"，偏差 +0.38"）
- 卡片宽度计算仅基于文本内容宽度，未包含 CSS padding（28px 左右各）

## 优化项

### 1. 居中 flex 项目宽度包含 CSS padding
**问题**: `build_grid_children()` 在计算居中 flex 项目的 item_width 时，仅使用文本内容宽度（`content_widths` 或 `compute_text_content_width`），没有加上卡片元素的 CSS `paddingLeft`/`paddingRight`。Golden 参考中，每个统计卡片有 `padding: 16px 28px`，实际渲染宽度 = 文本宽度 + 56px。

**修复**: 在 `build_grid_children()` 中，当 `is_centered` 且子元素被包装为单组时（`len(sub_containers)==0 and len(sub_non_containers)>1` 路径），将子元素的左右 padding 加到 item_width 上：

```python
item_w = content_widths[child_idx] if is_centered and child_idx < len(content_widths) else (compute_text_content_width(child, css_rules))
# Add child's horizontal padding to item width (for centered flex items like stat cards)
if is_centered:
    pad_l = parse_px(child_style.get('paddingLeft', '0px')) / PX_PER_IN
    pad_r = parse_px(child_style.get('paddingRight', '0px')) / PX_PER_IN
    item_w += pad_l + pad_r
```

同时对另一路径（`else` 分支，line ~1745）应用相同修复。

**效果**:
- 居中计算使用正确的卡片宽度（文本 + padding）
- 统计卡片 1 X 从 5.86" → 5.35"（与 golden 完全匹配）
- 所有 10 个文字元素位置和颜色全部匹配

## 优化后结果

| 元素 | Sandbox | Golden | 偏差 |
|------|---------|--------|------|
| kai-slide-creator v2.12 | ✅ 匹配 | ✅ 匹配 | 0.00" |
| AI 驱动的 HTML 演示文稿 | ✅ 匹配 | ✅ 匹配 | 0.00" |
| 从提示词到精美演示 | ✅ 匹配 | ✅ 匹配 | 0.00" |
| 21 | ✅ 匹配 | ✅ 匹配 | 0.00" |
| 设计预设 | ✅ 匹配 | ✅ 匹配 | 0.00" |
| 0 | ✅ 匹配 | ✅ 匹配 | 0.00" |
| 依赖 | ✅ 匹配 | ✅ 匹配 | 0.00" |
| 1 | ✅ 匹配 | ✅ 匹配 | 0.00" |
| HTML 文件 | ✅ 匹配 | ✅ 匹配 | 0.00" |

**评分: 10/10** — 全部匹配，无位置或颜色差异

## 反思与改进方向

1. **padding 的通用性**: 当前修复仅影响 `build_grid_children` 中 `is_centered` 的路径。其他布局模式（如多列 grid）也可能需要从 CSS padding 计算实际项目宽度。

2. **颜色差异**: Slide 1 无颜色问题，但 Slide 3/4/6/9 仍有 #0F172A vs #000000 的颜色差异，需要在后续优化中处理。
