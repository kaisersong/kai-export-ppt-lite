# 修复形状导出系统性缺陷 — 修复方案 v2

## 当前状态

基于 `demo/blue-sky-zh.html` + `demo/blue-sky-golden-native.pptx` 验证

- **Slide 1**: ✅ 元素全匹配 (4=4 shape, 10=10 text), 26 属性不匹配 (dy ~0.25-0.49")
- **Slide 2**: ✅ 文本全匹配 (16=16), decoration 1vs2, shape 8vs7
- **Slide 3**: ✅ 元素全匹配 (4=4 dec, 10=10 shape, 13=13 text) — 无多余元素
- **Slide 4**: ✅ 文本全匹配 (36=36), decoration 6vs2, shape 28vs32 (dy ~0.3-0.7")
- **Slide 5**: ✅ 元素全匹配 (1=1 decoration, 4=4 text)
- **Slide 6**: ✅ 文本全匹配 (18=18), shape 12=12, decoration 3vs2
- **Slide 7**: ✅ 文本全匹配 (26=26), decoration 2vs1, shape 8vs9
- **Slide 8**: ✅ 文本全匹配 (30=30), decoration 1vs2, shape 5vs4
- **Slide 9**: ✅ 文本全匹配 (18=18), decoration 1vs2, shape 5vs4
- **Slide 10**: ✅ 文本全匹配 (5=5), card_bg 1vs0 (位置/颜色不同), shape 1vs2

**总计**: 176/176 文字元素匹配, 0 多余/缺失文字元素。评分 9.6/10。

### Slide 3 剩余问题

| 问题 | 现状 | 目标 |
|------|------|------|
| borderLeft 被分类为 divider_v | sandbox w=0.04" < 0.05" 阈值 | golden w=0.19" 分类为 shape |
| 第 3 层 bg shape 缺失 | G#21 at (2.87",5.43") 缺失 | 需要创建 |
| `--generate` 多余 text | ~~`<code>` 被提取为独立 pill~~ ✅ 已修复 | 不再需要 |
| 文字宽度被覆盖 | "描述你的主题" 3.73" vs 6.49" | build_grid_children 中 this_item_width 覆盖 |
| step dot 位置 | "1" at (2.87",3.69") vs (3.07",3.30") | x 偏移需要包含 step dot 的 CSS 位置 |
| 主题小圆点缺失 (Slide 4) | 4 个 0.09"×0.09" 圆点未创建 | golden 有 #0F172A/#0EA5E9/#8B5CF6/#2563EB 圆点 |

### Slide 4 剩余问题

| 问题 | 现状 | 目标 |
|------|------|------|
| 主题小圆点缺失 | sandbox 无 0.09"×0.09" 圆点 | golden 4 个圆点 at y=3.17" |
| 多余 shape | sandbox 31 vs golden 28 | 3 个多余 shape 需过滤 |
| 垂直偏移 ~0.34" | 所有元素 y 偏上 | golden 参考位置 |

## 问题根因分析

### 问题 1：多余渐变覆盖层（Slide 1/10/4）
- **根因**：`build_grid_children()` 第 1610 行创建 gradient shape 时，没有经过 `is_gradient_overlay` 过滤
- **路径**：`.divider` 元素 → `flat_extract` → `build_grid_children` → 第 1610 行 `if has_grad:` 创建 shape
- **状态**：主路径 1358 行已修复，但 `build_grid_children` 路径未修复

### 问题 2：多余分割线（Slide 2/4/8/9）
- **根因**：同上，`.divider` 在 `build_grid_children()` 中被创建为全宽 shape
- **CSS**：`.divider { width: 56px; height: 3px; background: linear-gradient(...) }`
- **Sandbox**：宽度被强制为全宽 ~8"，颜色转为 #2563EB
- **Golden**：无此形状（gradient 装饰元素应跳过）

### 问题 3：卡片背景颜色错误
- **根因**：`slide_data['background']` 为 None 时 fallback 到 `(255,255,255)` 纯白
- **Golden**：卡片背景 `rgba(255,255,255,0.70)` 混合渐变起始色 `#f0f9ff` = `#FAFDFF`
- **Sandbox**：混合纯白 `(255,255,255)` = `#FFFFFF`

### 问题 4：装饰小圆点缺失（所有 slide）
- **根因**：由 JavaScript 动态创建 + `position: fixed`，静态 HTML 解析无法捕获
- **CSS**：`#nav-dots { position: fixed; bottom: 24px; ... }`
- **决策**：JS 动态内容不纳入导出范围（黄金参考是 Playwright 渲染，包含运行时 DOM）

### 问题 5：测试用例未覆盖
- 所有 6 个测试只覆盖 Slide 1
- 没有 shape 数量/颜色/去重检测

## 修复实施顺序

### 修复 1：build_grid_children 过滤 gradient 装饰元素 ✅ 已完成
**文件**: `scripts/export-sandbox-pptx.py`
**位置**: `build_grid_children()` 第 ~1610 行
**改动**: 在 `if has_visible_bg_or_border(child_style) or has_grad:` 前加 `is_gradient_overlay` 检查

### 修复 1b：flat_extract 中 `is_gradient_overlay` 小元素例外 ✅ 已完成
**改动**: 添加 `is_small_explicit` 检查 — 仅当 width/height 为 px 值且 < 200px/< 20px 时不过滤
**原因**: `.divider` (56px × 3px) 是可见设计元素，不应被过滤

### 修复 1c：`background-clip: text` 元素不创建 shape ✅ 已完成
**改动**: 在 leaf text container 路径和 standard container 路径中检查 `backgroundClip: text`
**原因**: `.chapter-num` 的渐变用于文字着色，不是背景形状

### 修复 1d：`gradient_to_solid` 支持 hex 颜色 ✅ 已完成
**改动**: 当 rgba 匹配为空时，提取 hex 颜色的第一个作为 solid fill
**原因**: `.divider` 使用 `#2563eb` hex 颜色，之前返回 None

### 修复 1e：小元素 shape 使用 CSS 尺寸 ✅ 已完成
**改动**: 当 CSS 有显式 px 宽度/高度时，设置 shape bounds 为对应尺寸
**原因**: `.divider` 的 golden 尺寸是 0.52"×0.03" (56px×3px)

### 修复 1f：小 shape 在 layout 中居中 ✅ 已完成
**改动**: layout 函数中为小于 max_width 50% 的 shape 添加 x 居中

### 修复 1g：居中文字宽度公式优化 ✅ 已完成
**改动**: 对于 text_width_in > 1.0 的居中元素，使用 `natural_w + 0.15"` 替代 `* 1.3 + 1.0`

### 修复 1h：百分比维度不作为小元素 ✅ 已完成
**改动**: `is_small_explicit` 只匹配 px 结尾的尺寸，忽略百分比
**原因**: 封面幻灯片的环境光球 (50%×60%) 被错误当作小元素

### 修复 1i：过滤装饰性模糊元素 ✅ 已完成
**改动**: 对 `position:absolute/fixed` + 有 `blur` filter + 无文字的元素跳过 shape 创建
**原因**: Slide 10 的 cloud orbs (300px/400px rgba) 不应成为 PPTX shapes

### 修复 1j：pill 文字保留在父级文本元素中 ✅ 已完成
**文件**: `scripts/export-sandbox-pptx.py`
**位置**: `build_text_element()` 调用、pill shape 创建、`get_text_content()`、`extract_text_segments()`
**改动**:
1. 移除 `pill_elements` set 跟踪 — 不再从父级文本中排除 pill 子元素
2. 移除 pill shape 的 `pill_text` 和 `pill_color` 字段 — pill 变为纯视觉装饰（无文本）
3. 移除 `get_text_content()` 和 `extract_text_segments()` 的 `exclude_elements` 参数
4. 移除 `build_text_element()` 的 `exclude_elements` 参数

**原因**: 黄金参考中 "Blue Sky 当前" 是合并的文本元素（两个 text runs："Blue Sky" 正常色，"当前" 不同色+粗体），pill shape 是无文本的装饰形状。之前的 `exclude_elements` 逻辑将 pill 文字从父级文本中排除，导致 "Blue Sky" 和 "当前" 分离为两个独立文本元素。

**结果**: Slide 4 的 "Blue Sky 当前" 现在是合并的文本元素，pill shape 为纯视觉装饰。文本元素计数从 37→36（匹配 golden 36）。

### 修复 1k：无文本但有显式尺寸的小元素创建装饰 shape ✅ 已完成
**文件**: `scripts/export-sandbox-pptx.py`
**位置**: `flat_extract()` 叶级文本容器路径、`build_grid_children()` 路径
**改动**:
1. 在 `is_leaf_text_container` 路径中，当 `build_text_element` 返回 None 但元素有可见背景 + 显式 CSS px 尺寸时，创建 `_is_decoration` shape
2. 在 `build_grid_children` 的 `is_leaf_text_container` 路径中，同样处理无文本但有小尺寸背景的元素
3. 在 `content_only` 过滤器中添加 `_is_decoration` 条件

**原因**: 主题小圆点 (10px×10px, border-radius: 50%, 纯色背景) 是无文本的装饰元素，之前完全被跳过。
**限制**: 如果圆点是 CSS 伪元素 (`::before`/`::after`) 或 JS 生成，静态 HTML 解析无法捕获。这是架构限制。

### 修复 2：slide background fallback 使用渐变起始色 ✅ 已完成
**文件**: `scripts/export-sandbox-pptx.py`
**位置**: `build_grid_children()` 第 1610 行
**改动**: 在 `if has_visible_bg_or_border(child_style) or has_grad:` 前加 `is_gradient_overlay` 检查

```python
# 在 line ~1607 之后，创建 shape 之前：
# Skip gradient-only decorative elements (dividers, ornament divs)
is_decorative_gradient = (
    has_grad
    and not has_visible_bg_or_border(child_style)
    and not total_txt  # no text content
    and not any(child_style.get(side, '') for side in ('border', 'borderLeft', 'borderRight', 'borderTop', 'borderBottom'))
)
if is_decorative_gradient:
    continue  # Skip decorative gradient element
```

### 修复 2：slide background fallback 使用渐变起始色
**文件**: `scripts/export-sandbox-pptx.py`
**位置**: `export_shape_background()` 调用处的 `slide_bg` fallback
**改动**: 从 `(255, 255, 255)` 改为使用 `slide_data.get('gradient_start')` 或提取渐变起始色

需要在 `extract_slide_background()` 中存储渐变起始色，并在 export 时使用。

### 修复 3：装饰小圆点 — 不修复
JS 动态生成的内容不在静态 HTML 导出范围内。Golden 参考是 Playwright 渲染（有完整运行时 DOM），沙盒只解析静态 HTML。这是架构差异，不是 bug。

### 修复 4：测试用例扩展
- 添加 Slide 4 shape 数量测试
- 添加卡片背景颜色测试
- 添加无多余分割线测试

## Slide 5 剩余问题 (9.3/10 → 9.5+ 目标)

| 元素 | 问题 | 原因 |
|------|------|------|
| "01" width | 1.94" vs golden 1.65" | font-size clamp() 解析值 (176px) 与 golden 实际渲染 (132pt=176px) 一致但宽度公式仍多 0.3" |
| "01" y-position | 2.47" vs golden 2.93" | justifyContent: center 计算的 content height 与 golden 不同 |
| "工程化能力" width | 2.08" vs golden 1.91" | 同上，clamp() 解析问题 |
| subtitle width | 2.08" vs golden 2.63" | `_full_subtitle` 未设置，使用了通用公式 |
| subtitle y-position | 5.36" vs golden 5.16" | 前面元素累积偏移 |
| decoration y-position | 5.18" vs golden 5.00" | 前面元素累积偏移 |

**根因**: clamp() CSS 函数在静态 HTML 解析时使用中间值，而 golden (Playwright 渲染) 使用视口决定的实际值。这导致 font-size 系统性差异，进而影响宽度和位置计算。

**可能方案**: 
1. 改进 clamp() 解析，使用视口宽度 (1440px) 计算实际值
2. 或在 layout 中对 chapter slides 使用特殊间距

## 对抗评审要点

1. **修复 1** 是否会影响合法的 gradient 卡片（如 Aurora Mesh 卡片的 gradient 背景）？
2. **修复 2** 的渐变起始色提取是否可靠？CSS 变量解析是否正确？
3. **修复 3** 的决策是否正确：nav-dots 是否应在后续版本中通过 JS 快照支持？
