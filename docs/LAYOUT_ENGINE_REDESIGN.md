# Layout Engine 重构设计（V2）

## 背景

### Codex 对抗性 Review 结果

Codex 提出 10 个问题（5 Major/Critical, 5 Minor）：
- **Critical**: 评分方法有缺陷（线性扣分不区分严重度）
- **Major**: 根因分析部分错误、架构过度设计、扁平化过早否定、重构顺序错误
- **Minor**: 细节描述问题

### CSS 覆盖率审计（42 个 demo HTML）

| CSS 属性 | 使用量 | 优先级 |
|---------|--------|--------|
| `display: flex` | 413 次 | 必须 |
| `grid-template-columns` | 125 次 | 必须 |
| `justify-content: center` | 135 次 | 必须 |
| `align-items: center` | 142 次 | 必须 |
| `box-sizing: border-box` | 47 次 | 关键 |
| `grid-template-areas` | 0 次 | 可忽略 |

---

## 问题诊断（修正版）

### 第 2 页布局对比

| 元素类型 | Golden X | Sandbox X | ΔX | Golden Y | Sandbox Y | ΔY |
|---------|----------|-----------|----|----------|-----------|----|
| 标题 | 2.68" | 2.68" | 0 | 2.15" | 1.68" | -0.47" |
| 左列 | 2.94" | 2.18" | -0.76" | 3.44" | 3.13" | -0.31" |
| 右列 | 7.00" | 6.19" | -0.81" | 3.44" | 3.13" | -0.31" |
| 列宽 | 3.42" | 3.91" | +0.49" | — | — | — |

### 根因分析（修正版）

1. **Grid 列宽解析只返回列数，忽略宽度比例**
   - `_parse_grid_columns()` 正确解析 `repeat(2, 1fr)` → 返回 2
   - 但 `_compute_column_widths()` 对所有列均分（忽略 `1fr 2fr` 比例、`200px 1fr` 固定+弹性）
   - Golden 的 3.42" 列宽 ≠ Sandbox 的 3.91" 均分列宽

2. **Slide 顶部边距来自硬编码常量**
   - `internal_margin = 1.5` 无任何 CSS 依据
   - Golden 实际从 y=2.15" 开始（body padding ~60px + 内部偏移）
   - Sandbox 从 y=1.68" 开始，偏差 0.47"

3. **`box-sizing: border-box` 未考虑**
   - 所有 demo 全局设置 `*, *::before, *::after { box-sizing: border-box }`
   - 代码把 width 和 padding 相加，实际 border-box 下 width 已包含 padding

4. **`justify-content: center` 支持不完整**
   - 135 次使用，是最常用的 flex/grid 对齐方式
   - 代码只在 grid 容器级别检测，不在 flex 容器级别处理

5. **评分方法缺陷**
   - 每个位置问题固定扣 0.01 分，0.01" 偏移和 0.76" 偏移同罚
   - 装饰元素（无文字）被计数
   - 覆盖率已 100%，唯一提升空间在位置分

---

## 修正后架构

### 设计原则

- **单文件重构**：先稳定函数边界，不拆分文件
- **IR schema 优先**：先定义元素数据结构，再提取布局器
- **3 阶段**：组件布局 → 页面布局 → 诊断层
- **保留扁平化**：当前 flatten 逻辑已有 sophistication（列检测、span 处理），不全盘否定

### 数据流

```
HTML Parse (不变)
    ↓
Stage 2: 组件布局器（4 个独立函数，同文件）
    ├── layout_grid_columns()      — Grid 列计算
    ├── layout_flex_row()           — Flex 行布局
    ├── layout_flex_column()        — Flex 列布局
    └── apply_alignment()           — 对齐计算（内联到布局器）
    ↓
Stage 3: 页面布局器（1 个函数）
    ├── layout_slide_elements()     — Y 堆叠 + CSS 边距
    └── overflow_splitter()         — 已实现 ✓
    ↓
Stage 4: 诊断层（独立模块）
    ├── diff_against_golden()       — 加权评分
    └── debug_layout()              — 布局追踪
```

---

## CSS 属性支持计划

| CSS 属性 | 优先级 | 实现位置 | 处理方式 |
|----------|--------|---------|---------|
| `grid-template-columns: repeat(N, 1fr)` | HIGH | Grid Layouter | 解析 fr 数量，均分可用宽度 |
| `grid-template-columns: 1fr 2fr` | HIGH | Grid Layouter | 解析比例，按 fr 权重分配 |
| `grid-template-columns: 200px 1fr` | MEDIUM | Grid Layouter | px 固定 + fr 剩余 |
| `justify-content: center` | HIGH | Flex Row/Column | 整体居中 |
| `justify-content: space-between` | MEDIUM | Flex Row/Column | 两端对齐 |
| `align-items: center` | HIGH | Flex Column | 容器内居中 |
| `box-sizing: border-box` | CRITICAL | 全局修正 | width 不再 + padding |
| `flex: 1` | MEDIUM | Flex Row | 弹性增长 |

---

## 实现顺序

1. **定义 IR schema** — 明确元素 dict 的字段和单位
2. **提取 Grid Layouter** — 独立函数，可单元测试
3. **提取 Flex Column Layouter** — 同上
4. **连线** — 在 `layout_slide_elements()` 中调用
5. **提取 Flex Row Layouter** — 同上
6. **清理 layout_slide_elements()** — 纯 Y 堆叠
7. **Golden Diff 加权评分** — 后验诊断

---

## 实现状态（2025-04-13）

### 已完成

| 修复项 | 状态 | 效果 |
|--------|------|------|
| Grid X 偏移修正 | ✅ | `container_abs_x - 0.50` → `container_abs_x`，列宽从 3.91" → 3.46"（golden 3.42"） |
| 卡片 padding 应用到 grid 单元格 | ✅ | `cell_has_bg` 检测 + `card_pad = 24/108` 背景单元格 |
| 溢出分页 | ✅ | `while content_bottom > safe_y` 自动拆分 |
| `clamp()` padding 解析修复 | ✅ | `_expand_padding` 使用 `re.findall` 尊重函数边界，`56px)` → `clamp(28px, 4vw, 56px)` |
| `internal_margin` 从 CSS 派生 | ✅ | `paddingTop` → `parse_px` → 英寸，替代硬编码 1.5" |
| 溢出页减少 | ✅ | 16 页 → 13 页（6 → 3 溢出） |

### 剩余问题

| 问题 | 影响 | 优先级 |
|------|------|--------|
| Y 偏移 -0.65" | 垂直居中逻辑与 golden 不一致 | MEDIUM |
| 列 gap 0.23" vs golden 0.34" | grid gap 计算偏差 | LOW |
| 元素提取覆盖率 | sandbox 52 text vs golden 176（架构差异） | HIGH |

---

## 预期效果

| 指标 | 当前 | 目标 |
|------|------|------|
| 溢出页面 | 3 页 | 0 页 |
| 文字覆盖率 | 100% ✓ | 100% |
| 位置偏差 > 0.5" | ~26 | < 20 |
| 位置偏差 > 0.2" | ~26 | < 80 |
| 评分（加权） | 待测 | 9.0+/10 |
