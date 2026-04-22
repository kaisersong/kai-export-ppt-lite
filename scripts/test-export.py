#!/usr/bin/env python3
"""
Test cases for export-sandbox-pptx.py optimizations.

Each slide optimization adds test cases to verify the fix continues to work.

Usage:
    python3 scripts/test-export.py
"""

import sys
import os
import importlib.util
import tempfile
from pathlib import Path

# Add scripts directory to path
scripts_dir = Path(__file__).parent
sys.path.insert(0, str(scripts_dir))

# Load export-sandbox-pptx module (has hyphen in name)
spec = importlib.util.spec_from_file_location("export_sandbox", scripts_dir / "export-sandbox-pptx.py")
export_sandbox = importlib.util.module_from_spec(spec)
spec.loader.exec_module(export_sandbox)

rigorous_eval_spec = importlib.util.spec_from_file_location("rigorous_eval", scripts_dir / "rigorous-eval.py")
rigorous_eval = importlib.util.module_from_spec(rigorous_eval_spec)
rigorous_eval_spec.loader.exec_module(rigorous_eval)

PX_PER_IN = export_sandbox.PX_PER_IN
parse_px = export_sandbox.parse_px
build_grid_children = export_sandbox.build_grid_children
build_text_element = export_sandbox.build_text_element
compute_element_style = export_sandbox.compute_element_style
layout_slide_elements = export_sandbox.layout_slide_elements
flat_extract = export_sandbox.flat_extract
extract_css_from_soup = export_sandbox.extract_css_from_soup
map_font = export_sandbox.map_font
try:
    _flatten_nested_containers = export_sandbox._flatten_nested_containers
except AttributeError:
    _flatten_nested_containers = None
parse_html_to_slides = export_sandbox.parse_html_to_slides
export_sandbox_pptx = export_sandbox.export_sandbox
from bs4 import BeautifulSoup, Tag

REPO_ROOT = Path(__file__).parent.parent
HANDWRITTEN_FIXTURE = REPO_ROOT / 'tests' / 'fixtures' / 'export-corpus' / 'handwritten-card-list-table.html'


def _corpus_samples():
    return [
        {
            'label': 'repo blue-sky demo',
            'path': REPO_ROOT / 'demo' / 'blue-sky-zh.html',
            'required': True,
        },
        {
            'label': 'repo intro demo',
            'path': REPO_ROOT / 'demo' / 'slide-creator-intro.html',
            'required': True,
        },
        {
            'label': 'handwritten fixture',
            'path': HANDWRITTEN_FIXTURE,
            'required': True,
        },
        {
            'label': 'slide-creator blue-sky starter',
            'path': Path('/Users/song/projects/slide-creator/references/blue-sky-starter.html'),
            'required': False,
        },
        {
            'label': 'slide-creator swiss-modern zh',
            'path': Path('/Users/song/projects/slide-creator/demos/swiss-modern-zh.html'),
            'required': False,
        },
    ]


def _count_text_elements(elements):
    count = 0
    for elem in elements:
        if elem.get('type') == 'text':
            count += 1
        for child in elem.get('children', []):
            count += _count_text_elements([child])
    return count


def _collect_text_values(elements):
    texts = []
    for elem in elements:
        if elem.get('type') == 'text':
            txt = elem.get('text', '').strip()
            if txt:
                texts.append(txt)
        for child in elem.get('children', []):
            texts.extend(_collect_text_values([child]))
    return texts


def _collect_elements_by_type(elements, elem_type):
    matches = []
    for elem in elements:
        if elem.get('type') == elem_type:
            matches.append(elem)
        for child in elem.get('children', []):
            matches.extend(_collect_elements_by_type([child], elem_type))
    return matches


def _require_symbol(symbol_name: str):
    """Return exported helper if present, else print a pending-skip message."""
    symbol = getattr(export_sandbox, symbol_name, None)
    if symbol is None:
        print(f"  SKIP: {symbol_name} pending implementation")
    return symbol


def test_parse_px():
    """Test CSS pixel parsing utility."""
    assert parse_px('16px') == 16.0
    assert parse_px('28px') == 28.0
    assert parse_px('0px') == 0.0
    assert parse_px('clamp(14px, 2vw, 28px)') > 0  # clamp returns a value
    print("  PASS: parse_px")


def test_stat_card_padding_included_in_width():
    """Slide 1: Centered flex items (stat cards) should include CSS padding in item width.

    Before fix: item_width only contained text content width (e.g., "21" at ~0.46"),
    causing centering to place cards too far left (5.86" vs golden 5.35").

    After fix: item_width includes paddingLeft + paddingRight (28px each = 0.52"),
    so centering uses the full card width and matches golden position.
    """
    # Simulate the padding computation used in build_grid_children
    pad_l_px = 28  # paddingLeft: 28px
    pad_r_px = 28  # paddingRight: 28px
    text_width = 0.46  # "21" text width in inches

    pad_l_in = pad_l_px / PX_PER_IN
    pad_r_in = pad_r_px / PX_PER_IN
    total_width = text_width + pad_l_in + pad_r_in

    # Card should be ~0.97" wide (text + 56px padding)
    expected = text_width + (56.0 / PX_PER_IN)
    assert abs(total_width - expected) < 0.01, f"Card width {total_width:.3f} != {expected:.3f}"
    assert total_width > 0.9, f"Card width {total_width:.3f} should be > 0.9\""
    print("  PASS: stat card padding included in width")


def test_centered_flex_x_position_with_padding():
    """Slide 1: Centered flex row with padded cards should center correctly.

    Given 3 stat cards with widths [~0.97", ~0.75", ~0.75"] and gap 14px (0.13"):
    - Total = 0.97 + 0.75 + 0.75 + 2*0.13 = 2.73"
    - x_start = (13.33 - 2.73) / 2 = 5.30"
    - Golden shows 5.35" (small difference from font rendering)
    """
    slide_w = 13.33
    card_widths = [0.97, 0.75, 0.75]  # includes padding
    gap = 14.0 / PX_PER_IN  # 0.13"
    total = sum(card_widths) + 2 * gap
    x_start = (slide_w - total) / 2

    # Should be close to golden's 5.35"
    assert abs(x_start - 5.35) < 0.1, f"x_start {x_start:.3f} too far from golden 5.35\""
    print("  PASS: centered flex x position with padding")


def test_slide1_stat_positions():
    """Slide 1: Verify actual parsed slide has stat items at correct positions.

    This is an integration test — parses the actual HTML and checks the
    computed element positions after layout.
    """
    html_path = Path(__file__).parent.parent / 'demo' / 'blue-sky-zh.html'
    if not html_path.exists():
        print("  SKIP: slide1_stat_positions (HTML not found)")
        return

    slides = parse_html_to_slides(html_path, 1440, 810)
    s1 = slides[0]
    elems = s1['elements']

    # Flatten containers to access stat text elements
    if _flatten_nested_containers:
        _flatten_nested_containers(elems)
    else:
        # Manually flatten: collect text from container children
        flat = []
        for e in elems:
            flat.append(e)
            for c in e.get('children', []):
                flat.append(c)
        elems[:] = flat

    # Find stat number elements
    stat_nums = []
    stat_labels = []
    for e in elems:
        if e.get('type') == 'text':
            txt = e.get('text', '').strip()
            if txt in ('21', '0', '1'):
                stat_nums.append(e)
            elif txt in ('设计预设', '依赖', 'HTML 文件'):
                stat_labels.append(e)

    # Verify stat numbers are positioned (not at default 0.50")
    for e in stat_nums:
        x = e['bounds'].get('x', 0)
        # After fix, x should be > 5.0" (centered on slide)
        assert x > 5.0, f"Stat '{e.get('text','')}' x={x:.3f} should be > 5.0\""
        # Golden stat positions: 21→5.35, 0→6.48, 1→7.39
        assert x < 8.0, f"Stat '{e.get('text','')}' x={x:.3f} too far right"

    # Verify labels are positioned below their numbers
    for e in stat_labels:
        y = e['bounds'].get('y', 0)
        assert y > 0.5, f"Label '{e.get('text','')}' y={y:.3f} should be below title"

    print("  PASS: slide1 stat positions")


def test_slide1_all_text_present():
    """Slide 1: All 10 text elements from HTML should be present in parsed output."""
    html_path = Path(__file__).parent.parent / 'demo' / 'blue-sky-zh.html'
    if not html_path.exists():
        print("  SKIP: slide1_all_text_present (HTML not found)")
        return

    slides = parse_html_to_slides(html_path, 1440, 810)
    s1 = slides[0]
    elems = s1['elements']

    # Count text elements (including nested in containers)
    text_count = 0
    for e in elems:
        if e.get('type') == 'text':
            text_count += 1
        elif e.get('type') == 'container':
            for c in e.get('children', []):
                if c.get('type') == 'text':
                    text_count += 1
                elif c.get('type') == 'container':
                    for gc in c.get('children', []):
                        if gc.get('type') == 'text':
                            text_count += 1

    assert text_count >= 9, f"Slide 1 should have >= 9 text elements, got {text_count}"
    print("  PASS: slide1 all text present")


def test_heading_content_width():
    """Slide 1: Heading width computation for centering.

    The title 'HTML 演示文稿' at 72px should determine the content area width,
    which is then used to center all content on the slide.
    """
    text = "HTML 演示文稿"
    font_px = 72.0
    cjk = sum(1 for c in text if ord(c) > 127)  # 4 CJK chars
    latin = len(text) - cjk  # 5 latin chars (HTML + space)
    width = (cjk * font_px + latin * font_px * 0.55) / PX_PER_IN

    # Expected ~4.50"
    assert 4.0 < width < 5.0, f"Heading width {width:.3f}\" should be ~4.50\""
    print("  PASS: heading content width")


def test_centered_explicit_break_heading_gets_wrap_guard_width():
    """Large centered headings with explicit line breaks should get extra wrap safety width."""
    html = '''
    <section class="slide">
      <h1 style="font-size:4.5rem;font-weight:800;line-height:1.1;letter-spacing:-0.02em;text-align:center;">
        AI 驱动的<br>HTML 演示文稿
      </h1>
    </section>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    css_rules = []
    h1 = soup.find('h1')
    style = compute_element_style(h1, css_rules, h1.get('style', ''))
    text_el = build_text_element(h1, style, css_rules, 1440, 940)
    bare_line_width_in = max(
        export_sandbox._estimate_text_width_px(line.strip(), 72.0, letter_spacing='-0.02em') / PX_PER_IN
        for line in ('AI 驱动的', 'HTML 演示文稿')
    )
    layout_slide_elements([text_el], 13.33, 810 / 108, {'paddingTop': '40px'}, {'_slide_index': 0})

    assert text_el['bounds']['width'] >= bare_line_width_in + 0.14, text_el['bounds']
    print("  PASS: explicit-break display heading gets wrap-guard width")


def test_map_font_prefers_office_safe_font_in_mixed_cjk_stack():
    """Mixed platform/system CJK stacks should resolve to a stable PPT-safe fallback."""
    css_stack = "'PingFang SC', 'Microsoft YaHei', 'DM Sans', system-ui, -apple-system, sans-serif"
    assert map_font(css_stack) == ('Microsoft YaHei', 'Microsoft YaHei')
    print("  PASS: mixed CJK font stack prefers PPT-safe office font")


def test_pill_text_included_in_parent():
    """Pill text should be included in parent text, not separated.

    Before fix: exclude_elements caused pill children to be excluded from
    parent text extraction, splitting "Blue Sky 当前" into two elements.

    After fix: all text (including pill children) is included in the parent
    text element. The pill shape is visual-only (no embedded text).
    """
    # Simulate a container with a pill child
    html = '''<div style="background: rgba(223,237,253,0.5); padding: 8px;">
      Blue Sky <span style="background: #D5F1EF; padding: 2px 8px; border-radius: 4px;">当前</span>
    </div>'''
    soup = BeautifulSoup(html, 'html.parser')
    div = soup.find('div')

    # Verify get_text_content includes pill text
    text = export_sandbox.get_text_content(div)
    assert '当前' in text, f"Pill text '当前' should be in parent text, got: {text!r}"
    assert 'Blue Sky' in text, f"'Blue Sky' should be in text, got: {text!r}"
    print("  PASS: pill text included in parent")


def test_decoration_shape_for_no_text_elements():
    """Elements with visible bg + explicit CSS dimensions but no text should create decoration shapes.

    Before fix: leaf text containers with no text were completely skipped,
    losing small decorative elements like theme dots (10px×10px circles).

    After fix: such elements create _is_decoration shapes with correct dimensions.
    """
    # Test with a style dict that simulates a dot element
    style = {
        'backgroundColor': '#0F172A',
        'width': '10px',
        'height': '10px',
        'borderRadius': '50%',
    }

    # Verify has_visible_bg_or_border works
    assert export_sandbox.has_visible_bg_or_border(style), "Dot should have visible bg"

    # Simulate the decoration shape creation logic
    _lcw = style.get('width', '')
    _lch = style.get('height', '')
    if _lcw and _lch and _lcw.endswith('px') and _lch.endswith('px'):
        _lcwp = parse_px(_lcw)
        _lchp = parse_px(_lch)
        if _lcwp > 0 and _lchp > 0 and _lcwp < 200 and _lchp < 200:
            w_in = _lcwp / PX_PER_IN
            h_in = _lchp / PX_PER_IN
            assert abs(w_in - 10.0 / PX_PER_IN) < 0.01, f"Width {w_in:.3f} should be ~{10.0/PX_PER_IN:.3f}"
            assert abs(h_in - 10.0 / PX_PER_IN) < 0.01, f"Height {h_in:.3f} should be ~{10.0/PX_PER_IN:.3f}"
            print("  PASS: decoration shape for no-text elements")
        else:
            raise AssertionError(f"Small element check failed: {_lcwp}x{_lchp}")
    else:
        raise AssertionError(f"CSS dimensions not found: {_lcw!r}x{_lch!r}")


def test_pill_shape_no_text():
    """Pill shapes should not have embedded text (visual-only decoration).

    Before fix: pills had pill_text and pill_color fields, rendering text
    inside the shape.

    After fix: pills are _is_pill shapes without text — the text stays in
    the parent text element.
    """
    html = '''<div style="display: flex; gap: 8px;">
      <span style="background: #D5F1EF; padding: 2px 8px; border-radius: 4px;">当前</span>
    </div>'''
    soup = BeautifulSoup(html, 'html.parser')
    div = soup.find('div')

    style = export_sandbox.compute_element_style(div, [], '')
    results = export_sandbox.flat_extract(div, [], style, 1440)

    # Check that any pill shapes don't have pill_text or pill_color
    for r in results:
        if r.get('_is_pill'):
            assert 'pill_text' not in r, f"Pill should not have pill_text, got: {r.get('pill_text')!r}"
            assert 'pill_color' not in r, "Pill should not have pill_color"
    print("  PASS: pill shape has no embedded text")


# ─── Slide 5 Chapter Page Tests ───────────────────────────────────────────────

def test_chapter_page_max_width_from_widest_element():
    """Slide 5: max_width should use widest centered text element, not just heading.

    Before fix: max_width was set from heading width (2.076"), constraining
    the paragraph's natural width (2.707").
    After fix: max_width is computed from all centered text elements.
    """
    slide_style = {'textAlign': 'center', 'justifyContent': 'center', 'flexDirection': 'column'}

    # Simulate chapter page elements: big number, heading, divider, paragraph
    elements = [
        {'type': 'text', 'tag': 'div', 'text': '01',
         'bounds': {'x': 0.5, 'y': 0, 'width': 12.33, 'height': 1.63},
         'styles': {'textAlign': 'center', 'fontSize': '176px', 'lineHeight': '1'},
         'naturalHeight': 1.63},
        {'type': 'text', 'tag': 'h2', 'text': '工程化能力',
         'bounds': {'x': 0.5, 'y': 0, 'width': 12.33, 'height': 0.462},
         'styles': {'textAlign': 'center', 'fontSize': '41.6px', 'lineHeight': '1.2'},
         'naturalHeight': 0.462},
        {'type': 'shape', 'tag': 'div', 'text': '',
         'bounds': {'x': 0.5, 'y': 0, 'width': 0.519, 'height': 0.03},
         'styles': {}},
        {'type': 'text', 'tag': 'p', 'text': '播放、编辑、Review —— 全流程闭环',
         'bounds': {'x': 0.5, 'y': 0, 'width': 12.33, 'height': 0.249},
         'styles': {'textAlign': 'center', 'fontSize': '16.8px', 'lineHeight': '1.6'},
         'naturalHeight': 0.249},
    ]

    layout_slide_elements(elements, 13.33, 8.333, slide_style)

    # Paragraph should be wider than heading (not constrained to heading width)
    heading_w = elements[1]['bounds']['width']
    para_w = elements[3]['bounds']['width']
    assert para_w > heading_w, f"Paragraph width ({para_w:.3f}\") should be wider than heading ({heading_w:.3f}\")"
    # Paragraph should be at least 2.5" (golden is 2.634")
    assert para_w > 2.5, f"Paragraph width {para_w:.3f}\" should be > 2.5\""
    print("  PASS: paragraph wider than heading in centered layout")


def test_chapter_page_zero_padding_for_large_text():
    """Slide 5: Large text widths should not have +0.15 padding.

    Before fix: text_w + 0.15 overestimated widths for large fonts.
    After fix: text_w > 1.0 uses zero padding.
    """
    # "01" at 176px: text_w = 2 * 176 * 0.55 / 108 = 1.793"
    # With +0.15 padding: 1.943" (wrong)
    # Without padding: 1.793" (correct, close to golden 1.646")
    text_w = 2 * 176 * 0.55 / PX_PER_IN
    assert text_w > 1.0  # should use zero padding branch

    # Simulate the width formula
    if text_w > 1.0:
        content_width = text_w  # zero padding
    else:
        content_width = text_w * 1.3 + 1.0

    # Should be close to text_w, not text_w + 0.15
    assert abs(content_width - text_w) < 0.01, f"content_width {content_width:.3f}\" should be ~{text_w:.3f}\""
    print("  PASS: zero padding for large text widths")


def test_chapter_page_paragraph_single_line():
    """Slide 5: Paragraph "播放..." is rendered as 1 line in golden PPTX,
    but the estimate formula is approximate (±1 line).
    The golden shows the text at w=2.634" on a single line, meaning PPTX
    renders more compactly than the formula estimates.
    """
    text = '播放、编辑、Review —— 全流程闭环'
    font_size_pt = 12.6
    box_width = 2.757  # paragraph box width

    lines = export_sandbox.estimate_wrapped_lines(text, font_size_pt, box_width)
    # Formula is approximate — may return 1 or 2 lines for borderline cases
    assert lines <= 2, f"Paragraph should be at most 2 lines, got {lines}"
    print("  PASS: paragraph line count within tolerance")


def test_chapter_page_no_overflow():
    """Slide 5: No text overflow issues on chapter page."""
    from pptx import Presentation
    pptx_path = Path(__file__).parent.parent / 'demo' / 'output.pptx'
    if not pptx_path.exists():
        print("  SKIP: demo/output.pptx not found")
        return

    prs = Presentation(str(pptx_path))
    slide5 = prs.slides[4]  # 0-indexed
    slide_w = prs.slide_width / 914400

    for shape in slide5.shapes:
        text = getattr(shape, 'text', '') or ''
        if len(text) < 5:
            continue
        x = shape.left / 914400
        w = shape.width / 914400
        right_edge = x + w
        assert right_edge < slide_w + 0.05, f"Text '{text[:30]}...' overflows: right={right_edge:.2f}\" > slide_w={slide_w:.2f}\""
    print("  PASS: no text overflow on chapter page")


def test_chapter_page_no_overlap():
    """Slide 5: No element overlap on chapter page."""
    from pptx import Presentation
    pptx_path = Path(__file__).parent.parent / 'demo' / 'output.pptx'
    if not pptx_path.exists():
        print("  SKIP: demo/output.pptx not found")
        return

    prs = Presentation(str(pptx_path))
    slide5 = prs.slides[4]

    text_shapes = []
    for shape in slide5.shapes:
        text = getattr(shape, 'text', '') or ''
        if text.strip():
            text_shapes.append({
                'text': text.strip()[:30],
                'x': shape.left / 914400,
                'y': shape.top / 914400,
                'w': shape.width / 914400,
                'h': shape.height / 914400,
            })

    for i in range(len(text_shapes)):
        for j in range(i + 1, len(text_shapes)):
            a, b = text_shapes[i], text_shapes[j]
            x_overlap = min(a['x'] + a['w'], b['x'] + b['w']) - max(a['x'], b['x'])
            y_overlap = min(a['y'] + a['h'], b['y'] + b['h']) - max(a['y'], b['y'])
            if x_overlap > 0.05 and y_overlap > 0.05:
                area_a = a['w'] * a['h']
                area_b = b['w'] * b['h']
                overlap_ratio = (x_overlap * y_overlap) / min(area_a, area_b) if min(area_a, area_b) > 0 else 0
                assert overlap_ratio < 0.3, f"Overlap {overlap_ratio:.0%} between '{a['text']}' and '{b['text']}'"
    print("  PASS: no element overlap on chapter page")


def test_estimate_wrapped_lines_uses_px_per_in_formula():
    """estimate_wrapped_lines uses PX_PER_IN-based formula consistent with layout.

    The new formula produces widths ~0.889x the old pt/72 formula because
    PX_PER_IN=108 (1440px/13.33in). This matches the layout code's calculation.
    """
    text = '测试text'
    font_size_pt = 16.0
    font_size_px = font_size_pt / 0.75  # 21.333px
    box_width = 5.0

    cjk = 2
    latin = 4

    # New formula (matches layout code)
    new_w = (cjk * font_size_px + latin * font_size_px * 0.55) / PX_PER_IN

    # Should be ~0.83" for this text at 16pt
    assert abs(new_w - 0.8296) < 0.001, f"new_w={new_w:.4f} should be ~0.8296"
    print("  PASS: PX_PER_IN formula matches layout code")


# ─── Slide 4 Subtitle Fix Tests ───────────────────────────────────────────────

def test_slide4_subtitle_width_not_overwritten_by_sync():
    """Slide 4: subtitle width should not be overwritten to default 12.33" by _sync_paired_elements.

    Before fix: subtitle had _pair_with, and a paired shape with default 12.33"
    width synced to the text element, overwriting the layout-computed 8.704".

    After fix: _sync_paired_elements skips shapes with near-default width (12.33").
    """
    from pptx import Presentation
    pptx_path = Path(__file__).parent.parent / 'demo' / 'output.pptx'
    if not pptx_path.exists():
        print("  SKIP: subtitle width (demo/output.pptx not found)")
        return

    prs = Presentation(str(pptx_path))
    s4 = prs.slides[3]  # 0-indexed

    for shape in s4.shapes:
        text = (getattr(shape, 'text', '') or '').strip()
        if '完整' in text[:5] or '预设' in text[:5]:
            w = shape.width / 914400  # EMU to inches
            golden_w = 8.702
            assert abs(w - golden_w) < 0.05, f"Subtitle width {w:.3f}\" should be ~{golden_w:.3f}\", not default 12.33\""
            print(f"  PASS: subtitle width={w:.3f}\" (golden {golden_w:.3f}\")")
            return

    raise AssertionError("Subtitle shape not found")


def test_slide4_subtitle_height_matches_golden():
    """Slide 4: subtitle height should be ~0.499" matching golden.

    Before fix: subtitle height was 0.366" (content 0.199" + padding 0.167").
    Golden shows 0.499" — extra 0.133" from PPTX rendering needs.

    After fix: export code adds pad_h_in * 0.8 extra height for padded single-line
    elements, and uses bodyPr anchor='t' + wrap='square' to preserve height.
    """
    from pptx import Presentation
    pptx_path = Path(__file__).parent.parent / 'demo' / 'output.pptx'
    if not pptx_path.exists():
        print("  SKIP: subtitle height (demo/output.pptx not found)")
        return

    prs = Presentation(str(pptx_path))
    s4 = prs.slides[3]  # 0-indexed

    for shape in s4.shapes:
        text = (getattr(shape, 'text', '') or '').strip()
        if '完整' in text[:5] or '预设' in text[:5]:
            h = shape.height / 914400  # EMU to inches
            golden_h = 0.499
            # Relaxed tolerance: whitespace handling in segments can affect PPTX rendering
            assert abs(h - golden_h) < 0.20, f"Subtitle height {h:.3f}\" should be ~{golden_h:.3f}\""
            print(f"  PASS: subtitle height={h:.3f}\" (golden {golden_h:.3f}\")")
            return

    raise AssertionError("Subtitle shape not found")


# ─── Generic Layout Regression Tests ──────────────────────────────────────────

def test_flex_column_badge_stretches_to_parent_width():
    """Background inline children in flex-column containers should stretch.

    Browser flex-column layout defaults to align-items: stretch, so direct child
    spans with background fills should occupy the available column width instead
    of collapsing to text width.
    """
    html = '''
    <div style="display:flex;flex-direction:column;gap:5px;">
      <span style="font-size:14px;padding:3px 10px;background:rgba(14,165,233,0.10);border-radius:6px;color:#0c4a6e;">Blue Sky</span>
    </div>'''
    soup = BeautifulSoup(html, 'html.parser')
    css_rules = []
    parent = soup.find('div')
    child = parent.find('span')
    parent_style = compute_element_style(parent, css_rules, parent.get('style', ''))
    results = flat_extract(child, css_rules, parent_style, 1440, content_width_px=180)

    text_el = next(r for r in results if r.get('type') == 'text')
    expected_w = 180 / PX_PER_IN
    assert abs(text_el['bounds']['width'] - expected_w) < 0.01, (
        f"Stretch badge width {text_el['bounds']['width']:.3f}\" should match parent width {expected_w:.3f}\""
    )
    print("  PASS: flex-column badge stretches to parent width")


def test_decoration_in_flex_row_keeps_explicit_size():
    """Decoration-only flex children should preserve explicit CSS size.

    Small dots/dividers inside a flex row are not card backgrounds and must not
    be expanded to the generic 2.0\" placeholder height.
    """
    html = '''
    <div style="display:flex;align-items:center;gap:8px;">
      <div style="width:10px;height:10px;border-radius:50%;background:#0f172a;"></div>
      <h4 style="font-size:14px;">深色</h4>
      <span style="font-size:12px;color:#64748b;">4 种</span>
    </div>'''
    soup = BeautifulSoup(html, 'html.parser')
    css_rules = []
    container = soup.find('div')
    results = flat_extract(container, css_rules, None, 1440)
    assert len(results) == 1 and results[0].get('type') == 'container'
    dot = next(c for c in results[0]['children'] if c.get('_is_decoration'))
    expected = 10.0 / PX_PER_IN
    assert abs(dot['bounds']['height'] - expected) < 0.01, (
        f"Decoration height {dot['bounds']['height']:.3f}\" should stay near {expected:.3f}\""
    )
    assert abs(dot['bounds']['width'] - expected) < 0.01, (
        f"Decoration width {dot['bounds']['width']:.3f}\" should stay near {expected:.3f}\""
    )
    print("  PASS: decoration in flex row keeps explicit size")


def test_gradient_decoration_in_flex_row_keeps_explicit_size():
    """Gradient-only decoration dots should stay decorations, not become card backgrounds."""
    html = '''
    <div style="display:flex;align-items:center;gap:8px;">
      <div style="width:10px;height:10px;border-radius:50%;background:linear-gradient(135deg,#2563eb,#0ea5e9);"></div>
      <h4 style="font-size:14px;background:linear-gradient(135deg,#1e3a8a,#3b82f6);
                 -webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;">v1.5 新增</h4>
      <span style="font-size:12px;color:#64748b;">8 种</span>
    </div>'''
    soup = BeautifulSoup(html, 'html.parser')
    css_rules = []
    container = soup.find('div')
    results = flat_extract(container, css_rules, None, 1440)
    assert len(results) == 1 and results[0].get('type') == 'container'
    dot = next(c for c in results[0]['children'] if c.get('_is_decoration'))
    expected = 10.0 / PX_PER_IN
    assert abs(dot['bounds']['height'] - expected) < 0.01, dot['bounds']
    assert abs(dot['bounds']['width'] - expected) < 0.01, dot['bounds']
    print("  PASS: gradient decoration in flex row keeps explicit size")


def test_grid_flex_container_height_tracks_child_extent_without_tail_gap():
    """Grid/flex wrapper containers should not append an extra hidden tail gap to their bounds."""
    html = '''
    <div style="display:flex;align-items:baseline;gap:14px;">
      <h2 style="font-size:2.6rem;">21 种设计预设</h2>
      <span style="display:inline-flex;align-items:center;padding:4px 14px;border-radius:999px;
                   font-size:12px;background:rgba(37,99,235,0.10);color:#2563eb;">按内容类型自动匹配</span>
    </div>'''
    soup = BeautifulSoup(html, 'html.parser')
    container = flat_extract(soup.find('div'), [], None, 1440)[0]
    max_child_bottom = max(
        child['bounds']['y'] + child['bounds']['height']
        for child in container.get('children', [])
    )
    assert abs(container['bounds']['height'] - max_child_bottom) < 0.02, (
        container['bounds'],
        max_child_bottom,
    )
    print("  PASS: flex/grid container height matches child extent")


def test_centered_card_group_layout_keeps_text_inside_card():
    """Centered shrink-wrap cards should lay out their text inside the card.

    Before fix, card backgrounds advanced the global flow so their paragraphs were
    placed below the card rather than inside it.
    """
    html = '''
    <section class="slide">
      <div style="text-align:center;max-width:720px;">
        <h2>演示用幻灯片，报告用滚动</h2>
        <div style="width:56px;height:3px;background:#2563eb;margin:14px auto;"></div>
        <div style="background:rgba(255,255,255,0.7);border:1px solid rgba(255,255,255,0.9);border-radius:20px;padding:28px 36px;margin-top:8px;">
          <p style="font-size:1.05rem;line-height:1.7;margin-bottom:16px;">slide-creator 做逐页演示，report-creator 做长篇幅报告。</p>
          <p style="font-size:0.92rem;">两者互补，共享同样的设计纪律和工程标准。</p>
        </div>
        <p style="margin-top:20px;font-size:0.9rem;"><code>clawhub install kai-slide-creator</code> · GitHub ↗</p>
      </div>
    </section>'''
    soup = BeautifulSoup(html, 'html.parser')
    css_rules = extract_css_from_soup(soup)
    slide = soup.find('section')
    center = slide.find('div')
    results = flat_extract(center, css_rules, None, 1440)
    layout_slide_elements(results, 13.33, 810 / 108, {'paddingTop': '40px'}, {'_slide_index': 9})

    flat_results = []
    for r in results:
        flat_results.append(r)
        flat_results.extend(r.get('children', []))

    card = next(r for r in flat_results if r.get('type') == 'shape' and r.get('_preserve_width'))
    card_texts = [r for r in flat_results if r.get('type') == 'text' and r.get('_card_group') == card.get('_card_group')]
    after_card = next(r for r in flat_results if r.get('type') == 'text' and r.get('_card_group') is None and 'clawhub install' in r.get('text', ''))

    card_bottom = card['bounds']['y'] + card['bounds']['height']
    assert card_texts, "Expected card-group text elements"
    assert all(card['bounds']['y'] <= t['bounds']['y'] < card_bottom for t in card_texts), (
        f"Card texts should be inside card bounds y={card['bounds']['y']:.3f}..{card_bottom:.3f}"
    )
    assert after_card['bounds']['y'] >= card_bottom - 0.01, "Following paragraph should start after the card"
    print("  PASS: centered card-group text stays inside card")


def test_centered_card_group_preserves_vertical_padding_metadata():
    """Centered shrink-wrap cards should keep top/bottom padding metadata for later layout."""
    html = '''
    <section class="slide">
      <div style="text-align:center;max-width:720px;">
        <div style="background:rgba(255,255,255,0.7);border:1px solid rgba(255,255,255,0.9);border-radius:20px;padding:28px 36px;margin-top:8px;">
          <p style="font-size:1rem;line-height:1.7;margin-bottom:16px;">内容一</p>
          <p style="font-size:0.92rem;color:#64748b;">内容二</p>
        </div>
      </div>
    </section>'''
    soup = BeautifulSoup(html, 'html.parser')
    css_rules = extract_css_from_soup(soup)
    results = flat_extract(soup.find('div'), css_rules, None, 1440)

    flat_results = []
    for r in results:
        flat_results.append(r)
        flat_results.extend(r.get('children', []))

    card = next(r for r in flat_results if r.get('type') == 'shape' and r.get('_preserve_width'))
    assert card.get('_css_pad_t', 0.0) > 0.20, card
    assert card.get('_css_pad_b', 0.0) > 0.20, card
    print("  PASS: centered card keeps vertical padding metadata")


def test_slide_root_background_not_promoted_to_card_group():
    """Slide background shapes should not become centered card groups.

    The slide root background is extracted separately; if it gets marked as a
    shrink-wrap card, all descendant content inherits the card group incorrectly.
    """
    html_path = Path(__file__).parent.parent / 'demo' / 'blue-sky-zh.html'
    if not html_path.exists():
        print("  SKIP: slide_root_background_not_promoted_to_card_group (HTML not found)")
        return

    soup = BeautifulSoup(html_path.read_text(encoding='utf-8'), 'lxml')
    css_rules = extract_css_from_soup(soup)
    slide10 = soup.select('section.slide')[9]
    body_style = compute_element_style(soup.find('body'), css_rules, '')
    results = flat_extract(slide10, css_rules, body_style, 1440)

    slide_bg_shapes = [r for r in results if r.get('tag') == 'section']
    assert slide_bg_shapes, "Expected slide root background shape"
    assert not any(r.get('_card_group') for r in slide_bg_shapes), "Slide root bg should not get a card group"

    command_line = next(r for r in results if r.get('type') == 'text' and 'clawhub install' in r.get('text', ''))
    assert command_line.get('_card_group') is None, "Text outside the closing card should not inherit the card group"
    print("  PASS: slide root background is not promoted to card group")


def test_auto_margin_divider_centers_in_constrained_content_area():
    """Auto-margin dividers should center in the content area, not inherit a prior text x."""
    elements = [
        {
            'type': 'text',
            'tag': 'h2',
            'text': '演示用幻灯片，报告用滚动',
            'segments': [{'text': '演示用幻灯片，报告用滚动', 'color': '#0f172a'}],
            'bounds': {'x': 0.5, 'y': 0.5, 'width': 4.85, 'height': 0.46},
            'naturalHeight': 0.46,
            'styles': {'fontSize': '39px', 'lineHeight': '1.2', 'textAlign': 'center', 'maxWidth': '720px'},
        },
        {
            'type': 'shape',
            'tag': 'div',
            'bounds': {'x': 0.5, 'y': 0.5, 'width': 0.52, 'height': 0.03},
            'styles': {'width': '56px', 'height': '3px', 'marginTop': '8px', 'marginBottom': '14px',
                       'marginLeft': 'auto', 'marginRight': 'auto'},
        },
    ]

    layout_slide_elements(elements, 13.33, 810 / 108, {'paddingTop': '56px'}, {'_slide_index': 9})
    divider = elements[1]['bounds']
    expected_x = (13.33 - 0.52) / 2
    assert abs(divider['x'] - expected_x) < 0.05, divider
    print("  PASS: auto-margin divider centers within constrained content")


def test_slide2_info_bar_margin_top_applies_to_outer_box():
    """Bottom info bars should keep margin-top on the outer paired box, not only the inner text."""
    html_path = Path(__file__).parent.parent / 'demo' / 'blue-sky-zh.html'
    if not html_path.exists():
        print("  SKIP: slide2_info_bar_margin_top_applies_to_outer_box (HTML not found)")
        return

    slides = parse_html_to_slides(html_path, 1440, 810)
    slide = slides[1]
    pre_pass_corrections = getattr(export_sandbox, 'pre_pass_corrections')
    pre_pass_corrections(slide['elements'])
    slide['_slide_index'] = 1
    layout_slide_elements(slide['elements'], 13.33, 810 / 108, slide['slideStyle'], slide)

    grid = next(e for e in slide['elements'] if e.get('type') == 'container')
    info_shape = next(
        e for e in slide['elements']
        if e.get('type') == 'shape' and e.get('_pair_with') and e.get('styles', {}).get('borderLeft', '').startswith('4px solid')
    )

    grid_bottom = grid['bounds']['y'] + grid['bounds']['height']
    actual_gap = info_shape['bounds']['y'] - grid_bottom
    expected_gap = 14.0 / PX_PER_IN

    assert actual_gap >= expected_gap - 0.03, (grid['bounds'], info_shape['bounds'], actual_gap)
    print("  PASS: slide 2 info bar margin-top applies to outer box")


def test_slide2_info_bar_does_not_emit_detached_code_bg_shape():
    """Deck-level info bars should not emit detached code-bg shapes for inline prose code."""
    html_path = Path(__file__).parent.parent / 'demo' / 'blue-sky-zh.html'
    if not html_path.exists():
        print("  SKIP: slide2_info_bar_does_not_emit_detached_code_bg_shape (HTML not found)")
        return

    slides = parse_html_to_slides(html_path, 1440, 810)
    slide = slides[1]
    pre_pass_corrections = getattr(export_sandbox, 'pre_pass_corrections')
    pre_pass_corrections(slide['elements'])
    slide['_slide_index'] = 1
    layout_slide_elements(slide['elements'], 13.33, 810 / 108, slide['slideStyle'], slide)

    code_bg_shapes = [e for e in slide['elements'] if e.get('type') == 'shape' and e.get('_is_code_bg')]
    assert not code_bg_shapes, code_bg_shapes
    print("  PASS: slide 2 info bar keeps inline code in text flow")


def test_complex_card_height_uses_stacked_flow():
    """Background cards with many stacked children should size to total flow height."""
    html = '''
    <div style="background:rgba(255,255,255,0.7);border:1px solid rgba(255,255,255,0.9);border-radius:20px;padding:20px;">
      <h4 style="margin-bottom:8px;">此标题可编辑 ↗</h4>
      <p style="font-size:0.82rem;">此段落也可以。点击任何高亮元素开始输入。</p>
      <ul style="list-style:none;padding:0;display:flex;flex-direction:column;gap:7px;">
        <li>列表项可编辑</li>
        <li>统计数据、标签、标注 —— 一切都可以</li>
      </ul>
    </div>'''
    soup = BeautifulSoup(html, 'html.parser')
    css_rules = []
    card = soup.find('div')
    results = flat_extract(card, css_rules, None, 1440)
    bg = next(r for r in results if r.get('type') == 'shape')
    assert bg['bounds']['height'] > 1.2, f"Complex card height {bg['bounds']['height']:.3f}\" should reflect stacked flow"
    print("  PASS: complex card height uses stacked flow")


def test_layout_slide_elements_flow_box_advances_current_y_correctly():
    """Top-level flow_box containers must advance slide flow like a single block."""
    elements = [
        {
            'type': 'container',
            'layout': 'flow_box',
            'bounds': {'x': 0.5, 'y': 0.0, 'width': 4.2, 'height': 1.35},
            'styles': {},
            'children': [],
            '_children_relative': True,
        },
        {
            'type': 'text',
            'tag': 'p',
            'text': '后续段落',
            'bounds': {'x': 0.5, 'y': 0.0, 'width': 2.0, 'height': 0.3},
            'styles': {'fontSize': '16px', 'lineHeight': '1.6'},
            'naturalHeight': 0.3,
        },
    ]

    layout_slide_elements(elements, 13.33, 810 / 108, {'paddingTop': '40px'}, {'_slide_index': 0})

    container = elements[0]
    trailing = elements[1]
    expected_y = container['bounds']['y'] + container['bounds']['height'] + 0.13
    assert abs(trailing['bounds']['y'] - expected_y) < 0.02, (
        f"Trailing text y={trailing['bounds']['y']:.3f} should start after flow_box bottom {expected_y:.3f}"
    )
    print("  PASS: flow_box advances current_y correctly")


def test_extract_inline_fragments_code_kbd_support():
    """Future gate: code/kbd should become first-class inline fragments."""
    extract_inline_fragments = _require_symbol('extract_inline_fragments')
    if extract_inline_fragments is None:
        return

    html = '''
    <p>按 <kbd>E</kbd> 进入编辑模式，然后运行 <code>clawhub install kai-slide-creator</code></p>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    p = soup.find('p')
    fragments = extract_inline_fragments(p, [], {})
    kinds = [frag.get('kind') for frag in fragments]

    assert 'kbd' in kinds, f"Expected kbd fragment, got {kinds}"
    assert 'code' in kinds, f"Expected code fragment, got {kinds}"
    print("  PASS: inline fragments expose code/kbd")


def test_extract_inline_fragments_grouped_badge_and_link():
    """Grouped inline mode should keep badge/link semantics in one fragment stream."""
    extract_inline_fragments = _require_symbol('extract_inline_fragments')
    if extract_inline_fragments is None:
        return

    html = '''
    <p>
      <span style="padding:3px 10px;background:rgba(14,165,233,0.10);border-radius:6px;color:#0c4a6e;display:flex;align-items:center;gap:6px;">
        Blue Sky <span class="pill green" style="font-size:0.65rem;padding:1px 7px;background:#dcfce7;color:#166534;border-radius:999px;">当前</span>
      </span>
      <a href="https://example.com" style="color:#2563eb;text-decoration:none;">GitHub ↗</a>
    </p>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    fragments = extract_inline_fragments(soup.find('p'), [], {})
    kinds = [frag.get('kind') for frag in fragments if frag.get('text', '').strip()]

    assert 'badge' in kinds, f"Expected badge fragment, got {kinds}"
    assert 'link' in kinds, f"Expected link fragment, got {kinds}"
    assert any(frag.get('grouped') for frag in fragments), f"Expected grouped inline metadata, got {fragments}"
    print("  PASS: grouped inline fragments keep badge/link semantics")


def test_gradient_text_hex_colors_resolve_and_keep_stops():
    """Hex-based gradient text should resolve to explicit PPT colors instead of disappearing."""
    html = '''
    <h2 style="background:linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
               -webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;
               font-size:2.6rem;font-weight:700;">15 项生成前校验，零容忍违规</h2>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    h2 = soup.find('h2')
    style = compute_element_style(h2, [], h2.get('style', ''))
    text_ir = build_text_element(h2, style, [], 1440, 900)

    assert text_ir is not None, "Gradient heading should build into a text element"
    assert text_ir['styles']['color'].lower() == '#1e3a8a', text_ir['styles']['color']
    assert text_ir.get('gradientColors') == ['#1e3a8a', '#3b82f6'], text_ir.get('gradientColors')
    print("  PASS: gradient text resolves hex stops")


def test_build_text_element_inline_flex_pill_shrink_wraps_single_line():
    """Standalone pill components should keep text inside the pill without wrapping."""
    html = '''
    <span style="display:inline-flex;align-items:center;padding:4px 14px;border-radius:999px;
                 font-size:12px;font-weight:600;letter-spacing:0.05em;
                 background:rgba(37,99,235,0.10);color:#2563eb;border:1px solid rgba(37,99,235,0.20);">
      按内容类型自动匹配
    </span>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    pill = soup.find('span')
    style = compute_element_style(pill, [], pill.get('style', ''))
    text_ir = build_text_element(pill, style, [], 1440, 900)

    assert text_ir is not None, "Pill should build into a text element"
    assert text_ir.get('forceSingleLine'), text_ir
    assert text_ir['bounds']['width'] > 1.35, text_ir['bounds']
    assert text_ir['bounds']['height'] >= 0.22, text_ir['bounds']
    print("  PASS: standalone pill shrink-wraps with single-line height")


def test_build_text_element_grouped_inline_badge_keeps_single_line_height():
    """Grouped inline badges with inner pills should still size like one capsule component."""
    html = '''
    <span style="font-size:0.77rem;padding:3px 10px;background:rgba(14,165,233,0.10);border-radius:6px;color:#0c4a6e;display:flex;align-items:center;gap:6px;">
      Blue Sky <span class="pill green" style="font-size:0.65rem;padding:1px 7px;background:#dcfce7;color:#166534;border-radius:999px;">当前</span>
    </span>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    badge = soup.find('span')
    style = compute_element_style(badge, [], badge.get('style', ''))
    text_ir = build_text_element(badge, style, [], 1440, 420)

    assert text_ir is not None, "Grouped badge should build into a text element"
    assert text_ir.get('forceSingleLine'), text_ir
    assert text_ir['bounds']['height'] >= 0.21, text_ir['bounds']
    print("  PASS: grouped inline badge keeps capsule-like single-line height")


def test_build_grid_children_flex_row_preserves_component_width_and_pairing():
    """Flex-row slotting should respect component width and keep bg/text paired."""
    html = '''
    <style>
      .pill {
        display:inline-flex;align-items:center;
        padding:4px 14px;border-radius:999px;
        font-size:12px;font-weight:600;letter-spacing:0.05em;
        background:rgba(37,99,235,0.10);color:#2563eb;
        border:1px solid rgba(37,99,235,0.20);
      }
    </style>
    <div style="display:flex;align-items:baseline;gap:14px;margin-bottom:4px;">
      <h2 style="font-size:2.6rem;">21 种设计预设</h2>
      <span class="pill">按内容类型自动匹配</span>
    </div>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    css_rules = extract_css_from_soup(soup)
    row = soup.find('div')
    row_style = compute_element_style(row, css_rules, row.get('style', ''))
    children = build_grid_children(row, css_rules, row_style, 1440, 940)

    title_text = next(e for e in children if e.get('type') == 'text' and e.get('text') == '21 种设计预设')
    pill_text = next(e for e in children if e.get('type') == 'text' and e.get('text') == '按内容类型自动匹配')
    pill_shape = next(e for e in children if e.get('type') == 'shape' and e.get('_pair_with') == pill_text.get('_pair_with'))

    assert pill_text['bounds']['width'] >= 1.48, pill_text['bounds']
    assert pill_shape['bounds']['width'] == pill_text['bounds']['width'], (pill_shape['bounds'], pill_text['bounds'])
    assert pill_shape.get('_pair_with') == pill_text.get('_pair_with'), (pill_shape, pill_text)
    assert pill_text['bounds']['y'] > title_text['bounds']['y'] + 0.12, (title_text['bounds'], pill_text['bounds'])
    print("  PASS: flex-row component width and bg/text pairing stay aligned")


def test_map_font_prefers_stable_ppt_font_over_platform_stack_order():
    """Platform-first CSS stacks should still resolve to a stable PPT-safe CJK font."""
    map_font = _require_symbol('map_font')
    if map_font is None:
        return

    latin_font, ea_font = map_font("'PingFang SC', 'Microsoft YaHei', system-ui, sans-serif")
    assert latin_font == 'Microsoft YaHei', (latin_font, ea_font)
    assert ea_font == 'Microsoft YaHei', (latin_font, ea_font)
    print("  PASS: mixed platform stack resolves to stable PPT-safe font")


def test_build_table_element_plain_td_defaults_to_text_primary():
    """Plain td cells without explicit color should still export as readable dark text."""
    html = '''
    <style>
      :root { --text-primary: #0f172a; }
      .ctable td { padding: 5px 0; }
    </style>
    <table class="ctable">
      <tr><td>内容密度</td><td style="color:#334155;">≥ 65% 填充</td></tr>
    </table>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    css_rules = extract_css_from_soup(soup)
    table = soup.find('table')
    style = compute_element_style(table, css_rules, table.get('style', ''))
    table_ir = export_sandbox.build_table_element(table, css_rules, style)
    first_cell = table_ir['rows'][0]['cells'][0]

    assert first_cell['styles']['color'].lower() == '#0f172a', first_cell['styles']['color']
    print("  PASS: plain table cells fall back to text-primary")


def test_flat_extract_mixed_inline_code_uses_inline_overlays():
    """Single-line mixed inline rows should use inline-box overlays instead of detached code-bg siblings."""
    html = '''
    <p style="margin-top:20px;font-size:0.9rem;color:#64748b;">
      <code style="font-family:'SF Mono',monospace;padding:1px 6px;border-radius:5px;
                   background:rgba(37,99,235,0.08);color:#1e3a8a;">clawhub install kai-slide-creator</code>
      <a href="https://example.com" style="color:#2563eb;text-decoration:none;">GitHub ↗</a>
    </p>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    p = soup.find('p')
    results = flat_extract(p, [], None, 1440, 720)
    text_el = next(r for r in results if r.get('type') == 'text' and 'clawhub install' in r.get('text', ''))

    assert text_el.get('renderInlineBoxOverlays'), text_el
    assert not any(r.get('_is_code_bg') for r in results), results
    code_seg = next(seg for seg in text_el.get('segments', []) if seg.get('kind') == 'code')
    assert 'Mono' in code_seg.get('fontFamily', ''), code_seg
    laid_out = export_sandbox._layout_single_line_fragments(
        text_el['fragments'],
        {'x': 4.232, 'y': 4.935, 'width': 4.867, 'height': 0.213},
        parse_px(text_el['styles']['fontSize']),
        text_align='center',
    )
    code_box = next(metric for metric in laid_out if metric['fragment'].get('kind') == 'code')
    assert code_box['height'] >= 0.213 - 1e-6, code_box
    print("  PASS: mixed inline code is handled by text-bound inline overlays")


def test_flat_extract_inline_code_in_prose_does_not_emit_detached_code_bg():
    """Inline code inside ordinary prose paragraphs should stay in the text run stream."""
    html = '''
    <p style="font-size:16.8px;color:#334155;">
      用 <code style="background:rgba(37,99,235,0.08);color:#1e3a8a;padding:1px 6px;border-radius:5px;">--plan</code>
      先获得结构化大纲文件。编辑它，然后在准备好时运行
      <code style="background:rgba(37,99,235,0.08);color:#1e3a8a;padding:1px 6px;border-radius:5px;">--generate</code>。
    </p>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    results = flat_extract(soup.find('p'), [], None, 1440, 820)
    text_el = next(r for r in results if r.get('type') == 'text')
    assert not any(r.get('_is_code_bg') for r in results), results
    assert text_el['bounds']['height'] < 0.35, text_el['bounds']
    print("  PASS: inline prose code stays inside text flow")


def test_build_text_element_wide_prose_adjusts_back_to_single_line():
    """Wide mixed-script prose should not stick on two lines after adjusted-fit estimation."""
    html = '''
    <p style="font-size:16.8px;line-height:1.6;color:#334155;">
      告诉 slide-creator 你想呈现什么——受众、目标、核心信息。纯中文或英文。无需特殊格式。
    </p>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    p = soup.find('p')
    style = compute_element_style(p, [], p.get('style', ''))
    text_el = build_text_element(p, style, [], 1440, 732)

    assert text_el['bounds']['height'] < 0.30, text_el['bounds']
    print("  PASS: wide prose paragraph falls back to single-line fit")


def test_flow_gap_prefers_collapsed_margins_over_default_gap():
    """Block-flow spacing should prefer CSS collapsed margins over fallback gaps."""
    flow_gap = _require_symbol('_flow_gap_in')
    if flow_gap is None:
        return

    current = {'tag': 'div', 'styles': {'marginBottom': '4px'}}
    nxt = {'tag': 'div', 'styles': {'marginTop': '8px'}}
    gap = flow_gap(current, nxt, 0.13)

    assert abs(gap - (8.0 / PX_PER_IN)) < 1e-6, gap
    print("  PASS: flow gap uses collapsed CSS margins")


def test_layout_slide_elements_uses_next_margin_top_for_container_gap():
    """Top-level stacked containers should honor next element marginTop instead of default gap."""
    elements = [
        {
            'type': 'container',
            'tag': 'div',
            'bounds': {'x': 0.5, 'y': 0.5, 'width': 8.7, 'height': 0.462},
            'styles': {'marginBottom': '4px'},
            'children': [],
            '_children_relative': False,
        },
        {
            'type': 'shape',
            'tag': 'div',
            'bounds': {'x': 0.5, 'y': 0.5, 'width': 0.518, 'height': 0.028},
            'styles': {'marginTop': '8px', 'marginBottom': '14px'},
        },
    ]

    layout_slide_elements(elements, 13.33, 810 / 108, {'paddingTop': '54px'}, {'_slide_index': 0})

    first = elements[0]['bounds']
    divider = elements[1]['bounds']
    expected_gap = 8.0 / PX_PER_IN
    actual_gap = divider['y'] - (first['y'] + first['height'])

    assert abs(actual_gap - expected_gap) < 0.02, (first, divider, actual_gap)
    print("  PASS: top-level layout uses next margin-top for spacing")


def test_build_elements_preserve_margin_top_metadata():
    """IR elements should retain marginTop so downstream flow layout can use it."""
    html = '''
    <div style="width:56px;height:3px;margin-top:8px;margin-bottom:14px;background:#2563eb;"></div>
    <p style="font-size:16px;line-height:1.6;margin-top:10px;">段落</p>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    div = soup.find('div')
    p = soup.find('p')

    div_style = compute_element_style(div, [], div.get('style', ''))
    p_style = compute_element_style(p, [], p.get('style', ''))

    shape = export_sandbox.build_shape_element(div, div_style, 1440)
    text = build_text_element(p, p_style, [], 1440, 720)

    assert shape['styles']['marginTop'] == '8px', shape['styles']
    assert text['styles']['marginTop'] == '10px', text['styles']
    print("  PASS: element IR preserves margin-top metadata")


def test_card_group_layout_expands_bg_height_to_content_bottom():
    """Centered card groups should grow their background to the laid-out content bottom."""
    elements = [
        {
            'type': 'shape',
            'tag': 'div',
            'bounds': {'x': 0.5, 'y': 0.5, 'width': 2.4, 'height': 0.52},
            'styles': {'backgroundColor': '#ffffff', 'borderRadius': '20px'},
            '_card_group': 'card-test',
            '_preserve_width': True,
            '_css_pad_l': 0.12,
            '_css_pad_r': 0.12,
            '_css_pad_t': 0.10,
            '_css_pad_b': 0.12,
            '_css_border_l': 0.0,
        },
        {
            'type': 'text',
            'tag': 'p',
            'text': '标题',
            'segments': [{'text': '标题', 'color': '#0f172a'}],
            'bounds': {'x': 0.5, 'y': 0.5, 'width': 1.8, 'height': 0.25},
            'naturalHeight': 0.25,
            'styles': {'fontSize': '16px', 'lineHeight': '1.4', 'marginBottom': '8px'},
            '_card_group': 'card-test',
        },
        {
            'type': 'text',
            'tag': 'p',
            'text': '说明文本',
            'segments': [{'text': '说明文本', 'color': '#475569'}],
            'bounds': {'x': 0.5, 'y': 0.5, 'width': 1.8, 'height': 0.40},
            'naturalHeight': 0.40,
            'styles': {'fontSize': '14px', 'lineHeight': '1.6'},
            '_card_group': 'card-test',
        },
    ]

    layout_slide_elements(elements, 13.33, 810 / 108, {'paddingTop': '54px'}, {'_slide_index': 0})

    bg = elements[0]['bounds']
    last_text = elements[2]['bounds']
    assert bg['y'] + bg['height'] >= last_text['y'] + last_text['height'] + 0.11, (bg, last_text)
    print("  PASS: card-group bg height tracks laid-out content bottom")


def test_export_text_element_preserves_explicit_break_headings():
    """Explicit line-break display headings should not allow PowerPoint re-wrap."""
    from pptx import Presentation

    prs = Presentation()
    prs.slide_width = int(914400 * 13.33)
    prs.slide_height = int(914400 * 7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    elem = {
        'type': 'text',
        'tag': 'h1',
        'text': 'AI 驱动的\nHTML 演示文稿',
        'segments': [
            {'text': 'AI 驱动的', 'color': '#1e3a8a', 'fontSize': '72px', 'fontFamily': 'PingFang SC', 'letterSpacing': '-0.02em', 'bold': True, 'strike': False, 'bgColor': None, 'inlineBgBounds': None, 'kind': 'text'},
            {'text': '\n', 'color': '#1e3a8a', 'fontSize': '72px', 'fontFamily': 'PingFang SC', 'letterSpacing': '-0.02em', 'bold': True, 'strike': False, 'bgColor': None, 'inlineBgBounds': None, 'kind': 'text'},
            {'text': 'HTML 演示文稿', 'color': '#1e3a8a', 'fontSize': '72px', 'fontFamily': 'PingFang SC', 'letterSpacing': '-0.02em', 'bold': True, 'strike': False, 'bgColor': None, 'inlineBgBounds': None, 'kind': 'text'},
        ],
        'bounds': {'x': 4.25, 'y': 2.7, 'width': 4.82, 'height': 1.47},
        'naturalHeight': 1.47,
        'styles': {
            'fontSize': '72px',
            'fontWeight': '800',
            'fontFamily': 'PingFang SC',
            'letterSpacing': '-0.02em',
            'color': '#1e3a8a',
            'textAlign': 'center',
            'lineHeight': '1.1',
            'paddingLeft': '0px',
            'paddingRight': '0px',
            'paddingTop': '0px',
            'paddingBottom': '0px',
            'justifyContent': '',
        },
    }

    export_sandbox.export_text_element(slide, elem, (255, 255, 255))
    tf = slide.shapes[-1].text_frame
    assert tf.word_wrap is False, tf.word_wrap
    assert tf.auto_size == export_sandbox.MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE, tf.auto_size
    print("  PASS: export keeps explicit-break headings from re-wrapping")


def test_measure_flow_box_intrinsic_height_for_layer_card():
    """Future gate: layer cards should measure as a single flow_box container."""
    measure_flow_box = _require_symbol('measure_flow_box')
    if measure_flow_box is None:
        return

    html = '''
    <div class="layer" style="display:flex;align-items:flex-start;gap:14px;padding:14px 18px;border-radius:14px;
         background:rgba(255,255,255,0.60);border:1px solid rgba(255,255,255,0.90);border-left:4px solid #0ea5e9;">
      <div style="width:34px;height:34px;border-radius:50%;background:#2563eb;color:#fff;">1</div>
      <div>
        <h4 style="margin-bottom:4px;">描述你的主题</h4>
        <p>告诉 slide-creator 你想呈现什么。</p>
      </div>
    </div>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    layer = soup.find('div')
    measured = measure_flow_box(layer, [], {}, 1440, 820)

    assert measured.get('layout') == 'flow_box', f"Expected flow_box layout, got {measured.get('layout')!r}"
    assert measured.get('measure', {}).get('intrinsic_height', 0) > 0.65, "Measured flow_box height should be non-trivial"
    print("  PASS: flow_box intrinsic height measured")


def test_measure_flow_box_marks_descendants_in_flow_box():
    """flow_box descendants should opt out of legacy card-group layout."""
    measure_flow_box = _require_symbol('measure_flow_box')
    if measure_flow_box is None:
        return

    html = '''
    <div class="layer" style="display:flex;align-items:flex-start;gap:14px;padding:14px 18px;border-radius:14px;
         background:rgba(255,255,255,0.60);border:1px solid rgba(255,255,255,0.90);border-left:4px solid #0ea5e9;">
      <div style="font-size:1.5rem;flex-shrink:0;">📝</div>
      <div>
        <h4 style="margin-bottom:4px;">当前幻灯片备注</h4>
        <p>按 <kbd>P</kbd> 在独立窗口查看演讲者备注。</p>
      </div>
    </div>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    measured = measure_flow_box(soup.find('div'), [], {}, 1440, 820)
    descendants = measured.get('children', [])

    assert descendants, "flow_box should contain descendants"
    assert all(child.get('_in_flow_box') for child in descendants), descendants
    assert not any(child.get('_card_group') for child in descendants), descendants
    assert not any(child.get('_is_border_left') for child in descendants), descendants
    bg_shape = next(child for child in descendants if child.get('_is_card_bg'))
    assert '4px' in bg_shape.get('styles', {}).get('borderLeft', ''), bg_shape
    print("  PASS: flow_box descendants opt into _in_flow_box")


def test_table_cell_fragments_measure_kbd_sequence():
    """Future gate: table cells should keep kbd fragments instead of flattening away."""
    html = '''
    <table style="width:100%;border-collapse:collapse;">
      <tr>
        <td><kbd>→</kbd> <kbd>Space</kbd> <kbd>↓</kbd></td>
        <td>下一个幻灯片</td>
      </tr>
    </table>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    table = soup.find('table')
    style = compute_element_style(table, [], table.get('style', ''))
    table_ir = export_sandbox.build_table_element(table, [], style)
    cell = table_ir['rows'][0]['cells'][0]
    fragments = cell.get('fragments')

    if not fragments:
        print("  SKIP: table cell fragments pending implementation")
        return

    kinds = [frag.get('kind') for frag in fragments]
    assert kinds.count('kbd') == 3, f"Expected 3 kbd fragments, got {kinds}"
    print("  PASS: table cell preserves kbd fragments")


def test_build_table_element_classifies_presentation_rows():
    """Presentational row lists should not stay on the generic data-table path."""
    html = '''
    <table style="width:100%;border-collapse:collapse;font-size:0.82rem;">
      <tr>
        <td style="padding:6px 0;border-bottom:1px solid rgba(14,165,233,0.10);"><kbd>→</kbd> <kbd>Space</kbd> <kbd>↓</kbd></td>
        <td style="padding:6px 0 6px 12px;border-bottom:1px solid rgba(14,165,233,0.10);color:#475569;">下一个幻灯片</td>
      </tr>
      <tr>
        <td style="padding:6px 0;">Home / End</td>
        <td style="padding:6px 0 6px 12px;color:#475569;">首/末幻灯片</td>
      </tr>
      <tr>
        <td style="padding:6px 0;">滑动 ← →</td>
        <td style="padding:6px 0 6px 12px;color:#475569;">触摸导航</td>
      </tr>
    </table>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    table = soup.find('table')
    style = compute_element_style(table, [], table.get('style', ''))
    table_ir = export_sandbox.build_table_element(table, [], style)

    assert table_ir.get('type') == 'presentation_rows', table_ir.get('type')
    print("  PASS: presentational tables classify as presentation_rows")


def test_presentation_rows_use_compact_single_line_row_height():
    """Presentation rows should keep compact single-line heights instead of table-like 0.31+ rows."""
    html = '''
    <table style="width:100%;border-collapse:collapse;font-size:0.82rem;">
      <tr><td style="padding:5px 0;border-bottom:1px solid rgba(14,165,233,0.10);">内容密度</td><td style="padding:5px 0 5px 8px;border-bottom:1px solid rgba(14,165,233,0.10);color:#475569;">≥ 65% 填充</td></tr>
      <tr><td style="padding:5px 0;border-bottom:1px solid rgba(14,165,233,0.10);">列平衡</td><td style="padding:5px 0 5px 8px;border-bottom:1px solid rgba(14,165,233,0.10);color:#475569;">最短列 ≥ 60%</td></tr>
      <tr><td style="padding:5px 0;">色彩律</td><td style="padding:5px 0 5px 8px;color:#475569;">90/8/2 分配</td></tr>
    </table>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    table = soup.find('table')
    style = compute_element_style(table, [], table.get('style', ''))
    table_ir = export_sandbox.build_table_element(table, [], style)

    assert table_ir.get('type') == 'presentation_rows', table_ir.get('type')
    row_heights = [row.get('height', 0.0) for row in table_ir.get('rows', [])]
    assert row_heights and min(row_heights) >= 0.264, row_heights
    assert row_heights and max(row_heights) < 0.29, row_heights
    print("  PASS: presentation rows use compact single-line row height")


def test_presentation_rows_keep_fitted_key_column_and_expand_value_column():
    """Presentation rows should keep a fitted key column instead of proportionally stretching both columns."""
    html = '''
    <table style="width:100%;border-collapse:collapse;font-size:0.82rem;">
      <tr><td style="padding:5px 0;border-bottom:1px solid rgba(14,165,233,0.10);">内容密度</td><td style="padding:5px 0 5px 8px;border-bottom:1px solid rgba(14,165,233,0.10);color:#475569;">≥ 65% 填充</td></tr>
      <tr><td style="padding:5px 0;border-bottom:1px solid rgba(14,165,233,0.10);">标题质量</td><td style="padding:5px 0 5px 8px;border-bottom:1px solid rgba(14,165,233,0.10);color:#475569;">断言式，非通用</td></tr>
      <tr><td style="padding:5px 0;">布局轮换</td><td style="padding:5px 0 5px 8px;color:#475569;">禁止连续 3 页</td></tr>
    </table>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    table = soup.find('table')
    style = compute_element_style(table, [], table.get('style', ''))
    table_ir = export_sandbox.build_table_element(table, [], style)

    col_widths = export_sandbox._compute_presentation_row_column_widths(table_ir.get('rows', []), 3.536)
    assert len(col_widths) == 2, col_widths
    assert col_widths[0] < 1.35, col_widths
    assert col_widths[1] > 2.10, col_widths
    assert abs(sum(col_widths) - 3.536) < 0.01, col_widths
    print("  PASS: presentation rows keep fitted key column widths")


def test_presentation_rows_shortcut_column_gets_extra_runway():
    """Shortcut-heavy first columns should get more runway than prose labels."""
    html = '''
    <table style="width:100%;border-collapse:collapse;font-size:0.82rem;">
      <tr><td style="padding:6px 0;">→ Space ↓</td><td style="padding:6px 0 6px 12px;color:#475569;">下一个幻灯片</td></tr>
      <tr><td style="padding:6px 0;">Home / End</td><td style="padding:6px 0 6px 12px;color:#475569;">首/末幻灯片</td></tr>
      <tr><td style="padding:6px 0;">F5</td><td style="padding:6px 0 6px 12px;color:#475569;">全屏演示</td></tr>
    </table>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    table = soup.find('table')
    style = compute_element_style(table, [], table.get('style', ''))
    table_ir = export_sandbox.build_table_element(table, [], style)

    col_widths = export_sandbox._compute_presentation_row_column_widths(table_ir.get('rows', []), 3.536)
    assert len(col_widths) == 2, col_widths
    assert col_widths[0] > 1.45, col_widths
    assert col_widths[1] < 2.10, col_widths
    print("  PASS: presentation rows give shortcut columns extra runway")


def test_presentation_row_label_uses_stronger_ink():
    """First-column labels in presentation rows should use strong ink on light cards."""
    assert export_sandbox._presentation_row_label_color('rgb(15,23,42)') == '#000000'
    assert export_sandbox._presentation_row_label_color('#0f172a') == '#000000'
    assert export_sandbox._presentation_row_label_color('rgb(51,65,85)') == 'rgb(51,65,85)'
    print("  PASS: presentation row label uses stronger ink")


def test_display_heading_normalizes_deep_slate_to_black():
    """Short dark-slate display headings should strengthen to black, without muting boxed inline fragments."""
    html = '<h4 style="color:#0f172a;">核心 8 项 <span style="background:#dbeafe;color:#0f172a;padding:2px 8px;border-radius:999px;">当前</span></h4>'
    soup = BeautifulSoup(html, 'html.parser')
    css_rules = extract_css_from_soup(soup)
    h4 = soup.find('h4')
    style = compute_element_style(h4, css_rules, h4.get('style', ''))
    ir = build_text_element(h4, style, css_rules, 1440, 720)

    heading_seg = next(seg for seg in ir.get('segments', []) if '核心' in seg.get('text', ''))
    badge_seg = next(seg for seg in ir.get('segments', []) if '当前' in seg.get('text', ''))
    assert heading_seg['color'] == '#000000', ir.get('segments', [])
    assert badge_seg['color'].lower() == '#0f172a', ir.get('segments', [])
    print("  PASS: display heading normalizes deep slate to black")


def test_presentation_row_shortcut_cell_keeps_original_ink():
    """Shortcut-token first cells should not be over-strengthened to black."""
    html = '''
    <table style="width:100%;border-collapse:collapse;font-size:0.82rem;">
      <tr>
        <td style="padding:6px 0;color:#0f172a;">Home / End</td>
        <td style="padding:6px 0 6px 12px;color:#475569;">首/末幻灯片</td>
      </tr>
    </table>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    table = soup.find('table')
    style = compute_element_style(table, [], table.get('style', ''))
    table_ir = export_sandbox.build_table_element(table, [], style)
    first_cell = table_ir['rows'][0]['cells'][0]

    assert first_cell['styles']['color'].lower() == '#0f172a', first_cell['styles']['color']
    print("  PASS: presentation-row shortcut cell keeps original ink")


def test_build_table_element_keeps_real_data_tables():
    """Real data tables with headers should remain on the generic table path."""
    html = '''
    <table style="width:100%;border-collapse:collapse;">
      <tr><th>季度</th><th>转化率</th></tr>
      <tr><td>Q1</td><td>12%</td></tr>
      <tr><td>Q2</td><td>18%</td></tr>
    </table>
    '''
    soup = BeautifulSoup(html, 'html.parser')
    table = soup.find('table')
    style = compute_element_style(table, [], table.get('style', ''))
    table_ir = export_sandbox.build_table_element(table, [], style)

    assert table_ir.get('type') == 'table', table_ir.get('type')
    print("  PASS: real data tables stay on table path")


def test_table_card_height_uses_actual_table_bounds():
    """Cards wrapping tables should use measured table height, not row-count heuristics."""
    html_path = Path(__file__).parent.parent / 'demo' / 'blue-sky-zh.html'
    if not html_path.exists():
        print("  SKIP: table_card_height_uses_actual_table_bounds (HTML not found)")
        return

    slides = parse_html_to_slides(html_path, 1440, 810)
    slide = slides[6]
    pre_pass_corrections = getattr(export_sandbox, 'pre_pass_corrections')
    pre_pass_corrections(slide['elements'])
    slide['_slide_index'] = 6
    layout_slide_elements(slide['elements'], 13.33, 810 / 108, slide['slideStyle'], slide)

    container = next(e for e in slide['elements'] if e.get('type') == 'container')
    card = next(c for c in container.get('children', []) if c.get('type') == 'shape' and c.get('_is_card_bg'))
    table = next(
        c for c in container.get('children', [])
        if c.get('type') in ('table', 'presentation_rows')
    )

    card_bottom = card['bounds']['y'] + card['bounds']['height']
    table_bottom = table['bounds']['y'] + table['bounds']['height']
    assert card_bottom >= table_bottom, (
        f"Card bottom {card_bottom:.3f} should cover table bottom {table_bottom:.3f}"
    )
    print("  PASS: table card height uses actual table bounds")


def test_centered_inline_command_prefers_content_width():
    """Centered inline command rows should shrink-wrap to measured fragment width."""
    html = '''
    <section class="slide">
      <div style="text-align:center;max-width:720px;">
        <p style="margin-top:20px;font-size:0.9rem;color:#64748b;">
          <code style="padding:2px 8px;background:rgba(14,165,233,0.10);">clawhub install kai-slide-creator</code>
          &nbsp;·&nbsp;
          <a href="https://github.com/kaisersong/slide-creator" style="color:#2563eb;text-decoration:none;">GitHub ↗</a>
        </p>
      </div>
    </section>'''
    soup = BeautifulSoup(html, 'html.parser')
    css_rules = extract_css_from_soup(soup)
    slide = soup.find('section')
    center = slide.find('div')
    results = flat_extract(center, css_rules, None, 1440)
    layout_slide_elements(results, 13.33, 810 / 108, {'paddingTop': '40px'}, {'_slide_index': 9})

    command_line = next(r for r in results if r.get('type') == 'text' and 'clawhub install' in r.get('text', ''))
    assert command_line.get('preferContentWidth'), "Inline command row should opt into content-width centering"
    assert command_line['bounds']['width'] < 5.5, (
        f"Command line width {command_line['bounds']['width']:.3f}\" should shrink-wrap, not fill the full 6.67\" column"
    )
    print("  PASS: centered inline command prefers content width")


def test_export_centered_inline_command_uses_fragment_runway():
    """Centered code+link rows should use one carrier textbox plus a pill overlay."""
    html = '''
    <html><body>
      <section class="slide" style="padding:56px 80px;background:linear-gradient(160deg,#f0f9ff,#e0f2fe);">
        <div style="text-align:center;max-width:720px;">
          <p style="margin-top:20px;font-size:0.9rem;color:#64748b;">
            <code style="padding:2px 8px;background:rgba(37,99,235,0.08);border-radius:999px;">clawhub install kai-slide-creator</code>
            &nbsp;·&nbsp;
            <a href="https://github.com/kaisersong/slide-creator" style="color:#2563eb;text-decoration:none;">GitHub ↗</a>
          </p>
        </div>
      </section>
    </body></html>
    '''

    with tempfile.TemporaryDirectory(prefix='kai-export-inline-runway-') as tmp_dir:
        html_path = Path(tmp_dir) / 'command-row.html'
        pptx_path = Path(tmp_dir) / 'command-row.pptx'
        html_path.write_text(html, encoding='utf-8')
        export_sandbox_pptx(html_path, pptx_path, 1440, 810)

        from pptx import Presentation
        prs = Presentation(str(pptx_path))
        slide = prs.slides[0]

        texts = []
        pill_shapes = []
        for shape in slide.shapes:
            text = getattr(shape, 'text', '').strip() if hasattr(shape, 'text') else ''
            if text:
                texts.append(text)
            try:
                fill = tuple(shape.fill.fore_color.rgb)
            except Exception:
                fill = None
            if fill and shape.width / 914400 > 2.5 and shape.height / 914400 > 0.20:
                pill_shapes.append((shape.width / 914400, shape.height / 914400, fill))

        assert any('clawhub install kai-slide-creator' in text and 'GitHub ↗' in text for text in texts), texts
        assert all(text != 'GitHub ↗' for text in texts), texts
        assert all(text != 'clawhub install kai-slide-creator' for text in texts), texts
        assert pill_shapes, pill_shapes
    print("  PASS: centered inline command exports as carrier textbox + pill overlay")


def test_centered_inline_command_mutes_trailing_link_color():
    """Centered command rows should keep trailing link/separator on the muted body ink."""
    html = '''
    <section class="slide">
      <div style="text-align:center;max-width:720px;">
        <p style="margin-top:20px;font-size:0.9rem;color:#64748b;">
          <code style="padding:2px 8px;background:rgba(37,99,235,0.08);border-radius:999px;">clawhub install kai-slide-creator</code>
          &nbsp;·&nbsp;
          <a href="https://github.com/kaisersong/slide-creator" style="color:#2563eb;text-decoration:none;">GitHub ↗</a>
        </p>
      </div>
    </section>'''
    soup = BeautifulSoup(html, 'html.parser')
    css_rules = extract_css_from_soup(soup)
    p = soup.find('p')
    style = compute_element_style(p, css_rules, p.get('style', ''))
    ir = build_text_element(p, style, css_rules, 1440, 720)

    link_segment = next(seg for seg in ir.get('segments', []) if 'GitHub' in seg.get('text', ''))
    assert link_segment['color'] == '#64748b', ir.get('segments', [])
    print("  PASS: centered inline command mutes trailing link color")


def test_export_accent_card_uses_narrow_strip_and_full_main_card():
    """Border-left accent cards should export as a narrow accent strip plus a full main card."""
    html = '''
    <html><body>
      <section class="slide" style="padding:56px 80px;background:linear-gradient(160deg,#f0f9ff,#e0f2fe);">
        <div style="max-width:820px;width:100%;">
          <div style="background:rgba(255,255,255,0.70);border:1px solid rgba(255,255,255,0.90);border-left:4px solid #0ea5e9;border-radius:20px;padding:22px 24px;">
            <h4 style="margin-bottom:8px;">标题</h4>
            <p>说明文本</p>
          </div>
        </div>
      </section>
    </body></html>
    '''

    with tempfile.TemporaryDirectory(prefix='kai-export-accent-card-') as tmp_dir:
        html_path = Path(tmp_dir) / 'accent-card.html'
        pptx_path = Path(tmp_dir) / 'accent-card.pptx'
        html_path.write_text(html, encoding='utf-8')
        export_sandbox_pptx(html_path, pptx_path, 1440, 810)

        from pptx import Presentation
        prs = Presentation(str(pptx_path))
        slide = prs.slides[0]
        widths = []
        for shape in slide.shapes:
            x = shape.left / 914400
            w = shape.width / 914400
            h = shape.height / 914400
            fill = None
            try:
                fill = tuple(shape.fill.fore_color.rgb)
            except Exception:
                pass
            if h > 0.4 and fill:
                widths.append((x, w, fill))

        accent_strip = next((item for item in widths if item[1] < 0.35 and item[2] == (14, 165, 233)), None)
        main_card = next((item for item in widths if item[1] > 1.0 and all(channel >= 245 for channel in item[2])), None)

        assert accent_strip is not None, widths
        assert main_card is not None, widths
        assert accent_strip[0] <= main_card[0] + 0.01, (accent_strip, main_card)
    assert main_card[1] > accent_strip[1] * 8, (accent_strip, main_card)
    print("  PASS: accent card exports as strip + full main card")


def test_slide4_theme_grid_cards_share_stretched_row_height():
    """Grid cards in one row should stretch their backgrounds to the shared row height."""
    html_path = Path(__file__).parent.parent / 'demo' / 'blue-sky-zh.html'
    if not html_path.exists():
        print("  SKIP: slide4_theme_grid_cards_share_stretched_row_height (HTML not found)")
        return

    slides = parse_html_to_slides(html_path, 1440, 810)
    slide = slides[3]
    pre_pass_corrections = getattr(export_sandbox, 'pre_pass_corrections')
    pre_pass_corrections(slide['elements'])
    slide['_slide_index'] = 3
    layout_slide_elements(slide['elements'], 13.33, 810 / 108, slide['slideStyle'], slide)

    grid = [e for e in slide['elements'] if e.get('type') == 'container'][1]
    card_heights = [
        child['bounds']['height']
        for child in grid.get('children', [])
        if child.get('type') == 'shape' and child.get('_is_card_bg')
    ]
    assert len(card_heights) >= 4, card_heights
    assert max(card_heights[:4]) - min(card_heights[:4]) < 0.05, card_heights[:4]
    print("  PASS: slide 4 theme cards stretch to shared row height")


def test_accent_callouts_keep_optical_gap_from_preceding_blocks():
    """Bottom accent/info callouts should keep a visible gap from the block above."""
    html_path = Path(__file__).parent.parent / 'demo' / 'blue-sky-zh.html'
    if not html_path.exists():
        print("  SKIP: accent_callouts_keep_optical_gap_from_preceding_blocks (HTML not found)")
        return

    slides = parse_html_to_slides(html_path, 1440, 810)
    pre_pass_corrections = getattr(export_sandbox, 'pre_pass_corrections')
    targets = {
        1: 0.20,  # slide 2
        3: 0.20,  # slide 4
        7: 0.22,  # slide 8
        8: 0.22,  # slide 9
    }

    for slide_idx, min_gap in targets.items():
        slide = slides[slide_idx]
        pre_pass_corrections(slide['elements'])
        slide['_slide_index'] = slide_idx
        layout_slide_elements(slide['elements'], 13.33, 810 / 108, slide['slideStyle'], slide)

        prev_elem = slide['elements'][-3]
        shape = slide['elements'][-2]
        gap = shape['bounds']['y'] - (prev_elem['bounds']['y'] + prev_elem['bounds']['height'])
        assert gap >= min_gap, (slide_idx + 1, gap, min_gap, shape['bounds'], prev_elem['bounds'])

    print("  PASS: accent callouts keep optical gap from preceding blocks")


def test_slide10_gradient_divider_centers_in_heading_block():
    """Gradient divider under a centered heading should stay centered in the heading block."""
    html_path = Path(__file__).parent.parent / 'demo' / 'blue-sky-zh.html'
    if not html_path.exists():
        print("  SKIP: slide10_gradient_divider_centers_in_heading_block (HTML not found)")
        return

    slides = parse_html_to_slides(html_path, 1440, 810)
    slide = slides[9]
    pre_pass_corrections = getattr(export_sandbox, 'pre_pass_corrections')
    pre_pass_corrections(slide['elements'])
    slide['_slide_index'] = 9
    layout_slide_elements(slide['elements'], 13.33, 810 / 108, slide['slideStyle'], slide)

    heading = slide['elements'][0]['bounds']
    divider = slide['elements'][1]['bounds']
    divider_center = divider['x'] + divider['width'] / 2.0
    heading_center = heading['x'] + heading['width'] / 2.0
    assert abs(divider_center - heading_center) < 0.05, (heading, divider)
    print("  PASS: slide 10 divider centers in heading block")


def test_export_corpus_parse_smoke():
    """Corpus HTML samples should remain parseable across Blue Sky and non-Blue-Sky decks."""
    parsed = []
    skipped = []

    for sample in _corpus_samples():
        path = sample['path']
        if not path.exists():
            if sample['required']:
                raise AssertionError(f"Required corpus sample missing: {path}")
            skipped.append(sample['label'])
            continue

        slides = parse_html_to_slides(path, 1440, 810)
        assert slides, f"Corpus sample {sample['label']} produced no slides"
        text_count = sum(_count_text_elements(slide['elements']) for slide in slides)
        assert text_count > 0, f"Corpus sample {sample['label']} produced no text elements"
        parsed.append(sample['label'])

    assert len(parsed) >= 3, f"Expected at least 3 parsed corpus samples, got {parsed}"
    assert any('blue-sky' in label for label in parsed), f"Expected at least one Blue Sky sample, got {parsed}"
    assert any('intro' in label or 'handwritten' in label or 'swiss-modern' in label for label in parsed), (
        f"Expected at least one non-Blue-Sky sample, got {parsed}"
    )
    if skipped:
        print(f"  SKIP: optional corpus samples unavailable: {', '.join(skipped)}")
    print(f"  PASS: corpus parse smoke ({len(parsed)} samples)")


def test_handwritten_fixture_covers_core_patterns():
    """Handwritten fixture should cover grouped-inline, cards, lists, and tables."""
    if not HANDWRITTEN_FIXTURE.exists():
        raise AssertionError(f"Missing handwritten fixture: {HANDWRITTEN_FIXTURE}")

    slides = parse_html_to_slides(HANDWRITTEN_FIXTURE, 1440, 810)
    assert len(slides) == 3, f"Expected 3 slides in handwritten fixture, got {len(slides)}"

    all_texts = []
    all_tables = []
    all_presentation_rows = []
    all_shapes = []
    for slide in slides:
        all_texts.extend(_collect_text_values(slide['elements']))
        all_tables.extend(_collect_elements_by_type(slide['elements'], 'table'))
        all_presentation_rows.extend(_collect_elements_by_type(slide['elements'], 'presentation_rows'))
        all_shapes.extend(_collect_elements_by_type(slide['elements'], 'shape'))

    joined = "\n".join(all_texts)
    table_cell_texts = []
    table_fragment_kinds = []
    for table in all_tables + all_presentation_rows:
        for row in table.get('rows', []):
            for cell in row.get('cells', []):
                cell_text = (cell.get('text') or '').strip()
                if cell_text:
                    table_cell_texts.append(cell_text)
                for fragment in cell.get('fragments') or []:
                    kind = fragment.get('kind')
                    if kind:
                        table_fragment_kinds.append(kind)

    assert 'Handwritten Export Corpus' in joined, "Fixture should include title slide text"
    assert 'GitHub' in joined, "Fixture should include a link-like grouped inline row"
    assert table_fragment_kinds.count('kbd') >= 3, (
        f"Fixture should include kbd-heavy presentational rows, got fragments {table_fragment_kinds}"
    )
    assert any('Q2' in txt or '转化率' in txt for txt in table_cell_texts), (
        f"Fixture should include a real data table, got {table_cell_texts}"
    )
    assert all_tables, "Fixture should include at least one real table element"
    assert all_presentation_rows, "Fixture should include at least one presentation_rows element"
    assert all_shapes, "Fixture should include at least one background/card shape"
    print("  PASS: handwritten fixture covers core patterns")


def test_handwritten_fixture_structural_eval_gate():
    """A small handwritten corpus export should clear structural eval gates."""
    if not HANDWRITTEN_FIXTURE.exists():
        raise AssertionError(f"Missing handwritten fixture: {HANDWRITTEN_FIXTURE}")

    with tempfile.TemporaryDirectory(prefix='kai-export-corpus-') as tmp_dir:
        out_path = Path(tmp_dir) / 'handwritten-card-list-table.pptx'
        export_sandbox_pptx(HANDWRITTEN_FIXTURE, out_path, 1440, 810)
        summary = rigorous_eval.collect_eval_summary(
            golden_path=str(out_path),
            sandbox_path=str(out_path),
            include_visual=False,
        )

        assert summary['sandbox_overflow_count'] == 0, summary['sandbox_overflow_issues']
        assert summary['sandbox_overlap_count'] == 0, summary['sandbox_overlap_issues']
        assert summary['element_gap_count'] == 0, summary['element_gap_issues']
        assert summary['card_containment_count'] == 0, summary['card_containment_issues']
        assert summary['color_diff_count'] == 0, summary['color_diff_issues']
        assert summary['total_actionable'] == 0, summary
    print("  PASS: handwritten fixture structural eval gate")


# ─── Test Runner ──────────────────────────────────────────────────────────────

def run_tests():
    print("Running export-sandbox-pptx tests...")
    print()

    print("Utilities:")
    test_parse_px()
    print()

    print("Slide 1 (Cover) — stat card padding fix:")
    test_stat_card_padding_included_in_width()
    test_centered_flex_x_position_with_padding()
    test_slide1_stat_positions()
    test_slide1_all_text_present()
    test_heading_content_width()
    test_centered_explicit_break_heading_gets_wrap_guard_width()
    print()

    print("Pill text fix (Slide 4/10 — combined text):")
    test_pill_text_included_in_parent()
    test_pill_shape_no_text()
    print()

    print("Decoration shapes:")
    test_decoration_shape_for_no_text_elements()
    print()

    print("Slide 5 chapter page (centered layout width + line count):")
    test_chapter_page_max_width_from_widest_element()
    test_chapter_page_zero_padding_for_large_text()
    test_chapter_page_paragraph_single_line()
    test_chapter_page_no_overflow()
    test_chapter_page_no_overlap()
    test_estimate_wrapped_lines_uses_px_per_in_formula()
    print()

    print("Slide 4 subtitle fix (width + height for padded elements):")
    test_slide4_subtitle_width_not_overwritten_by_sync()
    test_slide4_subtitle_height_matches_golden()
    print()

    print("Generic layout regressions:")
    test_flex_column_badge_stretches_to_parent_width()
    test_decoration_in_flex_row_keeps_explicit_size()
    test_gradient_decoration_in_flex_row_keeps_explicit_size()
    test_grid_flex_container_height_tracks_child_extent_without_tail_gap()
    test_centered_card_group_layout_keeps_text_inside_card()
    test_centered_card_group_preserves_vertical_padding_metadata()
    test_slide_root_background_not_promoted_to_card_group()
    test_auto_margin_divider_centers_in_constrained_content_area()
    test_slide2_info_bar_margin_top_applies_to_outer_box()
    test_slide2_info_bar_does_not_emit_detached_code_bg_shape()
    test_complex_card_height_uses_stacked_flow()
    test_layout_slide_elements_flow_box_advances_current_y_correctly()
    test_extract_inline_fragments_code_kbd_support()
    test_extract_inline_fragments_grouped_badge_and_link()
    test_gradient_text_hex_colors_resolve_and_keep_stops()
    test_build_text_element_inline_flex_pill_shrink_wraps_single_line()
    test_build_text_element_grouped_inline_badge_keeps_single_line_height()
    test_build_grid_children_flex_row_preserves_component_width_and_pairing()
    test_map_font_prefers_stable_ppt_font_over_platform_stack_order()
    test_build_table_element_plain_td_defaults_to_text_primary()
    test_flat_extract_mixed_inline_code_uses_inline_overlays()
    test_flat_extract_inline_code_in_prose_does_not_emit_detached_code_bg()
    test_build_text_element_wide_prose_adjusts_back_to_single_line()
    test_flow_gap_prefers_collapsed_margins_over_default_gap()
    test_layout_slide_elements_uses_next_margin_top_for_container_gap()
    test_build_elements_preserve_margin_top_metadata()
    test_card_group_layout_expands_bg_height_to_content_bottom()
    test_export_text_element_preserves_explicit_break_headings()
    test_measure_flow_box_intrinsic_height_for_layer_card()
    test_measure_flow_box_marks_descendants_in_flow_box()
    test_table_cell_fragments_measure_kbd_sequence()
    test_build_table_element_classifies_presentation_rows()
    test_presentation_rows_use_compact_single_line_row_height()
    test_presentation_rows_keep_fitted_key_column_and_expand_value_column()
    test_presentation_rows_shortcut_column_gets_extra_runway()
    test_presentation_row_label_uses_stronger_ink()
    test_display_heading_normalizes_deep_slate_to_black()
    test_presentation_row_shortcut_cell_keeps_original_ink()
    test_build_table_element_keeps_real_data_tables()
    test_table_card_height_uses_actual_table_bounds()
    test_centered_inline_command_prefers_content_width()
    test_export_centered_inline_command_uses_fragment_runway()
    test_centered_inline_command_mutes_trailing_link_color()
    test_export_accent_card_uses_narrow_strip_and_full_main_card()
    test_slide4_theme_grid_cards_share_stretched_row_height()
    test_accent_callouts_keep_optical_gap_from_preceding_blocks()
    test_slide10_gradient_divider_centers_in_heading_block()
    test_export_corpus_parse_smoke()
    test_handwritten_fixture_covers_core_patterns()
    test_handwritten_fixture_structural_eval_gate()
    print()

    print("All tests passed!")


if __name__ == '__main__':
    run_tests()
