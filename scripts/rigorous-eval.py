#!/usr/bin/env python3
"""
rigorous-eval.py — Multi-dimensional HTML→PPTX export quality evaluation.

Runs 5 independent checks in parallel:
1. compare-visual-comprehensive.py — pixel/property match vs golden
2. overflow-check — text wider than containing shape/card
3. overlap-check — adjacent elements overlapping beyond tolerance
4. position-drift — systematic offset analysis (Y shift, X shift distribution)
5. element-count-gap — missing/extra element categorization

Usage:
    python3 scripts/rigorous-eval.py
"""

import sys
import subprocess
from pathlib import Path
from pptx import Presentation
from pptx.util import Emu
from pptx.enum.dml import MSO_FILL
from pptx.enum.shapes import MSO_SHAPE_TYPE
from typing import List, Dict, Tuple, Optional

GOLDEN = "/tmp/kai-html-export-golden.pptx"
SANDBOX = "output.pptx"

EMU_PER_IN = 914400


# ─── 1. Visual comprehensive (wrapper) ─────────────────────────────────

def run_visual_comprehensive() -> str:
    """Run compare-visual-comprehensive.py and return summary."""
    result = subprocess.run(
        [sys.executable, "scripts/compare-visual-comprehensive.py"],
        capture_output=True, text=True, timeout=60
    )
    # Extract per-slide scores and total
    lines = result.stdout.split("\n")
    summary = []
    in_summary = False
    for line in lines:
        if "SUMMARY" in line:
            in_summary = True
        if in_summary:
            summary.append(line)
    return "\n".join(summary) if summary else result.stdout[-500:]


# ─── 2. Overflow check ─────────────────────────────────────────────────

def _is_nav_dot(shape) -> bool:
    try:
        if shape.width / EMU_PER_IN < 0.1 and shape.height / EMU_PER_IN < 0.1:
            return True
    except Exception:
        pass
    return False


def _get_fill_color(shape) -> Optional[str]:
    try:
        if shape.fill.type == MSO_FILL.SOLID:
            return f"#{shape.fill.fore_color.rgb}"
    except Exception:
        pass
    return None


def _text_runs_info(shape) -> List[Dict]:
    """Extract text run info: text, font_size_pt, is_bold, color."""
    runs = []
    if not hasattr(shape, "text_frame"):
        return runs
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            runs.append({
                "text": run.text,
                "size": run.font.size.pt if run.font.size else None,
                "bold": bool(run.font.bold),
            })
    return runs


def check_overflow(pptx_path: str, label: str) -> List[str]:
    """Check if text elements overflow their container width."""
    prs = Presentation(pptx_path)
    issues = []
    slide_w = prs.slide_width / EMU_PER_IN

    for si, slide in enumerate(prs.slides):
        shapes = list(slide.shapes)
        for i, shape in enumerate(shapes):
            if _is_nav_dot(shape):
                continue
            if not hasattr(shape, "text_frame") or not shape.text:
                continue
            text = shape.text.strip()
            if len(text) < 5:
                continue

            w_in = shape.width / EMU_PER_IN
            h_in = shape.height / EMU_PER_IN
            x_in = shape.left / EMU_PER_IN

            # Check: text element extends beyond slide right edge
            right_edge = x_in + w_in
            if right_edge > slide_w + 0.05:
                issues.append(
                    f"  [{label}] Slide {si+1}: element '{text[:40]}' "
                    f"extends beyond slide right edge "
                    f"(x={x_in:.2f}\", w={w_in:.2f}\", right={right_edge:.2f}\", slide_w={slide_w:.2f}\")"
                )

            # Check: text height severely underestimated
            # Estimate minimum height needed for the text
            runs = _text_runs_info(shape)
            if runs:
                max_fs = max(r["size"] or 12 for r in runs)
                lines_est = max(1, len(text) / max(w_in * 6 / (max_fs / 12), 1))
                min_h = lines_est * max_fs * 1.3 / 72
                if h_in < min_h * 0.5 and len(text) > 30:
                    issues.append(
                        f"  [{label}] Slide {si+1}: text '{text[:30]}...' "
                        f"height {h_in:.3f}\" << estimated min {min_h:.3f}\" "
                        f"(likely truncated)"
                    )

    return issues


# ─── 3. Overlap check ──────────────────────────────────────────────────

def check_overlap(pptx_path: str, label: str) -> List[str]:
    """Check if non-background elements overlap each other excessively."""
    prs = Presentation(pptx_path)
    issues = []

    for si, slide in enumerate(prs.slides):
        shapes = [s for s in slide.shapes if not _is_nav_dot(s)]
        # Only check text elements against each other
        text_shapes = []
        for s in shapes:
            if hasattr(s, "text_frame") and s.text.strip():
                text_shapes.append({
                    "text": s.text.strip()[:40],
                    "x": s.left / EMU_PER_IN,
                    "y": s.top / EMU_PER_IN,
                    "w": s.width / EMU_PER_IN,
                    "h": s.height / EMU_PER_IN,
                })

        # Check pairwise overlap
        for i in range(len(text_shapes)):
            for j in range(i + 1, len(text_shapes)):
                a, b = text_shapes[i], text_shapes[j]
                # Compute overlap
                x_overlap = min(a["x"] + a["w"], b["x"] + b["w"]) - max(a["x"], b["x"])
                y_overlap = min(a["y"] + a["h"], b["y"] + b["h"]) - max(a["y"], b["y"])

                if x_overlap > 0.05 and y_overlap > 0.05:
                    # Significant overlap — check if one isn't a bg shape
                    area_a = a["w"] * a["h"]
                    area_b = b["w"] * b["h"]
                    overlap_area = x_overlap * y_overlap
                    overlap_ratio = overlap_area / min(area_a, area_b) if min(area_a, area_b) > 0 else 0

                    if overlap_ratio > 0.3:
                        issues.append(
                            f"  [{label}] Slide {si+1}: overlap {overlap_ratio:.0%} between "
                            f"'{a['text'][:25]}' and '{b['text'][:25]}'"
                        )

    return issues


# ─── 4. Position drift analysis ────────────────────────────────────────

def check_position_drift() -> List[str]:
    """Analyze systematic position offsets between golden and sandbox."""
    from collections import defaultdict

    try:
        from compare_visual_comprehensive import match_elements_comprehensive
    except ImportError:
        # Inline minimal matching
        pass

    prs_g = Presentation(GOLDEN)
    prs_s = Presentation(SANDBOX)
    num_slides = min(len(prs_g.slides), len(prs_s.slides))

    issues = []
    all_dx = []
    all_dy = []
    per_slide_drift = []

    for si in range(num_slides):
        g_shapes = [(s, s.text.strip()[:30] if hasattr(s, "text") and s.text.strip() else None)
                    for s in prs_g.slides[si].shapes if not _is_nav_dot(s)]
        s_shapes = [(s, s.text.strip()[:30] if hasattr(s, "text") and s.text.strip() else None)
                    for s in prs_s.slides[si].shapes if not _is_nav_dot(s)]

        dxs, dys = [], []
        # Match by text content
        for gs, gt in g_shapes:
            if not gt:
                continue
            for ss, st in s_shapes:
                if st and gt == st:
                    dx = (ss.left - gs.left) / EMU_PER_IN
                    dy = (ss.top - gs.top) / EMU_PER_IN
                    dxs.append(dx)
                    dys.append(dy)
                    break

        if dxs:
            avg_dx = sum(dxs) / len(dxs)
            avg_dy = sum(dys) / len(dys)
            max_dx = max(abs(d) for d in dxs)
            max_dy = max(abs(d) for d in dys)
            per_slide_drift.append((si + 1, avg_dx, avg_dy, max_dx, max_dy))
            all_dx.extend(dxs)
            all_dy.extend(dys)

    if all_dx:
        issues.append(f"\n  [POSITION DRIFT] Analysis across {num_slides} slides:")
        issues.append(f"    Mean X drift: {sum(all_dx)/len(all_dx):+.3f}\"  |  Mean Y drift: {sum(all_dy)/len(all_dy):+.3f}\"")
        issues.append(f"    Max |X| drift: {max(abs(d) for d in all_dx):.3f}\"  |  Max |Y| drift: {max(abs(d) for d in all_dy):.3f}\"")

        for slide_num, avg_dx, avg_dy, max_dx, max_dy in per_slide_drift:
            flag = ""
            if abs(avg_dy) > 0.15 or abs(avg_dx) > 0.15:
                flag = " ⚠ SYSTEMATIC"
            issues.append(
                f"    Slide {slide_num}: avg(dx={avg_dx:+.3f}\", dy={avg_dy:+.3f}\") "
                f"max(|dx|={max_dx:.3f}\", |dy|={max_dy:.3f}\"){flag}"
            )

    return issues


# ─── 5. Element count gap analysis ──────────────────────────────────────

def check_element_gaps() -> List[str]:
    """Categorize missing and extra elements."""
    prs_g = Presentation(GOLDEN)
    prs_s = Presentation(SANDBOX)
    num_slides = min(len(prs_g.slides), len(prs_s.slides))

    issues = []
    for si in range(num_slides):
        g_texts = set()
        s_texts = set()
        g_shapes = 0
        s_shapes = 0

        for s in prs_g.slides[si].shapes:
            if _is_nav_dot(s):
                continue
            if hasattr(s, "text") and s.text.strip():
                g_texts.add(s.text.strip()[:50])
            else:
                g_shapes += 1

        for s in prs_s.slides[si].shapes:
            if _is_nav_dot(s):
                continue
            if hasattr(s, "text") and s.text.strip():
                s_texts.add(s.text.strip()[:50])
            else:
                s_shapes += 1

        missing_texts = g_texts - s_texts
        extra_texts = s_texts - g_texts

        if missing_texts:
            for t in sorted(missing_texts)[:5]:
                issues.append(f"  [MISSING] Slide {si+1}: text '{t[:50]}'")
        if extra_texts:
            for t in sorted(extra_texts)[:5]:
                issues.append(f"  [EXTRA]   Slide {si+1}: text '{t[:50]}'")

        shape_diff = s_shapes - g_shapes
        if abs(shape_diff) > 1:
            direction = "extra" if shape_diff > 0 else "missing"
            issues.append(f"  [SHAPES]  Slide {si+1}: {abs(shape_diff)} {direction} shapes (golden={g_shapes}, sandbox={s_shapes})")

    return issues


# ─── 6. Color diff analysis ─────────────────────────────────────────────

def check_color_diffs() -> List[str]:
    """Find color mismatches between golden and sandbox."""
    prs_g = Presentation(GOLDEN)
    prs_s = Presentation(SANDBOX)
    num_slides = min(len(prs_g.slides), len(prs_s.slides))

    issues = []
    for si in range(num_slides):
        g_map = {}
        s_map = {}

        for s in prs_g.slides[si].shapes:
            if _is_nav_dot(s):
                continue
            if hasattr(s, "text") and s.text.strip():
                g_map[s.text.strip()[:40]] = _get_fill_color(s)

        for s in prs_s.slides[si].shapes:
            if _is_nav_dot(s):
                continue
            if hasattr(s, "text") and s.text.strip():
                s_map[s.text.strip()[:40]] = _get_fill_color(s)

        for key in g_map:
            if key in s_map and g_map[key] and s_map[key] and g_map[key] != s_map[key]:
                issues.append(
                    f"  [COLOR] Slide {si+1}: '{key[:40]}' golden={g_map[key]} sandbox={s_map[key]}"
                )

    return issues


# ─── Main ──────────────────────────────────────────────────────────────

def main():
    print("=" * 80)
    print("RIGOROUS MULTI-DIMENSIONAL EVALUATION")
    print(f"  Golden:  {GOLDEN}")
    print(f"  Sandbox: {SANDBOX}")
    print("=" * 80)

    # 1. Visual comprehensive
    print("\n" + "=" * 80)
    print("DIMENSION 1: Visual Comprehensive (compare-visual-comprehensive.py)")
    print("=" * 80)
    print(run_visual_comprehensive())

    # 2. Overflow check
    print("\n" + "=" * 80)
    print("DIMENSION 2: Text Overflow Check")
    print("=" * 80)
    golden_overflows = check_overflow(GOLDEN, "GOLDEN")
    sandbox_overflows = check_overflow(SANDBOX, "SANDBOX")
    all_overflows = golden_overflows + sandbox_overflows
    if all_overflows:
        for issue in all_overflows:
            print(issue)
    else:
        print("  ✓ No overflow issues detected")
    print(f"\n  Summary: {len(sandbox_overflows)} sandbox overflow(s), {len(golden_overflows)} golden overflow(s)")

    # 3. Overlap check
    print("\n" + "=" * 80)
    print("DIMENSION 3: Element Overlap Check")
    print("=" * 80)
    sandbox_overlaps = check_overlap(SANDBOX, "SANDBOX")
    golden_overlaps = check_overlap(GOLDEN, "GOLDEN")
    all_overlaps = golden_overlaps + sandbox_overlaps
    if all_overlaps:
        for issue in all_overlaps:
            print(issue)
    else:
        print("  ✓ No overlap issues detected")
    print(f"\n  Summary: {len(sandbox_overlaps)} sandbox overlap(s), {len(golden_overlaps)} golden overlap(s)")

    # 4. Position drift
    print("\n" + "=" * 80)
    print("DIMENSION 4: Position Drift Analysis")
    print("=" * 80)
    drift_issues = check_position_drift()
    for issue in drift_issues:
        print(issue)

    # 5. Element gaps
    print("\n" + "=" * 80)
    print("DIMENSION 5: Element Gap Analysis")
    print("=" * 80)
    gap_issues = check_element_gaps()
    if gap_issues:
        for issue in gap_issues:
            print(issue)
    else:
        print("  ✓ No element gaps detected")

    # 6. Color diffs
    print("\n" + "=" * 80)
    print("DIMENSION 6: Color Difference Analysis")
    print("=" * 80)
    color_issues = check_color_diffs()
    if color_issues:
        for issue in color_issues:
            print(issue)
    else:
        print("  ✓ No color differences detected")

    # Final summary
    print("\n" + "=" * 80)
    print("FINAL SUMMARY")
    print("=" * 80)
    total_issues = len(sandbox_overflows) + len(sandbox_overlaps) + len(gap_issues) + len(color_issues)
    print(f"  Overflow issues:     {len(sandbox_overflows)}")
    print(f"  Overlap issues:      {len(sandbox_overlaps)}")
    print(f"  Position drift:      see above")
    print(f"  Element gaps:        {len(gap_issues)}")
    print(f"  Color differences:   {len(color_issues)}")
    print(f"  Total actionable:    {total_issues}")
    print("=" * 80)

    return 0


if __name__ == "__main__":
    sys.exit(main())
