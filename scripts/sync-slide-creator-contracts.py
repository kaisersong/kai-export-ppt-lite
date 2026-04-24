#!/usr/bin/env python3
"""Sync slide-creator preset contracts into this repo.

Development-time only. This script reads stable reference/demo material from a
local slide-creator checkout, emits versioned preset contracts, and writes a
manifest that records the upstream version used for the sync.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional

from bs4 import BeautifulSoup


REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SLIDE_CREATOR_ROOT = REPO_ROOT.parent / "slide-creator"
CONTRACTS_ROOT = REPO_ROOT / "contracts" / "slide_creator"
PRESETS_ROOT = CONTRACTS_ROOT / "presets"

COMMON_CHROME_SELECTORS = [
    ".progress-bar",
    ".nav-dots",
    ".edit-hotzone",
    ".edit-toggle",
    "#notes-panel",
    "#present-btn",
    "#present-counter",
]


PRESET_SPECS: Dict[str, Dict[str, Any]] = {
    "blue-sky": {
        "slug": "blue-sky",
        "preset": "Blue Sky",
        "family": "bright-card-stack",
        "reference_refs": [
            "references/blue-sky-starter.html",
            "references/html-template.md",
        ],
        "demo_refs": [
            "demos/blue-sky-zh.html",
            "demos/blue-sky-en.html",
        ],
        "decorative_layers": [],
        "component_selectors": {
            "pill": [".pill"],
            "card": [".g", ".layer"],
            "cta_row": [".co", ".cmd", ".info"],
            "table_card": [".ctable"],
        },
        "component_slot_models": {
            "pill": {
                "layout": "inline_pill",
                "slots": ["text"],
            },
            "card": {
                "layout": "stack_card",
                "slots": ["eyebrow", "title", "body", "table_or_list", "footer"],
            },
            "cta_row": {
                "layout": "inline_command_row",
                "slots": ["command", "separator", "link"],
            },
        },
        "layout_variations": [
            "hero-stats",
            "pill-headers",
            "layer-cards",
            "closing-cta",
        ],
        "text_expectations": {"tabular_numbers": False},
    },
    "enterprise-dark": {
        "slug": "enterprise-dark",
        "preset": "Enterprise Dark",
        "family": "consulting-dark",
        "reference_refs": [
            "references/enterprise-dark.md",
            "references/html-template.md",
        ],
        "demo_refs": [
            "demos/enterprise-dark-zh.html",
            "demos/enterprise-dark-en.html",
        ],
        "decorative_layers": [
            {
                "selector": "body::before",
                "kind": "density-grid",
                "export_strategy": "background-layer",
            }
        ],
        "component_selectors": {
            "split_layout": [".ent-split"],
            "split_label": [".ent-split-label"],
            "metric_grid": [".ent-kpi-grid"],
            "metric_card": [".ent-kpi-card"],
            "data_table": [".ent-table"],
            "pill": [".ent-pill"],
            "progress_track": [".ent-progress-track"],
            "progress_fill": [".ent-progress-fill"],
            "trend": [".ent-trend"],
            "cta_pill": [".cta-pill"],
        },
        "component_slot_models": {
            "split_layout": {
                "layout": "grid_two_column",
                "slots": ["label_rail", "content_rail"],
                "track_template": "clamp(160px,22%,240px) 1fr",
            },
            "metric_card": {
                "layout": "vertical_card",
                "slots": ["metric", "title_or_label", "body", "progress_or_trend"],
                "allow_inline_variant": True,
            },
            "data_table": {
                "layout": "open_table",
                "slots": ["header", "body"],
            },
            "cta_pill": {
                "layout": "cta_pill",
                "slots": ["command"],
            },
        },
        "layout_variations": [
            "hero-stats",
            "consulting-split",
            "kpi-dashboard",
            "data-table",
            "installation-cards",
            "closing-cta",
        ],
        "text_expectations": {"tabular_numbers": True},
    },
    "swiss-modern": {
        "slug": "swiss-modern",
        "preset": "Swiss Modern",
        "family": "editorial-swiss",
        "reference_refs": [
            "references/swiss-modern.md",
            "references/html-template.md",
        ],
        "demo_refs": [
            "demos/swiss-modern-zh.html",
            "demos/swiss-modern-en.html",
        ],
        "decorative_layers": [],
        "component_selectors": {
            "editorial_grid": [".grid", ".editorial-grid"],
            "section_label": [".swiss-label", ".label", ".eyebrow"],
            "card": [".swiss-card", ".feat-card", ".pain-item"],
            "rule": [".swiss-rule", ".swiss-rule-thin"],
            "terminal_line": [".terminal-line"],
        },
        "component_slot_models": {
            "card": {
                "layout": "editorial_card",
                "slots": ["label", "title", "body", "rule_or_stat"],
            },
            "terminal_line": {
                "layout": "inline_command_row",
                "slots": ["command", "label", "link"],
            },
        },
        "layout_variations": [
            "editorial-hero",
            "swiss-grid",
            "section-divider",
        ],
        "text_expectations": {"tabular_numbers": True},
    },
    "data-story": {
        "slug": "data-story",
        "preset": "Data Story",
        "family": "data-visual-narrative",
        "reference_refs": [
            "references/data-story.md",
            "references/html-template.md",
        ],
        "demo_refs": [
            "demos/data-story-zh.html",
            "demos/data-story-en.html",
        ],
        "decorative_layers": [
            {
                "selector": ".slide::before",
                "kind": "data-grid",
                "export_strategy": "background-layer",
            },
            {
                "selector": "body::before",
                "kind": "data-grid",
                "export_strategy": "background-layer",
            },
        ],
        "component_selectors": {
            "split_layout": [".ds-split-layout"],
            "left_panel": [".left-panel"],
            "metric_grid": [".ds-kpi-grid"],
            "metric_card": [".ds-kpi-card"],
            "hero_metric": [".ds-kpi"],
            "metric_label": [".ds-kpi-label"],
            "trend": [".ds-trend"],
            "insight": [".ds-insight"],
            "style_card": [".style-card"],
            "solution_card": [".sol-card"],
            "solution_icon": [".sol-icon"],
            "chart_svg": [".ds-chart-svg"],
            "chart_line": [".ds-line"],
            "chart_axis": [".chart-axis", ".ds-axis"],
            "chart_grid": [".chart-grid", ".ds-grid"],
            "matrix": [".ds-matrix"],
            "matrix_cell": [".ds-matrix-cell"],
            "feature_grid": [".feat-grid"],
            "feature_card": [".feat-card"],
            "feature_stat": [".feat-stat"],
            "steps": [".steps"],
            "step_item": [".step-item"],
            "step_num": [".step-num"],
            "install_row": [".install-row"],
            "install_label": [".install-label"],
            "install_cmd": [".install-cmd"],
        },
        "component_slot_models": {
            "split_layout": {
                "layout": "grid_two_column",
                "slots": ["metric_stack", "chart"],
                "track_template": "1fr 1.5fr",
            },
            "metric_card": {
                "layout": "vertical_card",
                "slots": ["metric", "label", "trend"],
                "minimum_height_in": 1.10,
                "gaps": {"after_metric": 0.05, "after_label": 0.05},
                "metric_single_line": True,
                "bottom_anchor_last_slot": True,
                "metric_vertical_align": "center_remaining",
                "metric_max_height_ratio": 0.80,
            },
            "style_card": {
                "layout": "vertical_card",
                "slots": ["preview", "title", "body", "trend"],
                "minimum_height_in": 1.28,
                "gaps": {"after_metric": 0.10, "after_label": 0.06},
                "stretch_first_slot": True,
                "bottom_anchor_last_slot": False,
            },
            "solution_card": {
                "layout": "vertical_card",
                "slots": ["icon", "title", "body"],
                "minimum_height_in": 1.02,
                "gaps": {"after_metric": 0.08, "after_label": 0.06},
                "bottom_anchor_last_slot": False,
            },
            "feature_card": {
                "layout": "vertical_card",
                "slots": ["metric", "title", "body"],
                "minimum_height_in": 1.06,
                "gaps": {"after_metric": 0.10, "after_label": 0.06},
                "metric_single_line": True,
                "bottom_anchor_last_slot": False,
                "metric_vertical_align": "center_remaining",
                "metric_max_height_ratio": 0.58,
            },
            "install_row": {
                "layout": "split_rail",
                "slots": ["label", "command"],
                "keep_command_single_line": True,
                "minimum_height_in": 0.44,
                "label_min_width_px": 120,
                "gap_px": 16,
            },
            "chart_svg": {
                "layout": "svg_chart",
                "slots": ["axes", "grid", "series", "labels"],
            },
            "matrix_cell": {
                "layout": "vertical_card",
                "slots": ["label", "metric", "body"],
            },
        },
        "layout_variations": [
            "hero-number",
            "kpi-row-chart",
            "chart-layout",
            "comparison-matrix",
            "feature-grid",
            "install-rows",
            "closing-kpi",
        ],
        "text_expectations": {"tabular_numbers": True},
    },
}


def _read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _write_json(path: Path, payload: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


def _relative_paths(root: Path, refs: List[str]) -> List[str]:
    return [str((root / ref).resolve().relative_to(root.resolve())) for ref in refs]


def _detect_slide_creator_version(html: str) -> Optional[str]:
    matches = re.findall(r"kai-slide-creator v(\d+\.\d+\.\d+)", html)
    return matches[-1] if matches else None


def _collect_observed_classes(soup: BeautifulSoup) -> List[str]:
    classes = set()
    for node in soup.find_all(True):
        for class_name in node.get("class") or []:
            if (
                class_name.startswith("ds-")
                or class_name.startswith("ent-")
                or class_name.startswith("swiss-")
                or class_name in {"feat-card", "install-row", "install-label", "install-cmd", "terminal-line"}
                or class_name.startswith("chart-")
            ):
                classes.add(class_name)
    return sorted(classes)


def _collect_demo_metadata(slide_creator_root: Path, demo_refs: List[str]) -> Dict[str, Any]:
    discovered_version: Optional[str] = None
    discovered_preset: Optional[str] = None
    observed_classes: set[str] = set()

    for ref in demo_refs:
        demo_path = slide_creator_root / ref
        if not demo_path.exists():
            continue
        html = _read_text(demo_path)
        soup = BeautifulSoup(html, "lxml")
        body = soup.find("body")
        if body and body.get("data-preset") and not discovered_preset:
            discovered_preset = body.get("data-preset", "").strip()
        version = _detect_slide_creator_version(html)
        if version and not discovered_version:
            discovered_version = version
        observed_classes.update(_collect_observed_classes(soup))

    return {
        "producer_version_tested": discovered_version,
        "preset": discovered_preset,
        "observed_component_classes": sorted(observed_classes),
    }


def _get_upstream_commit(slide_creator_root: Path) -> Optional[str]:
    try:
        result = subprocess.run(
            ["git", "-C", str(slide_creator_root), "rev-parse", "HEAD"],
            check=True,
            capture_output=True,
            text=True,
        )
    except Exception:
        return None
    return result.stdout.strip() or None


def _producer_version_range(version: Optional[str]) -> str:
    if not version:
        return ">=0.0.0 <3.0.0"
    return f">={version} <3.0.0"


def build_contract(
    slide_creator_root: Path,
    spec: Dict[str, Any],
    generated_at: str,
    upstream_commit: Optional[str],
) -> Dict[str, Any]:
    demo_meta = _collect_demo_metadata(slide_creator_root, spec["demo_refs"])
    producer_version = demo_meta.get("producer_version_tested")
    source_version = f"slide-creator@{upstream_commit[:12]}" if upstream_commit else (
        f"slide-creator@v{producer_version}" if producer_version else "slide-creator@unknown"
    )

    return {
        "producer": "slide-creator",
        "contract_id": f"slide-creator/{spec['slug']}",
        "preset": demo_meta.get("preset") or spec["preset"],
        "family": spec["family"],
        "contract_version": "1.0.0",
        "producer_version_range": _producer_version_range(producer_version),
        "producer_version_tested": producer_version or "unknown",
        "contract_source_version": source_version,
        "upstream_commit": upstream_commit or "",
        "generated_at": generated_at,
        "source_refs": _relative_paths(slide_creator_root, spec["reference_refs"]),
        "demo_refs": _relative_paths(slide_creator_root, spec["demo_refs"]),
        "runtime_chrome_selectors": list(COMMON_CHROME_SELECTORS),
        "decorative_layers": spec["decorative_layers"],
        "component_selectors": spec["component_selectors"],
        "component_slot_models": spec["component_slot_models"],
        "layout_variations": spec["layout_variations"],
        "text_expectations": spec["text_expectations"],
        "producer_detection": {
            "body_attrs": {
                "data-preset": spec["preset"],
                "data-export-progress": "true|false",
            },
            "meta_generator_contains": "kai-slide-creator",
            "watermark_contains": "By kai-slide-creator v",
        },
        "observed_component_classes": demo_meta.get("observed_component_classes", []),
    }


def build_manifest(
    slide_creator_root: Path,
    generated_at: str,
    upstream_commit: Optional[str],
) -> Dict[str, Any]:
    presets = []
    for slug, spec in PRESET_SPECS.items():
        demo_meta = _collect_demo_metadata(slide_creator_root, spec["demo_refs"])
        presets.append(
            {
                "slug": slug,
                "preset": demo_meta.get("preset") or spec["preset"],
                "contract_id": f"slide-creator/{slug}",
                "contract_version": "1.0.0",
                "family": spec["family"],
                "producer_version_tested": demo_meta.get("producer_version_tested") or "unknown",
                "contract_path": f"presets/{slug}.json",
                "source_refs": _relative_paths(slide_creator_root, spec["reference_refs"]),
                "demo_refs": _relative_paths(slide_creator_root, spec["demo_refs"]),
            }
        )

    return {
        "producer": "slide-creator",
        "manifest_version": "1.0.0",
        "generated_at": generated_at,
        "upstream_repo": "slide-creator",
        "upstream_commit": upstream_commit or "",
        "runtime_chrome_selectors": list(COMMON_CHROME_SELECTORS),
        "presets": presets,
    }


def sync_contracts(slide_creator_root: Path) -> Dict[str, Any]:
    generated_at = date.today().isoformat()
    upstream_commit = _get_upstream_commit(slide_creator_root)

    PRESETS_ROOT.mkdir(parents=True, exist_ok=True)

    manifest = build_manifest(slide_creator_root, generated_at, upstream_commit)
    for spec in PRESET_SPECS.values():
        contract = build_contract(slide_creator_root, spec, generated_at, upstream_commit)
        _write_json(PRESETS_ROOT / f"{spec['slug']}.json", contract)

    _write_json(CONTRACTS_ROOT / "manifest.json", manifest)
    return manifest


def main() -> int:
    parser = argparse.ArgumentParser(description="Sync slide-creator preset contracts into this repo.")
    parser.add_argument(
        "--slide-creator-root",
        type=Path,
        default=DEFAULT_SLIDE_CREATOR_ROOT,
        help="Path to the local slide-creator checkout.",
    )
    args = parser.parse_args()

    root = args.slide_creator_root.resolve()
    if not root.exists():
        raise SystemExit(f"slide-creator root not found: {root}")

    manifest = sync_contracts(root)
    print(
        f"Synced {len(manifest['presets'])} slide-creator preset contracts "
        f"from {root} @ {manifest.get('upstream_commit') or 'unknown'}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
