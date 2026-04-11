---
name: kai-export-ppt-lite
description: Use when exporting HTML presentations to editable PPTX in sandbox environments where Playwright cannot be installed. Pure Python: bs4 + python-pptx + Pillow. Keywords: export pptx sandbox, html to pptx no browser, editable powerpoint from html.
emoji: 📄
---

# kai-export-ppt-lite

Export HTML presentations to **editable PPTX** without a browser. Pure Python — no Playwright, no Chrome, no Node.js.

## Dependencies

| Package | Install |
|---------|---------|
| `beautifulsoup4` | `pip install beautifulsoup4` |
| `lxml` | `pip install lxml` |
| `python-pptx` | `pip install python-pptx` |
| `Pillow` | `pip install Pillow` |

## Usage

```bash
python3 scripts/export-sandbox-pptx.py <file.html> [output.pptx] [--width 1440] [--height 810]
```

### Options

| Flag | Description |
|------|-------------|
| `--width N` | Slide width in pixels (default: 1440) |
| `--height N` | Slide height in pixels (default: 810) |
| `--no-chrome` | Skip page counter and nav dots |

## What it extracts

| HTML Element | PPTX Output |
|-------------|-------------|
| `h1-h6, p, li, span` | Editable text boxes |
| `div` with background/border | Rounded rectangles with fill |
| `table, tr, td, th` | Cell rectangles with text frames |
| `img` (http/file/data-uri) | Embedded pictures |
| `svg` | Placeholder rectangle |

## Limitations (no browser)

- CSS variables: only `:root` static substitution
- Layout: simulated flex-column, not pixel-perfect
- Computed styles: unavailable, relies on inline style + `<style>` blocks
- JavaScript: not executed, so IntersectionObserver-gated content won't appear

## Comparison

| | kai-html-export (full) | kai-export-ppt-lite |
|---|---|---|
| Runtime | Playwright + Chrome | Pure Python |
| Sandbox compatible | ❌ | ✅ |
| Output quality | Pixel-perfect | Semantic layout |
| Editable text | ✅ | ✅ |
| Image mode PPTX | ✅ (screenshot) | ❌ |
| Dependencies | playwright, python-pptx | bs4, lxml, python-pptx, Pillow |
