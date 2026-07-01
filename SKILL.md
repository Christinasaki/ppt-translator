---
name: ppt-translator
description: >
  Translate PowerPoint (.pptx) presentations between Chinese and English while
  preserving layout, formatting, and visual design. Supports terminology glossaries
  (Excel/JSON), SmartArt text, grouped shapes, and auto font-size fitting.
  Use this skill when the user wants to translate a PPT/PPTX file, localize a
  presentation, or batch-translate slides with a custom glossary.
agent_created: true
---

# PPT Translator Skill

Translate PowerPoint presentations between Chinese (zh) and English (en) while
preserving the original layout, formatting, and visual design.

## Prerequisites

- Python >= 3.8
- Required packages: `python-pptx`, `Pillow`, `openpyxl`, `lxml`

Install dependencies if not already available:

```bash
pip install python-pptx Pillow openpyxl lxml
```

## Quick Start

Run the standalone script directly:

```bash
python ppt_translate_standalone.py <input.pptx> [-o output.pptx] [-d zh2en|en2zh] [-t terms.xlsx]
```

### Parameters

| Flag | Default | Description |
|------|---------|-------------|
| `input` | (required) | Input `.pptx` file path |
| `-o, --output` | `{name}_EN.pptx` or `_ZH.pptx` | Output file path |
| `-d, --direction` | `zh2en` | Translation direction |
| `-t, --terms` | none | Terminology file (`.xlsx` or `.json`) |
| `--font` | 保留原字体 | Target font family for translated text (omit to preserve original) |
| `--no-title-case` | off | Disable automatic Title Case for zh→en |

### Examples

```bash
# Basic: Chinese → English
python ppt_translate_standalone.py report.pptx -o report_en.pptx

# With terminology glossary
python ppt_translate_standalone.py report.pptx -t glossary.xlsx -o report_en.pptx

# English → Chinese
python ppt_translate_standalone.py deck_en.pptx -d en2zh -o deck_zh.pptx

# Custom font, no Title Case
python ppt_translate_standalone.py slides.pptx --font "Microsoft YaHei" --no-title-case
```

## Terminology Glossary

The glossary ensures domain-specific terms are translated consistently.

### Excel format (.xlsx)

Two columns, no header required:

| Column A (Source) | Column B (Target) |
|--------------------|-------------------|
| 人工智能 | Artificial Intelligence |
| 机器学习 | Machine Learning |
| 深度学习 | Deep Learning |

### JSON format (.json)

```json
{
  "人工智能": "Artificial Intelligence",
  "机器学习": "Machine Learning"
}
```

Or array format:

```json
[
  {"source": "人工智能", "target": "Artificial Intelligence"},
  {"source": "机器学习", "target": "Machine Learning"}
]
```

## Key Capabilities

### Layout Preservation
- Text frames, tables, SmartArt, and grouped shapes are all processed in place
- No slide rearrangement; original layout is maintained exactly

### Smart Font Fitting
- Auto-detects when translated text would overflow its text box
- Progressively reduces font size (down to 5pt) to fit
- Uses actual font metrics (via PIL) for accurate width measurement

### Title Case (zh→en)
- Automatically applies English Title Case rules to headings and short text
- Preserves acronyms (AI, API, KPI, etc.) and respects small-word exceptions
- Disable with `--no-title-case` if not desired

### Grouped & Nested Shapes
- Correctly handles coordinate transforms for nested group shapes
- Accumulates scale factors through group hierarchy

### SmartArt
- Extracts and translates text inside SmartArt diagrams
- Preserves SmartArt structure and relationships

## Project Structure

```
ppt-translator/
├── SKILL.md                          # This file
├── ppt_translate_standalone.py       # Standalone script (all-in-one)
├── README.md                         # Project documentation
├── upload_to_github.js               # GitHub upload helper
└── ppt-translator-v1.0.0.zip         # Release package
```

## Notes

- The standalone script (`ppt_translate_standalone.py`) is self-contained with no internal package imports — it can be copied and used independently
- For large presentations (100+ slides), translation may take a few minutes due to font metric calculations
- The tool works with `.pptx` files only (not legacy `.ppt` format)
- Encrypted or password-protected presentations are not supported

## Known Bug Fixes

### `word_wrap` AttributeError (Fixed 2026-06-29)
- **Issue**: `_Paragraph` object has no attribute `word_wrap` — `word_wrap` belongs to `TextFrame`, not `Paragraph`
- **Fix**: Pass `word_wrap` as parameter to `_process_paragraph`, read from `TextFrame` via `getattr(tf, 'word_wrap', True)`
