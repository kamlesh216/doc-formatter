# Thesis Formatter

A production-quality Python tool that automatically formats a thesis/dissertation DOCX document — fixing spacing, fonts, images, tables, headings, and captions — **without ever modifying academic content**.

## Quick Start

```powershell
pip install -r requirements.txt

python format_thesis.py my_thesis.docx
# → produces my_thesis_formatted.docx

python format_thesis.py my_thesis.docx output.docx --report --analysis
```

## Architecture

```
thesis_formatter/
├── __init__.py
├── config.py       ← ALL formatting constants (edit here to change style)
├── analyzer.py     ← Classifies every paragraph into a type
├── formatter.py    ← Orchestrator — runs all modules in order
├── page_layout.py  ← Page size, margins, orientation
├── headings.py     ← Chapter/H1/H2/H3 formatting, page-breaks
├── paragraphs.py   ← Body text, bibliography (run-safe, never clear())
├── lists.py        ← Bullet and numbered lists
├── tables.py       ← Table borders, width, centering, header repeat
├── captions.py     ← Table/figure captions, keep_with_next
├── images.py       ← Image centering, width-capping, aspect ratio
├── equations.py    ← Equation preservation and centering
├── validator.py    ← Post-save structural integrity check
└── cli.py          ← argparse CLI
format_thesis.py    ← Entry point
```

## Formatting Rules (Arial body font)

| Element         | Rule |
|-----------------|------|
| Body font       | Arial 12 pt |
| Justification   | Justified |
| Line spacing    | 1.15 multiple |
| Space after     | 6 pt |
| Chapter heading | Centered · 16 pt Bold · New page before |
| Heading 1       | Left · 14 pt Bold |
| Heading 2       | Left · 12 pt Bold |
| Heading 3       | Left · 12 pt Bold Italic |
| Table captions  | Centered · 11 pt · Above table · keep_with_next |
| Figure captions | Centered · 11 pt · Below image · keep_with_next |
| Images          | Centered · max 6 in wide · aspect ratio preserved |
| Tables          | Centered · Table Grid borders · header row repeat |
| Page margins    | Top 1 in · Bottom 1 in · Left 1.25 in · Right 1 in |

## Configuration

Edit `thesis_formatter/config.py` to change any rule:

```python
BODY_FONT       = "Arial"      # ← change body font here
BODY_SIZE       = Pt(12)
MAX_IMAGE_WIDTH = PRINTABLE_WIDTH   # cap image width
```

## Safety Guarantees

- **Never overwrites** the original input file
- **Never calls** `paragraph.clear()` on body paragraphs (preserves bold, italic, superscript, subscript, hyperlinks, fields)
- **Never changes** table/figure content
- **Never rewrites** bibliography entries
- **Validates** output on reload before reporting success

## CLI Options

| Flag | Purpose |
|------|---------|
| `--report` | Print formatting summary after completion |
| `--analysis` | Print document analysis before formatting |
| `--no-validate` | Skip post-save validation |
| `-v` / `--verbose` | Enable debug logging |

## Running Tests

```powershell
python -m pytest tests/ -v
# or
python tests/test_formatter.py
```

## Requirements

- Python 3.9+
- `python-docx >= 1.1.0`
- `lxml >= 4.9.0`

## Detected types

The analyzer classifies every paragraph as one of:

`BODY | CHAPTER | HEADING1 | HEADING2 | HEADING3 | BULLET | NUMBERED | TABLE_CAPTION | FIGURE_CAPTION | EQUATION | IMAGE | FRONT_MATTER | BIBLIOGRAPHY | BLANK`
