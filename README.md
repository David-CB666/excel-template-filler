# Excel Template Filler

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)

> Dual-engine architecture: openpyxl + ZIP | Batch fill Excel templates with perfect image/print-setting preservation

---

## Why This Tool?

openpyxl's `copy_worksheet()` loses images, print settings, and binary resources. This tool solves that with a **ZIP engine** that operates on XLSX internals directly — preserving everything openpyxl breaks.

**Auto-selects the best engine** based on template characteristics. Zero config needed.

---

## Quick Start

```bash
# Clone
git clone https://github.com/David-CB666/excel-template-filler.git
cd excel-template-filler

# Install dependency
pip install -r requirements.txt

# Run example
cd examples
python example_basic.py
```

## Basic Usage

```python
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from template_filler import TemplateFiller

# Initialize (auto-detects engine)
filler = TemplateFiller(
    data_source="data/sample_data.xlsx",
    template="templates/sample_template.xlsx"
)

# Check which engine is selected
print(f"Engine: {'ZIP' if filler.has_images() else 'openpyxl'}")

# Scan placeholders
placeholders = filler.scan_placeholders()

# Fill and export
output_files = filler.fill_and_export(
    field_map={
        "{ID}": "ID",
        "{Name}": "Name",
        "{Brand}": "Brand",
        "{Qty}": "Qty"
    },
    output_dir="./output"
)
```

---

## Dual-Engine Architecture

| Engine | Best For | Key Advantage |
|:---|:---|:---|
| **openpyxl** | Templates without images | Clean API, easy to maintain |
| **ZIP** | Templates with images | Preserves images, print settings, binary resources |

The `TemplateFiller` auto-detects which engine to use based on template content.

---

## Features

| Feature | Description |
|:---|:---|
| **Auto engine selection** | Detects template type, picks optimal engine |
| **Image preservation** | ZIP engine keeps images/print settings intact |
| **Placeholder scanning** | Finds `{field}` or `{{field}}` patterns automatically |
| **Batch export** | Excel or PDF output (PDF requires pywin32 on Windows) |
| **BQ page merging** | Merge application PDFs with Bill of Quantities pages |
| **Auto-linking** | Match filenames to cells and create hyperlinks |
| **Data validation** | Verify data source completeness before filling |

---

## Directory Structure

```
excel-template-filler/
├── src/
│   ├── template_filler.py       # Unified entry (auto engine selection)
│   ├── engines/
│   │   ├── base_engine.py       # Engine interface
│   │   ├── openpyxl_engine.py   # openpyxl engine (no images)
│   │   └── zip_engine.py        # ZIP engine (images preserved)
│   ├── exporters/
│   │   └── bq_merger.py         # BQ page merger
│   ├── auto_linker.py           # Auto hyperlink creator
│   ├── file_grabber.py          # Filename grabber
│   └── utils.py                 # Utilities
├── examples/
│   ├── data/                    # Sample data files
│   ├── templates/               # Sample templates
│   ├── example_basic.py         # Basic fill example
│   ├── example_batch_pdf.py    # Batch PDF example
│   └── example_auto_link.py    # Auto-link example
├── pyproject.toml
├── requirements.txt
└── README.md
```

---

## Examples

### Scenario 1: Simple Template Fill

```python
filler = TemplateFiller("data.xlsx", "template.xlsx")
output = filler.fill_and_export(
    field_map={"{ID}": "ID", "{Name}": "Name"},
    output_dir="./output"
)
```

### Scenario 2: Full Pipeline (Fill + BQ Merge)

```python
# Step 1: Generate filled PDFs
filler = TemplateFiller("data.xlsx", "template.xlsx")
filler.fill_and_export(
    export_format="pdf", output_dir="./pdfs"
)

# Step 2: Merge with BQ pages
from exporters.bq_merger import BQMerger
merger = BQMerger()
merger.load_bq_pdf("bq_reference.pdf")
merger.load_zongbiao("data.xlsx")
merger.match_bq_pages()
merger.merge_pdfs("./pdfs", "./final_output")
```

---

## Requirements

| Requirement | Version |
|:---|:---|
| Python | >= 3.10 |
| openpyxl | >= 3.1.0 |

**Optional:** pywin32 (for PDF export on Windows)

---

## Limitations

1. Placeholders must be in shared strings (normal cell text)
2. Mixed placeholder + text in a single cell is not supported
3. BQ merge requires PDFs with text layers (not scanned images)

---

## Changelog

### v2.1.0 (2026-06-02)
- Code refactoring and desensitization
- Removed internal project references
- Added pyproject.toml, .gitignore, example data

### v2.0.0 (2026-06-01)
- Dual-engine architecture (openpyxl + ZIP)
- Auto engine selection
- BQ page merging
- Unified API

### v1.0.0 (2026-06-01)
- Core: template filling, batch export, auto-linking
- PDF/Excel output support

---

## License

[MIT](https://opensource.org/licenses/MIT)
