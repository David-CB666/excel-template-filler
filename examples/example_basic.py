#!/usr/bin/env python3
"""Example: Basic template fill"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from template_filler import TemplateFiller


def main():
    # Example file paths (relative to examples/ directory)
    data_source = str(Path(__file__).parent / "data" / "sample_data.xlsx")
    template = str(Path(__file__).parent / "templates" / "sample_template.xlsx")

    if not Path(data_source).exists():
        print(f"[ERROR] Data source not found: {data_source}")
        return

    if not Path(template).exists():
        print(f"[ERROR] Template not found: {template}")
        return

    # Initialize filler (auto-detects engine)
    filler = TemplateFiller(template=template)

    # 1. Scan placeholders
    print("=== Scanning Placeholders ===")
    placeholders = filler.scan_placeholders()
    print(f"Found: {placeholders}")

    # 2. Load data with field mapping
    print("\n=== Loading Data ===")
    filler.load_data(
        data_source=data_source,
        field_map={
            "{ID}": "ID",
            "{Name}": "Name",
            "{Brand}": "Brand",
            "{Qty}": "Qty",
            "{Spec}": "Spec"
        }
    )
    print(f"Data loaded: {len(filler.data_list)} rows")

    # 3. Fill and export
    print("\n=== Filling & Exporting ===")
    output_dir = str(Path(__file__).parent / "output")
    output_files = filler.fill_and_export(
        start_row=2,
        end_row=11,
        export_format="excel",
        output_dir=output_dir,
        empty_handling="clear"
    )

    print(f"\nGenerated {len(output_files)} files:")
    for f in output_files:
        print(f"  - {f}")

    filler.close()
    print("\n[OK] Done")


if __name__ == "__main__":
    main()
