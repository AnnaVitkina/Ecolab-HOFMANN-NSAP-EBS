"""
HOFMANN Ecolab — end-to-end rate pipeline.

Steps:
  1. Select Excel files and sheets from input/
  2. Combine selected tabs into processing/
  3. Build shipment/cost matrix and save to output/

Colab usage:
  exec(open('/content/Ecolab-HOFMANN-NSAP-EBS/run_pipeline.py').read())
"""

from __future__ import annotations

import sys
from pathlib import Path


def _ensure_code_on_path() -> None:
    """Make project modules importable (local run, Colab exec, or notebook cwd)."""
    candidates: list[Path] = []
    file_path = globals().get("__file__")
    if file_path:
        candidates.append(Path(file_path).resolve().parent)
    candidates.append(Path("/content/Ecolab-HOFMANN-NSAP-EBS"))
    candidates.append(Path.cwd())

    seen: set[str] = set()
    for path in candidates:
        path = path.resolve()
        if not (path / "build_matrix.py").is_file():
            continue
        path_str = str(path)
        if path_str in seen:
            continue
        seen.add(path_str)
        if path_str not in sys.path:
            sys.path.insert(0, path_str)


_ensure_code_on_path()

from build_matrix import run_matrix_build
from process_tabs import run_combine_tabs


def main() -> None:
    print("=== Step 1/2: Select and combine tabs ===")
    combined_path = run_combine_tabs()

    print("\n=== Step 2/2: Build rate matrix ===")
    matrix_path = run_matrix_build(combined_path)

    print("\n=== Pipeline complete ===")
    print(f"  Combined workbook: {combined_path}")
    print(f"  Matrix output:     {matrix_path}")


if __name__ == "__main__":
    main()
