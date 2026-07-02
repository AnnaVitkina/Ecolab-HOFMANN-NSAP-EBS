"""
Select Excel files from input/, pick sheets, and write one combined workbook to processing/.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

from project_paths import INPUT_DIR, PROCESSING_DIR, ensure_workspace_dirs

EXCEL_SUFFIXES = {".xlsx", ".xls", ".xlsm"}


@dataclass(frozen=True)
class SheetSelection:
    file_path: Path
    sheet_name: str

    @property
    def label(self) -> str:
        return f"{self.file_path.name} -> {self.sheet_name}"


def ensure_directories() -> None:
    ensure_workspace_dirs()


def list_input_files() -> list[Path]:
    files = [
        path
        for path in sorted(INPUT_DIR.iterdir())
        if path.is_file()
        and path.suffix.lower() in EXCEL_SUFFIXES
        and not path.name.startswith("~$")
    ]
    return files


def parse_selection(raw: str, max_index: int) -> list[int]:
    """Parse '1,3-5' into zero-based indices."""
    raw = raw.strip().lower()
    if raw in {"all", "*"}:
        return list(range(max_index))

    indices: set[int] = set()
    for part in re.split(r"\s*,\s*", raw):
        if not part:
            continue
        if "-" in part:
            start_s, end_s = part.split("-", 1)
            start = int(start_s) - 1
            end = int(end_s) - 1
            if start > end or start < 0 or end >= max_index:
                raise ValueError(f"Invalid range: {part}")
            indices.update(range(start, end + 1))
        else:
            idx = int(part) - 1
            if idx < 0 or idx >= max_index:
                raise ValueError(f"Invalid index: {part}")
            indices.add(idx)
    return sorted(indices)


def prompt_selection(
    title: str,
    items: list[str],
    allow_empty: bool = False,
) -> list[int]:
    if not items:
        return []

    print(f"\n{title}")
    for i, item in enumerate(items, start=1):
        print(f"  {i}. {item}")

    hint = "Enter numbers (e.g. 1,3 or 1-3) or 'all'"

    while True:
        raw = input(f"{hint}: ").strip()
        if not raw and allow_empty:
            return []
        if not raw:
            print("Please enter at least one choice.")
            continue
        try:
            chosen = parse_selection(raw, len(items))
            if chosen or allow_empty:
                return chosen
            print("Please enter at least one choice.")
        except ValueError as exc:
            print(f"Invalid input: {exc}")


def select_input_files(files: list[Path]) -> list[Path]:
    if not files:
        print(f"No Excel files found in: {INPUT_DIR}")
        sys.exit(1)

    print("\nAvailable input files:")
    for i, path in enumerate(files, start=1):
        print(f"  {i}. {path.name}")

    indices = prompt_selection(
        "Select file(s) to process (multiple allowed):",
        [f.name for f in files],
    )
    return [files[i] for i in indices]


def select_sheets(file_path: Path, sheet_names: list[str]) -> list[str]:
    indices = prompt_selection(
        f"Choose sheet(s) for {file_path.name}:",
        sheet_names,
    )
    return [sheet_names[i] for i in indices]


def gather_sheet_selections(files: list[Path]) -> list[SheetSelection]:
    selections: list[SheetSelection] = []

    for file_path in files:
        try:
            workbook = pd.ExcelFile(file_path)
        except Exception as exc:
            print(f"Skipping {file_path.name}: could not open file ({exc})")
            continue

        selected_sheets = select_sheets(file_path, workbook.sheet_names)
        if not selected_sheets:
            print(f"No sheets selected for {file_path.name}; skipping.")
            continue

        for sheet_name in selected_sheets:
            selections.append(SheetSelection(file_path, sheet_name))

    return selections


MAX_SHEET_NAME_LEN = 31


def sanitize_sheet_part(text: str) -> str:
    return re.sub(r"[\[\]:*?/\\]", "_", text).strip() or "Sheet"


def combine_file_and_sheet(file_stem: str, sheet_name: str, max_len: int = MAX_SHEET_NAME_LEN) -> str:
    file_part = sanitize_sheet_part(file_stem)
    sheet_part = sanitize_sheet_part(sheet_name)
    separator = "_"

    combined = f"{file_part}{separator}{sheet_part}"
    if len(combined) <= max_len:
        return combined

    # Excel allows max 31 chars: keep the tab name, shorten the file part.
    max_file_len = max_len - len(separator) - len(sheet_part)
    if max_file_len >= 1:
        return f"{file_part[:max_file_len]}{separator}{sheet_part}"

    return sheet_part[:max_len]


def unique_sheet_name(file_stem: str, sheet_name: str, used: set[str]) -> str:
    base = combine_file_and_sheet(file_stem, sheet_name)
    if base not in used:
        used.add(base)
        return base

    for n in range(2, 1000):
        suffix = f"_{n}"
        trimmed = combine_file_and_sheet(
            file_stem,
            sheet_name,
            max_len=MAX_SHEET_NAME_LEN - len(suffix),
        )
        candidate = trimmed + suffix
        if candidate not in used:
            used.add(candidate)
            return candidate

    raise RuntimeError(f"Could not create a unique sheet name for: {file_stem} / {sheet_name}")


def collect_selected_frames(
    selections: list[SheetSelection],
) -> list[tuple[str, pd.DataFrame]]:
    frames: list[tuple[str, pd.DataFrame]] = []
    used_names: set[str] = set()

    for selection in selections:
        try:
            df = pd.read_excel(selection.file_path, sheet_name=selection.sheet_name)
        except Exception as exc:
            print(f"Skipping {selection.label}: could not read sheet ({exc})")
            continue

        sheet_label = unique_sheet_name(selection.file_path.stem, selection.sheet_name, used_names)
        frames.append((sheet_label, df))
        print(f"  Loaded: {selection.label} -> output tab '{sheet_label}' ({len(df)} rows)")

    return frames


def save_combined_workbook(frames: list[tuple[str, pd.DataFrame]]) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = PROCESSING_DIR / f"combined_{timestamp}.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in frames:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  Wrote tab: {sheet_name}")

    return output_path


def run_combine_tabs() -> Path:
    """Interactive step: combine selected input tabs into processing/."""
    ensure_directories()
    files = list_input_files()
    selected_files = select_input_files(files)
    selections = gather_sheet_selections(selected_files)

    if not selections:
        print("\nNothing to save. No sheets were selected.")
        sys.exit(1)

    frames = collect_selected_frames(selections)

    if not frames:
        print("\nNothing to save. No sheets could be loaded.")
        sys.exit(1)

    output_path = save_combined_workbook(frames)
    print(f"\nSaved {len(frames)} sheet(s) to: {output_path}")
    return output_path


def main() -> None:
    run_combine_tabs()


if __name__ == "__main__":
    main()
