"""Extract accessorial costs from unused combined workbook tabs."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Any

import pandas as pd

DEFAULT_RATE_BY = "Per shipment"
DEFAULT_APPLY_IF = "Apply to all shipments"


@dataclass(frozen=True)
class AccessorialCostSpec:
    patterns: tuple[str, ...]
    cost_name: str
    rate_by: str | None = None
    apply_if: str | None = None
    exclude_patterns: tuple[str, ...] = ()


@dataclass
class AccessorialCostRow:
    cost_name: str
    cost_price: float
    rate_by: str
    apply_if: str


ACCESSORIAL_COST_SPECS: list[AccessorialCostSpec] = [
    AccessorialCostSpec(
        patterns=("senior", "spf"),
        cost_name="Transfer Service (Material/Senior)",
        rate_by="Material/Senior",
        apply_if="Service equals Transfer Service",
        exclude_patterns=("junior",),
    ),
    AccessorialCostSpec(
        patterns=("junior", "jpf"),
        cost_name="Transfer Service (Material/Junior)",
        rate_by="Material/Junior",
        apply_if="Service equals Transfer Service",
    ),
    AccessorialCostSpec(
        patterns=("mini", "mpf"),
        cost_name="Transfer Service (Material/Mini)",
        rate_by="Material/Mini",
        apply_if="Service equals Transfer Service",
        exclude_patterns=("micro", "mic"),
    ),
    AccessorialCostSpec(
        patterns=("micro", "mic"),
        cost_name="Transfer Service (Material/Micro)",
        rate_by="Material/Micro",
        apply_if="Service equals Transfer Service",
    ),
    AccessorialCostSpec(
        patterns=("nl nextday",),
        cost_name="Next Day fee",
    ),
    AccessorialCostSpec(
        patterns=("nd 12h",),
        cost_name="Next Day 12:00 fee",
    ),
    AccessorialCostSpec(
        patterns=("nd 10h",),
        cost_name="Next Day 10:00 fee",
    ),
    AccessorialCostSpec(
        patterns=("nd 8h",),
        cost_name="Next Day 8:00 fee",
    ),
    AccessorialCostSpec(
        patterns=("mitnahmestapler",),
        cost_name="Forklift surcharge",
    ),
]


def cell_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def parse_price(value: Any) -> float | None:
    if pd.isna(value):
        return None
    text = cell_text(value)
    if not text or text.lower() == "on request":
        return None
    try:
        return float(Decimal(str(value).replace(",", ".")))
    except (TypeError, ValueError, InvalidOperation, ArithmeticError):
        return None


def row_label(row: pd.Series) -> str:
    if row.empty:
        return ""
    return cell_text(row.iloc[0])


def row_price(row: pd.Series) -> float | None:
    for column_index in range(1, len(row)):
        price = parse_price(row.iloc[column_index])
        if price is not None:
            return price
    return None


def label_matches(label: str, spec: AccessorialCostSpec) -> bool:
    lowered = label.lower()
    if spec.exclude_patterns and any(token in lowered for token in spec.exclude_patterns):
        return False
    return any(pattern in lowered for pattern in spec.patterns)


def find_spec_price_in_sheet(raw: pd.DataFrame, spec: AccessorialCostSpec) -> float | None:
    if spec.patterns == ("mitnahmestapler",):
        return find_forklift_price(raw)

    for row_index in range(len(raw)):
        label = row_label(raw.iloc[row_index])
        if not label or not label_matches(label, spec):
            continue
        price = row_price(raw.iloc[row_index])
        if price is not None:
            return price
    return None


def find_forklift_price(raw: pd.DataFrame) -> float | None:
    in_section = False
    for row_index in range(len(raw)):
        label = row_label(raw.iloc[row_index])
        lowered = label.lower()
        if "mitnahmestapler" in lowered:
            in_section = True
            continue
        if not in_section:
            continue
        if not label:
            continue
        price = row_price(raw.iloc[row_index])
        if price is not None:
            return price
        if lowered.endswith(":"):
            continue
        break
    return None


def build_accessorial_row(spec: AccessorialCostSpec, price: float) -> AccessorialCostRow:
    return AccessorialCostRow(
        cost_name=spec.cost_name,
        cost_price=price,
        rate_by=spec.rate_by or DEFAULT_RATE_BY,
        apply_if=spec.apply_if or DEFAULT_APPLY_IF,
    )


def collect_accessorial_costs(
    combined_path: pd.ExcelFile | Any,
    used_tabs: set[str],
) -> list[AccessorialCostRow]:
    if not isinstance(combined_path, pd.ExcelFile):
        excel_file = pd.ExcelFile(combined_path)
    else:
        excel_file = combined_path

    unused_tabs = [tab for tab in excel_file.sheet_names if tab not in used_tabs]
    found_prices: dict[str, float] = {}

    for tab_name in unused_tabs:
        raw = pd.read_excel(excel_file, sheet_name=tab_name, header=None)
        for spec in ACCESSORIAL_COST_SPECS:
            if spec.cost_name in found_prices:
                continue
            price = find_spec_price_in_sheet(raw, spec)
            if price is not None:
                found_prices[spec.cost_name] = price
                print(f"  Accessorial: {spec.cost_name} = {price} (from tab '{tab_name}')")

    rows: list[AccessorialCostRow] = []
    for spec in ACCESSORIAL_COST_SPECS:
        price = found_prices.get(spec.cost_name)
        if price is None:
            continue
        rows.append(build_accessorial_row(spec, price))
    return rows


def get_used_matrix_tabs(
    parsed_by_role: dict[str, Any],
    return_tab: str,
) -> set[str]:
    used = {sheet.combined_tab for sheet in parsed_by_role.values()}
    used.add(return_tab)
    return used
