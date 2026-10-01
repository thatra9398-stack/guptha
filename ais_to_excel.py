#!/usr/bin/env python3
"""
AIS JSON -> structured Excel converter.

Converts the Annual Information Statement (AIS) JSON downloaded from the
Income Tax e-Filing portal into a multi-sheet Excel workbook.
"""

from __future__ import annotations

import argparse
import getpass
import sys
from decimal import Decimal
from pathlib import Path

from aisparser import read_ais
from aisparser.exceptions import AisParserError
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=13)
LABEL_FONT = Font(bold=True)

AMOUNT_HINTS = (
    "amount", "value", "consideration", "tax", "tds", "tcs", "cess",
    "surcharge", "salary", "interest", "fmv", "dividend", "deposit",
    "paid", "credited", "collected", "deducted", "refund", "balance",
    "cost", "acquisition", "purchase", "receipt", "expenditure", "gross",
    "net", "premium", "remittance", "turnover", "income", "payment", "rent",
)

PROVENANCE_LABELS = {
    "part": "Part",
    "section_key": "Section Key",
    "l1_src": "Source",
    "info_src_id": "Reporter ID",
    "title": "Category",
}


def looks_like_amount(column_name: str) -> bool:
    return any(hint in column_name.lower() for hint in AMOUNT_HINTS)


def sheet_base_name(key: str, df) -> str:
    if len(df) and "title" in df.columns:
        title = str(df["title"].iloc[0]).strip()
        if title:
            return title
    return key.replace("_", " ").title()


def sanitize_sheet_name(name: str, used: set[str]) -> str:
    cleaned = "".join(ch for ch in name if ch not in "[]:*?/\\").strip() or "Sheet"
    cleaned = cleaned[:31]
    candidate, n = cleaned, 1
    while candidate in used:
        suffix = f"_{n}"
        candidate = cleaned[: 31 - len(suffix)] + suffix
        n += 1
    used.add(candidate)
    return candidate


def style_header_row(ws, row_idx: int, ncols: int) -> None:
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def autosize_columns(ws, ncols: int, start_row: int = 1) -> None:
    for col in range(1, ncols + 1):
        letter = get_column_letter(col)
        longest = max(
            (len(str(ws.cell(row=r, column=col).value or "")) for r in range(start_row, ws.max_row + 1)),
            default=0,
        )
        ws.column_dimensions[letter].width = min(max(longest + 2, 10), 45)


def write_category_sheet(wb: Workbook, sheet_name: str, columns: list, rows: list) -> None:
    ws = wb.create_sheet(title=sheet_name)
    ws.append([PROVENANCE_LABELS.get(c, c) for c in columns])
    style_header_row(ws, 1, len(columns))
    amount_cols = {i for i, c in enumerate(columns) if looks_like_amount(c)}
    for row in rows:
        out = []
        for col in columns:
            v = row.get(col)
            if isinstance(v, Decimal):
                v = float(v)
            out.append(v)
        ws.append(out)
    for i in amount_cols:
        letter = get_column_letter(i + 1)
        for cell in ws[letter][1:]:
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0.00"
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    autosize_columns(ws, len(columns))


def write_summary_sheet(wb: Workbook, statement, frames: dict, sheet_name_by_key: dict) -> None:
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = "AIS Statement Summary"
    ws["A1"].font = TITLE_FONT

    r = 3
    meta = statement.metadata
    for label, value in [
        ("PAN", meta.loggedInPan),
        ("Assessment Year", statement.assessment_year),
        ("AIS Download Date", meta.downloadDate),
        ("JSON Version", meta.jsonVersion),
        ("Utility Version", meta.utilityVersion),
    ]:
        ws.cell(row=r, column=1, value=label).font = LABEL_FONT
        ws.cell(row=r, column=2, value=value if value is not None else "-")
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="Part").font = LABEL_FONT
    ws.cell(row=r, column=2, value="Category").font = LABEL_FONT
    ws.cell(row=r, column=3, value="Records").font = LABEL_FONT
    ws.cell(row=r, column=4, value="Sheet").font = LABEL_FONT
    style_header_row(ws, r, 4)
    r += 1
    for key, df in sorted(frames.items()):
        part = str(df["part"].iloc[0]) if len(df) and "part" in df.columns else "-"
        ws.cell(row=r, column=1, value=part)
        ws.cell(row=r, column=2, value=sheet_base_name(key, df))
        ws.cell(row=r, column=3, value=len(df))
        ws.cell(row=r, column=4, value=sheet_name_by_key.get(key) or "(no records)")
        r += 1

    if statement.parse_warnings:
        from collections import Counter

        NO_DATA_LOSS_CODES = {"row_column_count_mismatch", "duplicate_column_label"}
        RAW_ONLY_CODES = {"non_grid_structure_kept_raw", "unrecognized_top_level_key"}

        raw_only = [w for w in statement.parse_warnings if w.reason_code in RAW_ONLY_CODES]
        raw_only_partB = [w for w in raw_only if w.source_path and w.source_path[0] == "partB"]

        r += 2
        ws.cell(row=r, column=1, value=f"Parse Notes ({len(statement.parse_warnings)})").font = LABEL_FONT
        r += 1
        if raw_only_partB:
            note = (
                f"{len(raw_only_partB)} note(s) involve partB financial content not in grid shape "
                "-- check source JSON directly for raw structure."
            )
        elif raw_only:
            note = "All notes below are formatting quirks or metadata warnings; no category data lost."
        else:
            note = "All notes below are formatting quirks; no category data lost."
        ws.cell(row=r, column=1, value=note)
        r += 2

        ws.cell(row=r, column=1, value="Note type").font = LABEL_FONT
        ws.cell(row=r, column=2, value="Count").font = LABEL_FONT
        style_header_row(ws, r, 2)
        r += 1
        for code, n in Counter(w.reason_code for w in statement.parse_warnings).most_common():
            label = code.replace("_", " ").capitalize()
            if code not in NO_DATA_LOSS_CODES and code not in RAW_ONLY_CODES:
                label += " (check source)"
            ws.cell(row=r, column=1, value=label)
            ws.cell(row=r, column=2, value=n)
            r += 1

    autosize_columns(ws, 4)
    ws.column_dimensions["A"].width = 45


def convert(json_path: Path, password: str, output_path: Path) -> None:
    try:
        statement = read_ais(path=json_path, password=password)
    except AisParserError as exc:
        sys.exit(f"Could not read the AIS file: {exc}")

    frames = statement.to_df()
    used_names = {"Summary"}
    sheet_name_by_key = {
        key: sanitize_sheet_name(sheet_base_name(key, df), used_names) if len(df) > 0 else None
        for key, df in sorted(frames.items())
    }

    wb = Workbook()
    write_summary_sheet(wb, statement, frames, sheet_name_by_key)

    for key, df in sorted(frames.items()):
        if len(df) == 0:
            continue
        rows = df.to_dict(orient="records")
        write_category_sheet(wb, sheet_name_by_key[key], list(df.columns), rows)

    wb.save(output_path)
    print(f"Done: {len(frames)} categories processed -> {output_path}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert an AIS JSON download into a structured Excel workbook.")
    parser.add_argument("json_path", type=Path, help="Path to the AIS JSON file downloaded from the portal")
    parser.add_argument("-o", "--output", type=Path, default=None, help="Output .xlsx path")
    parser.add_argument("--pan", default=None, help="PAN")
    parser.add_argument("--dob", default=None, help="Date of birth as DDMMYYYY")
    args = parser.parse_args()

    if not args.json_path.exists():
        sys.exit(f"File not found: {args.json_path}")

    pan = args.pan or input("PAN: ").strip()
    dob = args.dob or getpass.getpass("Date of birth (DDMMYYYY, hidden): ").strip()
    password = f"{pan}{dob}"

    output_path = args.output or args.json_path.with_suffix(".xlsx")
    convert(args.json_path, password, output_path)


if __name__ == "__main__":
    main()
