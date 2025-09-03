#!/usr/bin/env python3
"""
Sort an Excel workbook by 'Broker' groups with a priority order.

- Keeps original row order *within* each Broker group (stable).
- Orders Broker groups using the PRIORITY list below.
- Any Broker not in PRIORITY is placed after, sorted alphabetically by Broker.
- Can sort one or multiple sheets in the workbook.
- Writes a new workbook that includes *all* original sheets. Sheets selected
  for sorting are rewritten in sorted order; others are copied unchanged.

Usage:
  python inter_sort_broker_priority.py input.xlsx
  python inter_sort_broker_priority.py input.xlsx --sheet "Sheet1"
  python inter_sort_broker_priority.py input.xlsx --sheets "Sheet1" "Sheet2"
  python inter_sort_broker_priority.py input.xlsx --sheets "Sheet1,Sheet2"
  python inter_sort_broker_priority.py input.xlsx --sheets 0 2  # by index (0-based)
  python inter_sort_broker_priority.py input.xlsx --sheets all  # sort all sheets
  python inter_sort_broker_priority.py input.xlsx --output result.xlsx

Requires: pandas, openpyxl
  pip install pandas openpyxl
"""

from __future__ import annotations
import argparse
from pathlib import Path
import sys
from typing import List
import pandas as pd

# ---- EDIT YOUR PRIORITY ORDER HERE ----
PRIORITY = ["DWM", "FEDEX", "DHLE", "POL", "UPS"]
# --------------------------------------


def find_broker_column(df: pd.DataFrame) -> str:
    """Find the 'Broker' column case-insensitively and return its exact name in the DataFrame."""
    for col in df.columns:
        if str(col).strip().lower() == "broker":
            return col
    raise KeyError(
        "Could not find a 'Broker' column (case-insensitive). "
        f"Available columns: {list(df.columns)}"
    )


def sort_by_broker_priority(df: pd.DataFrame, broker_col: str) -> pd.DataFrame:
    """
    Stable-sort rows by Broker groups according to PRIORITY.
    - Rows inside the same Broker group keep their original order (stable).
    - Priority groups appear first, in PRIORITY order.
    - Non-priority groups come after, ordered alphabetically by Broker.
    """
    df = df.copy()
    df["_orig_idx"] = range(len(df))

    broker_vals = df[broker_col].astype(str).fillna("").str.strip()
    priority_index = {b: i for i, b in enumerate(PRIORITY)}

    df["_is_nonprio"] = (~broker_vals.isin(PRIORITY)).astype(int)
    df["_prio_idx"] = broker_vals.map(priority_index).fillna(0).astype(int)
    df["_broker_sort"] = broker_vals.where(df["_is_nonprio"] == 1, "")

    df_sorted = df.sort_values(
        by=["_is_nonprio", "_prio_idx", "_broker_sort", "_orig_idx"],
        kind="mergesort",
    )

    return df_sorted.drop(columns=["_is_nonprio", "_prio_idx", "_broker_sort", "_orig_idx"])


def _normalize_sheet_tokens(tokens: List[str]) -> List[str]:
    """Split comma-separated tokens and strip whitespace."""
    out: List[str] = []
    for tok in tokens:
        for piece in str(tok).split(","):
            piece = piece.strip()
            if piece:
                out.append(piece)
    return out


def resolve_target_sheets(sheets_args: List[str] | None, sheet_arg: str | None, all_sheet_names: List[str]) -> List[str]:
    """
    Determine which sheet names to operate on.
    - Accepts names (case-insensitive) or 0-based indices.
    - 'all' sorts all sheets.
    - If nothing specified, defaults to the first sheet.
    """
    targets: List[str] = []
    if sheets_args:
        tokens = _normalize_sheet_tokens(sheets_args)
        if any(t.lower() == "all" for t in tokens):
            targets = list(all_sheet_names)
        else:
            for t in tokens:
                # index?
                try:
                    idx = int(t)
                    if idx < 0 or idx >= len(all_sheet_names):
                        print(f"WARNING: sheet index {idx} out of range; skipping.", file=sys.stderr)
                        continue
                    targets.append(all_sheet_names[idx])
                    continue
                except ValueError:
                    pass
                # name (case-insensitive exact match)
                matches = [s for s in all_sheet_names if s.lower() == t.lower()]
                if matches:
                    targets.append(matches[0])
                else:
                    print(f"WARNING: sheet '{t}' not found; skipping.", file=sys.stderr)
    elif sheet_arg is not None:
        t = str(sheet_arg).strip()
        try:
            idx = int(t)
            if idx < 0 or idx >= len(all_sheet_names):
                print(f"WARNING: sheet index {idx} out of range; defaulting to first sheet.", file=sys.stderr)
                targets = [all_sheet_names[0]]
            else:
                targets = [all_sheet_names[idx]]
        except ValueError:
            matches = [s for s in all_sheet_names if s.lower() == t.lower()]
            targets = [matches[0]] if matches else []
            if not targets:
                print(f"WARNING: sheet '{t}' not found; defaulting to first sheet.", file=sys.stderr)
                targets = [all_sheet_names[0]]
    else:
        targets = [all_sheet_names[0]]

    # Deduplicate while preserving order
    seen = set()
    deduped = []
    for s in targets:
        if s not in seen:
            seen.add(s)
            deduped.append(s)
    return deduped


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Sort Excel rows by Broker groups with priority order.")
    parser.add_argument("input", type=str, help="Path to the input .xlsx file")
    parser.add_argument("--sheet", "-s", type=str, default=None,
                        help='Single worksheet name or index (0-based). Alias for passing one sheet.')
    parser.add_argument("--sheets", "-S", nargs="+", default=None,
                        help='One or more sheet names or indices (0-based). '
                             'Can also pass a single comma-separated string or "all".')
    parser.add_argument("--output", "-o", type=str, default=None,
                        help="Path to the output .xlsx file. Defaults to <input>-sorted.xlsx")
    args = parser.parse_args(argv)

    in_path = Path(args.input)
    if not in_path.exists():
        print(f"ERROR: Input file not found: {in_path}", file=sys.stderr)
        return 2

    try:
        xls = pd.ExcelFile(in_path, engine="openpyxl")
    except Exception as e:
        print(f"ERROR: Failed to open Excel file: {e}", file=sys.stderr)
        return 3

    target_sheet_names = resolve_target_sheets(args.sheets, args.sheet, xls.sheet_names)
    if not target_sheet_names:
        print("ERROR: No matching sheets found to process.", file=sys.stderr)
        return 4

    out_path = Path(args.output) if args.output else in_path.with_stem(in_path.stem + "-sorted")

    out_frames: dict[str, pd.DataFrame] = {}
    sorted_any = False

    for name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=name, engine="openpyxl")
        except Exception as e:
            print(f"WARNING: Failed to read sheet '{name}': {e}. Copying as-is.", file=sys.stderr)
            out_frames[name] = pd.DataFrame()
            continue

        if name in target_sheet_names:
            try:
                broker_col = find_broker_column(df)
                out_frames[name] = sort_by_broker_priority(df, broker_col)
                sorted_any = True
                print(f"Sorted sheet: {name}")
            except KeyError as e:
                print(f"WARNING: {name}: {e}. Leaving sheet unchanged.", file=sys.stderr)
                out_frames[name] = df
        else:
            out_frames[name] = df

    if not sorted_any:
        print("WARNING: No sheets were sorted (Broker column missing or none matched). Writing workbook unchanged.", file=sys.stderr)

    try:
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            for name, df in out_frames.items():
                # If a sheet failed to read (rare), skip writing empty DataFrame to avoid confusion
                if df is None or (isinstance(df, pd.DataFrame) and df.empty and name not in target_sheet_names):
                    continue
                df.to_excel(writer, sheet_name=name, index=False)
    except Exception as e:
        print(f"ERROR: Failed to write output Excel: {e}", file=sys.stderr)
        return 5

    print(f"Done. Wrote: {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
