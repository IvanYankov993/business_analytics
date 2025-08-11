#!/usr/bin/env python3
"""
Parse a grouped Excel report where each order header (1 row) is followed by an items table.
Outputs a normalized table (one row per item) as CSV or to stdout.

Usage:
  python parse_grouped_report.py --input test.xlsx --sheet "Лист1" --output parsed.csv
"""

import argparse
from typing import List, Dict, Any, Optional
import pandas as pd

def parse_grouped_report(
    path: str,
    sheet: Optional[str] = None,
    keep_bulgarian_headers: bool = True
) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object)
    n = len(df)
    i = 0
    records: List[Dict[str, Any]] = []

    def row_is_empty(r: int) -> bool:
        return df.iloc[r].isna().all()

    def is_order_header_row(r: int) -> bool:
        try:
            return (str(df.iat[r, 0]).strip() == "№" and str(df.iat[r, 1]).strip() == "Тип документ")
        except Exception:
            return False

    def is_items_header_row(r: int) -> bool:
        try:
            return (str(df.iat[r, 5]).strip() == "Стока" and str(df.iat[r, 8]).strip() == "Количество")
        except Exception:
            return False

    while i < n:
        if row_is_empty(i):
            i += 1
            continue

        if not is_order_header_row(i):
            i += 1
            continue

        header_labels = df.iloc[i].tolist()
        if i + 1 >= n:
            break
        header_values = df.iloc[i + 1].tolist()

        order_info: Dict[str, Any] = {}
        for lab, val in zip(header_labels, header_values):
            if pd.notna(lab):
                order_info[str(lab).strip()] = val

        j = i + 2
        while j < n and not is_items_header_row(j) and not is_order_header_row(j):
            j += 1

        if j >= n or is_order_header_row(j):
            i = j
            continue

        j += 1

        while j < n and not is_order_header_row(j):
            row = df.iloc[j]
            item_name = row.iloc[5] if len(row) > 5 else None
            if pd.isna(item_name):
                if row_is_empty(j):
                    j += 1
                    continue
                else:
                    j += 1
                    continue

            item_qty = row.iloc[8] if len(row) > 8 else None
            item_price = row.iloc[9] if len(row) > 9 else None
            item_sum = row.iloc[10] if len(row) > 10 else None
            item_vat = row.iloc[11] if len(row) > 11 else None

            record = dict(order_info)
            record.update({
                "Стока": item_name,
                "Количество": item_qty,
                "Цена": item_price,
                "Сума": item_sum,
                "ДДС": item_vat,
            })
            records.append(record)
            j += 1

        i = j

    result = pd.DataFrame.from_records(records)

    if not keep_bulgarian_headers and len(result) > 0:
        rename_map = {
            "№": "Document No",
            "Тип документ": "Doc Type",
            "Дата": "Date",
            "Партньор": "Partner",
            "Булстат": "Company ID",
            "Тип плащане": "Payment Type",
            "Потребител": "User",
            "Дата на падеж": "Due Date",
            "Дни на падеж": "Days to Due",
            "Стока": "Item",
            "Количество": "Quantity",
            "Цена": "Unit Price",
            "Сума": "Line Total",
            "ДДС": "VAT",
        }
        result = result.rename(columns={k: v for k, v in rename_map.items() if k in result.columns})

    return result

def main():
    ap = argparse.ArgumentParser(description="Parse grouped Excel report into a normalized items table")
    ap.add_argument("--input", "-i", required=True, help="Path to input .xlsx file")
    ap.add_argument("--sheet", "-s", default=None, help="Sheet name (default: first sheet)")
    ap.add_argument("--output", "-o", default=None, help="Path to output CSV (default: print to stdout)")
    ap.add_argument("--english", action="store_true", help="Rename Bulgarian headers to English")
    args = ap.parse_args()

    df = parse_grouped_report(args.input, sheet=args.sheet, keep_bulgarian_headers=not args.english)

    if args.output:
        df.to_csv(args.output, index=False)
        print(f"Saved: {args.output}")
    else:
        # Print without index
        with pd.option_context('display.max_rows', None, 'display.max_columns', None, 'display.width', 200):
            print(df.to_string(index=False))

if __name__ == "__main__":
    main()
