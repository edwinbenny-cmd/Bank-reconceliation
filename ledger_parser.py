"""
ledger_parser.py

Parses Tally-style "Ledger Account" exports (the kind where each ledger's
transactions are dumped as multi-row vouchers with narration lines folded
into the same column as the account names) into a clean, flat table:

    Sheet | Ledger | Date | Type | Vch Type | Vch No | Particulars |
    Debit | Credit | Balance | Narration | Row Type

Works on a workbook with one or many sheets (e.g. a "CASH" tab and a
"BANK" tab), each containing one or many "Ledger:" blocks.

Usage:
    from ledger_parser import parse_workbook
    df = parse_workbook("CASH-BANK - copy.xlsx")
    df.to_excel("structured_ledger.xlsx", index=False)

Or from the command line:
    python ledger_parser.py "CASH-BANK - copy.xlsx" output.xlsx
"""

import sys
import datetime
import openpyxl
import pandas as pd


def _is_date(v):
    return isinstance(v, (datetime.datetime, datetime.date))


def _clean(v):
    """Blank out None / empty-string / whitespace-only cell values."""
    if v is None:
        return None
    if isinstance(v, str) and v.strip() == "":
        return None
    return v


def _is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def parse_sheet(ws, sheet_name):
    """
    Parse a single worksheet and return a list of row-dicts.

    Column layout (fixed by the source export):
        A = Date          B = Dr/Cr          C = Particulars
        D,E = spacers     F = Vch Type       G = Vch No
        H = Debit         I = Credit         J = Balance
    """
    records = []

    current_ledger = None
    current_period = None

    # State for the voucher currently being assembled.
    voucher = None   # dict with date/type/vchtype/vchno/balance
    lines = []        # list of dicts: {particulars, debit, credit, narrations: []}

    def flush_voucher():
        """Close out the voucher being built, back-filling narration and emitting rows."""
        if voucher is None or not lines:
            return
        # The last narration seen in this voucher is the "common" one Tally prints
        # once at the end. Any allocation line that has no narration of its own
        # inherits it, as the user described.
        common_narration = None
        for ln in reversed(lines):
            if ln["narrations"]:
                common_narration = " | ".join(ln["narrations"])
                break

        for ln in lines:
            narration = " | ".join(ln["narrations"]) if ln["narrations"] else common_narration
            records.append({
                "Sheet": sheet_name,
                "Ledger": current_ledger,
                "Period": current_period,
                "Date": voucher["date"],
                "Type": voucher["type"],
                "Vch Type": voucher["vch_type"],
                "Vch No": voucher["vch_no"],
                "Particulars": ln["particulars"],
                "Debit": ln["debit"],
                "Credit": ln["credit"],
                "Balance": voucher["balance"],
                "Narration": narration,
                "Row Type": voucher["row_type"],
            })

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        a = _clean(row[0].value) if len(row) > 0 else None
        b = _clean(row[1].value) if len(row) > 1 else None
        c = _clean(row[2].value) if len(row) > 2 else None
        f = _clean(row[5].value) if len(row) > 5 else None   # Vch Type
        g = _clean(row[6].value) if len(row) > 6 else None   # Vch No
        h = row[7].value if len(row) > 7 else None           # Debit
        i = row[8].value if len(row) > 8 else None           # Credit
        j = row[9].value if len(row) > 9 else None           # Balance
        h = h if _is_number(h) else None
        i = i if _is_number(i) else None
        j = j if _is_number(j) else None

        # --- New ledger block ---
        if a == "Ledger:":
            flush_voucher()
            voucher, lines = None, []
            current_ledger = b
            current_period = c
            continue

        # --- Header row ("Date | Particulars | Vch Type | ...") ---
        if a == "Date" and b == "Particulars":
            continue

        # --- Fully blank row: end of this ledger's transaction list ---
        if a is None and b is None and c is None and h is None and i is None and j is None:
            flush_voucher()
            voucher, lines = None, []
            continue

        # --- Ledger grand-total row, e.g. C=6037057.09 with no A/B/F/G ---
        if a is None and b is None and _is_number(c) and f is None and g is None:
            flush_voucher()
            voucher, lines = None, []
            records.append({
                "Sheet": sheet_name, "Ledger": current_ledger, "Period": current_period,
                "Date": None, "Type": None, "Vch Type": None, "Vch No": None,
                "Particulars": "Total", "Debit": h, "Credit": i, "Balance": None,
                "Narration": None, "Row Type": "Ledger Total",
            })
            continue

        # --- Opening/Closing balance row: B has Dr/Cr, C names it, no F/G ---
        if b in ("Dr", "Cr") and c in ("Opening Balance", "Closing Balance") and f is None:
            flush_voucher()
            voucher, lines = None, []
            records.append({
                "Sheet": sheet_name, "Ledger": current_ledger, "Period": current_period,
                "Date": a, "Type": b, "Vch Type": None, "Vch No": None,
                "Particulars": c, "Debit": h, "Credit": i, "Balance": j,
                "Narration": None,
                "Row Type": "Opening Balance" if c == "Opening Balance" else "Closing Balance",
            })
            continue

        # --- New voucher header: has a date in column A ---
        if _is_date(a):
            flush_voucher()
            voucher = {
                "date": a, "type": b, "vch_type": f, "vch_no": g,
                "balance": j, "row_type": "Transaction",
            }
            lines = [{"particulars": c, "debit": h, "credit": i, "narrations": []}]
            continue

        # --- Allocation line: has its own account name + an amount, no date ---
        if voucher is not None and c is not None and (h is not None or i is not None):
            lines.append({"particulars": c, "debit": h, "credit": i, "narrations": []})
            continue

        # --- Narration line: text only, no amount -> attach to the last line ---
        if voucher is not None and c is not None and h is None and i is None:
            if lines:
                lines[-1]["narrations"].append(str(c).strip())
            continue

        # Anything else (stray blank-ish rows) is ignored.

    flush_voucher()
    return records


def parse_workbook(path):
    """Parse every sheet in the workbook and return one combined DataFrame."""
    wb = openpyxl.load_workbook(path, data_only=True)
    all_records = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        all_records.extend(parse_sheet(ws, sheet_name.strip()))

    df = pd.DataFrame(all_records, columns=[
        "Sheet", "Ledger", "Period", "Date", "Type", "Vch Type", "Vch No",
        "Particulars", "Debit", "Credit", "Balance", "Narration", "Row Type",
    ])

    # Tidy date formatting for readability / downstream filtering.
    if not df.empty:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    return df


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python ledger_parser.py <input.xlsx> [output.xlsx]")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "structured_ledger.xlsx"

    df = parse_workbook(input_path)
    print(f"Parsed {len(df)} rows across {df['Sheet'].nunique()} sheet(s) "
          f"and {df['Ledger'].nunique()} ledger(s).")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Structured", index=False)
        ws = writer.sheets["Structured"]

        from openpyxl.styles import Font
        header_font = Font(name="Arial", bold=True)
        body_font = Font(name="Arial")
        widths = {
            "A": 10, "B": 24, "C": 20, "D": 12, "E": 8, "F": 10, "G": 22,
            "H": 12, "I": 12, "J": 12, "K": 55, "L": 16,
        }
        for col, width in widths.items():
            ws.column_dimensions[col].width = width

        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for cell in row:
                cell.font = header_font if cell.row == 1 else body_font

        # Vch No (col G) as text so "805" never turns into 805.0 downstream.
        for cell in ws["G"][1:]:
            cell.number_format = "@"
        # Debit/Credit/Balance as 2-decimal numbers.
        for col_letter in ("H", "I", "J"):
            for cell in ws[col_letter][1:]:
                cell.number_format = "#,##0.00"
        # Date column.
        for cell in ws["D"][1:]:
            cell.number_format = "DD-MMM-YY"

        ws.freeze_panes = "A2"

    print(f"Written to {output_path}")
