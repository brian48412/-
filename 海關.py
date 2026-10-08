"""
China Customs data update script
================================
Updates China.xls and Chinese Mainland's Trade Statistics.xls
using data from File_1.xls ~ File_4.xls.
"""

import os
from datetime import datetime
from dateutil.relativedelta import relativedelta

import pythoncom
import win32com.client as win32

# ============================================================
# CONFIG
# ============================================================

BASE_DIR = r"C:\Users\khchu1\Desktop\Brian2\中國海關"

CHINA_FILE       = "China.xls"
FILE_1           = "File_1.xls"
FILE_2           = "File_2.xls"
FILE_3           = "File_3.xls"
FILE_4           = "File_4.xls"
TRADE_STATS_FILE = "Chinese Mainland's Trade Statistics.xls"

xlPasteValues   = -4163
xlPasteFormats  = -4122
xlPasteFormulas = -4123
xlValues        = -4163

# Number formats
FORMAT_WITH_R    = '###0.0 "(r)"'
FORMAT_WITHOUT_R = '#,##0.0  '
FORMAT_HK_VALUE  = '0.00'          # 2 decimal places for R_China(HK) D & H

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# ============================================================
# HELPERS
# ============================================================

def get_excel():
    try:
        os.system('taskkill /F /IM EXCEL.EXE >nul 2>&1')
    except Exception:
        pass
    pythoncom.CoInitialize()
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    return excel


def close_all(excel, *workbooks):
    excel.ScreenUpdating = True
    excel.EnableEvents = True
    for wb in workbooks:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
    try:
        excel.Quit()
    except Exception:
        pass


def safe_float(val):
    """Convert Excel cell value to float safely."""
    if val is None:
        return None
    try:
        return float(val)
    except (TypeError, ValueError):
        try:
            return float(str(val).replace(",", "").strip())
        except Exception:
            return None


def next_month_name(month_str: str) -> str:
    idx = MONTHS.index(month_str.strip())
    return MONTHS[(idx + 1) % 12]


def find_latest_month_row(ws, start_row: int, end_row: int) -> int:
    latest_row = None
    for row in range(start_row, end_row + 1):
        val = ws.Cells(row, 2).Value
        if val and str(val).strip() in MONTHS:
            latest_row = row
    if latest_row is None:
        raise ValueError(f"No month found in B{start_row}:B{end_row} on {ws.Name}")
    return latest_row


def insert_new_row_after_latest(ws, search_start: int, search_end: int, add_star: bool = False):
    """
    Insert a new row after the latest month.
    If add_star=True (for R_China(HK)), also write "*" into column C.
    """
    latest_row = find_latest_month_row(ws, search_start, search_end)
    latest_month = str(ws.Cells(latest_row, 2).Value).strip()
    next_m = next_month_name(latest_month)

    ws.Rows(latest_row + 1).Insert()
    new_row = latest_row + 1
    ws.Cells(new_row, 2).Value = next_m

    if add_star:
        ws.Cells(new_row, 3).Value = "*"          # column C

    print(f"  [{ws.Name}] Inserted new_row = {new_row}  (month = {next_m})")
    return new_row


def get_year_of_row(ws, row: int) -> int:
    for r in range(row, 0, -1):
        val = ws.Cells(r, 1).Value
        if val is not None:
            try:
                y = int(val)
                if 2000 <= y <= 2100:
                    return y
            except (TypeError, ValueError):
                pass
    raise ValueError(f"Cannot determine year for row {row} on {ws.Name}")


def find_month_row(ws, year: int, month: str, search_start: int = 300, search_end: int = 450) -> int:
    for row in range(search_start, search_end + 1):
        b_val = str(ws.Cells(row, 2).Value or "").strip()
        if b_val != month:
            continue
        try:
            y = get_year_of_row(ws, row)
            if y == year:
                return row
        except ValueError:
            continue
    raise ValueError(f"Cannot find {year} {month} on sheet {ws.Name}")


def try_find_month_row(ws, year: int, month: str, search_start: int = 300, search_end: int = 450):
    """Like find_month_row but returns None when the month is missing."""
    try:
        return find_month_row(ws, year, month, search_start, search_end)
    except ValueError:
        return None


def autofill_column(ws, src_row: int, dest_row: int, col: int):
    """Drag-fill a formula column so relative references adjust (Excel AutoFill)."""
    src = ws.Cells(src_row, col)
    dest_range = ws.Range(ws.Cells(src_row, col), ws.Cells(dest_row, col))
    src.AutoFill(Destination=dest_range)


def write_hk_value(ws, row: int, col: int, value: float):
    """
    Write a R_China(HK) D/H value with consistent format/alignment.
    Prefer the number format already used in the same column on the previous month row.
    """
    cell = ws.Cells(row, col)
    cell.Value = value
    prev_fmt = ws.Cells(row - 1, col).NumberFormat
    if prev_fmt and str(prev_fmt).strip() and str(prev_fmt) != "General":
        cell.NumberFormat = prev_fmt
    else:
        cell.NumberFormat = FORMAT_HK_VALUE
    cell.HorizontalAlignment = ws.Cells(row - 1, col).HorizontalAlignment



FORMAT_PLAIN_FALLBACK = '#,##0.0'


def has_r_mark(number_format) -> bool:
    """True if an Excel number format shows an (r) mark (any quoting/escaping style)."""
    if not number_format:
        return False
    s = str(number_format).replace('\\', '').replace('"', '').lower()
    return '(r)' in s


def safe_set_number_format(cell, fmt, fallback=FORMAT_PLAIN_FALLBACK):
    """Set NumberFormat; if Excel rejects it, use the fallback instead of crashing."""
    try:
        cell.NumberFormat = fmt
    except Exception:
        print(f"    Warning: Excel rejected format {fmt!r} at {cell.Address}; using {fallback!r}")
        cell.NumberFormat = fallback


def find_plain_trade_format(ws, col: int, start_row: int, max_rows_up: int = 120) -> str:
    """
    Sheet 3's own format for cells without (r) in this column:
    scan upward from start_row for the nearest cell whose format has no (r).
    """
    for r in range(start_row, max(start_row - max_rows_up, 1) - 1, -1):
        cell = ws.Cells(r, col)
        if cell.Value is None:
            continue
        fmt = cell.NumberFormat
        if fmt and str(fmt).strip() and not has_r_mark(fmt):
            return fmt
    return FORMAT_PLAIN_FALLBACK


def paste_trade_cell(ws, dst_row, dst_col, value, src_fmt, plain_fmt, force_no_r: bool = False):
    """
    Write a value into sheet 3.
    - Source shows (r) and not force_no_r -> copy the source (r) format.
    - Otherwise -> keep sheet 3's own plain format for that column.
    """
    cell = ws.Cells(dst_row, dst_col)
    cell.Value = value
    if not force_no_r and has_r_mark(src_fmt):
        safe_set_number_format(cell, src_fmt, plain_fmt)
    else:
        safe_set_number_format(cell, plain_fmt)


def write_value_with_r_logic(cell, new_value, tracked_cells: list, force_no_r: bool = False):
    """
    Write value and decide (r) mark.
    force_no_r=True → always without (r)  (used for new_row)
    """
    old_value = cell.Value
    cell.Value = new_value

    if force_no_r:
        cell.NumberFormat = FORMAT_WITHOUT_R
    else:
        try:
            changed = (old_value is None) or (abs(float(old_value) - float(new_value)) > 1e-9)
        except (TypeError, ValueError):
            changed = True
        cell.NumberFormat = FORMAT_WITH_R if changed else FORMAT_WITHOUT_R

    tracked_cells.append(cell)


# ============================================================
# MAIN
# ============================================================

def main():
    excel = get_excel()
    wb_china = wb_f1 = wb_f2 = wb_f3 = wb_f4 = wb_trade = None
    tracked = []

    try:
        print("Opening workbooks ...")
        wb_china = excel.Workbooks.Open(os.path.join(BASE_DIR, CHINA_FILE), UpdateLinks=False)
        wb_f1    = excel.Workbooks.Open(os.path.join(BASE_DIR, FILE_1), UpdateLinks=False, ReadOnly=True)
        wb_f2    = excel.Workbooks.Open(os.path.join(BASE_DIR, FILE_2), UpdateLinks=False, ReadOnly=True)
        wb_f3    = excel.Workbooks.Open(os.path.join(BASE_DIR, FILE_3), UpdateLinks=False, ReadOnly=True)
        wb_f4    = excel.Workbooks.Open(os.path.join(BASE_DIR, FILE_4), UpdateLinks=False, ReadOnly=True)
        wb_trade = excel.Workbooks.Open(os.path.join(BASE_DIR, TRADE_STATS_FILE), UpdateLinks=False)

        ws_r_china    = wb_china.Sheets("R_China")
        ws_r_china_hk = wb_china.Sheets("R_China(HK)")
        ws_f1 = wb_f1.Sheets(1)
        ws_f2 = wb_f2.Sheets(1)
        ws_f3 = wb_f3.Sheets(1)
        ws_f4 = wb_f4.Sheets(1)

        try:
            ws_trade_main = wb_trade.Sheets("The Mainland Trade Statistics")
        except Exception:
            ws_trade_main = wb_trade.Sheets(1)

        # ==================================================
        # PART 1 – R_China
        # ==================================================
        print("\n========== PART 1 : R_China ==========")

        # 1.1 Insert new_row
        new_row = insert_new_row_after_latest(ws_r_china, 400, 411)
        prev_row = new_row - 1

        # 1.2 File_4 → new_row  (force_no_r = True)
        print("  Writing File_4 values into new_row ...")
        v = safe_float(ws_f4.Range("C7").Value)
        if v is not None:
            write_value_with_r_logic(ws_r_china.Cells(new_row, 4), v / 10, tracked, force_no_r=True)
        v = safe_float(ws_f4.Range("C6").Value)
        if v is not None:
            write_value_with_r_logic(ws_r_china.Cells(new_row, 8), v / 10, tracked, force_no_r=True)
        v = safe_float(ws_f4.Range("F7").Value)
        if v is not None:
            write_value_with_r_logic(ws_r_china.Cells(new_row, 6), v, tracked, force_no_r=True)
        v = safe_float(ws_f4.Range("F6").Value)
        if v is not None:
            write_value_with_r_logic(ws_r_china.Cells(new_row, 10), v, tracked, force_no_r=True)

        # 1.3 Drag-fill formulas L:O (AutoFill adjusts relative refs; plain copy does not)
        print("  AutoFilling formulas L:O ...")
        for col in range(12, 16):
            autofill_column(ws_r_china, prev_row, new_row, col)

        # 1.4 Last-year January block
        curr_year = get_year_of_row(ws_r_china, new_row)
        last_year = curr_year - 1
        last_year_jan_row = find_month_row(ws_r_china, last_year, "Jan")
        print(f"  Last-year January found at row {last_year_jan_row}")

        for i in range(12):
            src_e = safe_float(ws_f3.Cells(6 + i, 5).Value)
            src_d = safe_float(ws_f3.Cells(6 + i, 4).Value)
            if src_e is not None:
                write_value_with_r_logic(ws_r_china.Cells(last_year_jan_row + i, 4), src_e / 1000, tracked)
            if src_d is not None:
                write_value_with_r_logic(ws_r_china.Cells(last_year_jan_row + i, 8), src_d / 1000, tracked)

        # 1.5 Current-year Import (green)
        print("  Updating current-year Import side (green) ...")
        jan_row = find_month_row(ws_r_china, curr_year, "Jan")
        month_idx = 0
        src_base = 19
        while True:
            target_row = jan_row + month_idx
            if target_row >= new_row:
                break
            val_cell = safe_float(ws_f3.Cells(src_base, 5).Value)
            yoy_cell = safe_float(ws_f3.Cells(src_base + 1, 5).Value)
            if val_cell is not None:
                write_value_with_r_logic(ws_r_china.Cells(target_row, 4), val_cell / 1000, tracked)
            if yoy_cell is not None:
                write_value_with_r_logic(ws_r_china.Cells(target_row, 6), yoy_cell, tracked)
            month_idx += 1
            src_base += 3

        # 1.6 Current-year Export (blue)
        print("  Updating current-year Export side (blue) ...")
        month_idx = 0
        src_base = 19
        while True:
            target_row = jan_row + month_idx
            if target_row >= new_row:
                break
            val_cell = safe_float(ws_f3.Cells(src_base, 4).Value)
            yoy_cell = safe_float(ws_f3.Cells(src_base + 1, 4).Value)
            if val_cell is not None:
                write_value_with_r_logic(ws_r_china.Cells(target_row, 8), val_cell / 1000, tracked)
            if yoy_cell is not None:
                write_value_with_r_logic(ws_r_china.Cells(target_row, 10), yoy_cell, tracked)
            month_idx += 1
            src_base += 3

        # ==================================================
        # PART 2 – R_China(HK)
        # ==================================================
        print("\n========== PART 2 : R_China(HK) ==========")

        # Insert new_row + put "*" in column C
        new_row_hk = insert_new_row_after_latest(ws_r_china_hk, 398, 409, add_star=True)
        prev_row_hk = new_row_hk - 1

        # Resolve write targets by year+month (dynamic; survives row shifts)
        new_month_hk = str(ws_r_china_hk.Cells(new_row_hk, 2).Value).strip()
        prev_month_hk = str(ws_r_china_hk.Cells(prev_row_hk, 2).Value).strip()
        new_year_hk = get_year_of_row(ws_r_china_hk, new_row_hk)
        prev_year_hk = get_year_of_row(ws_r_china_hk, prev_row_hk)
        target_new_hk = find_month_row(ws_r_china_hk, new_year_hk, new_month_hk)
        target_prev_hk = find_month_row(ws_r_china_hk, prev_year_hk, prev_month_hk)
        print(f"  HK targets: {prev_year_hk} {prev_month_hk} → row {target_prev_hk}, "
              f"{new_year_hk} {new_month_hk} → row {target_new_hk}")

        # File_1 → new month row (2 decimal places, aligned like previous month)
        print("  Writing File_1 values into new month row ...")
        v = safe_float(ws_f1.Range("E21").Value)
        if v is not None:
            write_hk_value(ws_r_china_hk, target_new_hk, 8, v / 1000)   # H

        v = safe_float(ws_f1.Range("G21").Value)
        if v is not None:
            write_hk_value(ws_r_china_hk, target_new_hk, 4, v / 1000)   # D

        # AutoFill formulas for F, J, L, N  (all four columns)
        print("  Filling formulas F / J / L / N (AutoFill) ...")
        for col in [6, 10, 12, 14]:          # F, J, L, N
            autofill_column(ws_r_china_hk, target_prev_hk, target_new_hk, col)

        # File_2 → previous month row (2 decimal places)
        print("  Writing File_2 values into previous month row ...")
        v = safe_float(ws_f2.Range("E17").Value)
        if v is not None:
            write_hk_value(ws_r_china_hk, target_prev_hk, 8, v / 1000000)  # H

        v = safe_float(ws_f2.Range("G17").Value)
        if v is not None:
            write_hk_value(ws_r_china_hk, target_prev_hk, 4, v / 1000000)  # D

        # ==================================================
        # PART 3 – Trade Statistics
        # ==================================================
        print("\n========== PART 3 : Trade Statistics ==========")

        # 3.1 Insert new month row (same pattern as R_China)
        print("  Inserting new month row on Trade Stats ...")
        new_row_trade = insert_new_row_after_latest(ws_trade_main, 380, 450)
        prev_row_trade = new_row_trade - 1
        month_trade = str(ws_trade_main.Cells(new_row_trade, 2).Value or "").strip()
        print(f"  Trade Stats new_row_trade = {new_row_trade}, prev_row_trade = {prev_row_trade}")

        # 3.2 Copy row formatting from previous month onto the new row
        print("  Copying row formats from prev_row_trade → new_row_trade ...")
        ws_trade_main.Rows(prev_row_trade).Copy()
        ws_trade_main.Rows(new_row_trade).PasteSpecial(Paste=xlPasteFormats)
        excel.CutCopyMode = False
        # Re-write month name in column B (PasteFormats may overwrite it)
        ws_trade_main.Cells(new_row_trade, 2).Value = month_trade

        # Sheet 3's own plain (no-(r)) format per column, read before writing
        trade_cols = [4, 6, 11, 13, 15, 17, 19, 21]   # D F K M O Q S U
        plain_fmt = {c: find_plain_trade_format(ws_trade_main, c, prev_row_trade)
                     for c in trade_cols}

        # 3.3 New month only — paste mapped values; new row never has (r)
        print(f"  Pasting new-month values → Trade Stats row {new_row_trade} ...")
        print(f"    sources: R_China row {new_row}, R_China(HK) row {target_new_hk}")

        new_row_pairs = [
            (ws_r_china,    new_row,       8,  4),   # R_China H → D
            (ws_r_china,    new_row,       10, 6),   # R_China J → F
            (ws_r_china,    new_row,       4,  15),  # R_China D → O
            (ws_r_china,    new_row,       6,  17),  # R_China F → Q
            (ws_r_china_hk, target_new_hk, 8,  11),  # HK H → K
            (ws_r_china_hk, target_new_hk, 10, 13),  # HK J → M
            (ws_r_china_hk, target_new_hk, 4,  19),  # HK D → S
            (ws_r_china_hk, target_new_hk, 6,  21),  # HK F → U
        ]
        for src_ws, src_row, src_col, dst_col in new_row_pairs:
            src_cell = src_ws.Cells(src_row, src_col)
            paste_trade_cell(ws_trade_main, new_row_trade, dst_col, src_cell.Value,
                             src_cell.NumberFormat, plain_fmt[dst_col], force_no_r=True)
        print("  New-month paste done (D/F/O/Q from R_China; K/M/S/U from R_China(HK)), no (r)")

        # 3.3b AutoFill formula columns W and Y (drag prev → new; relative refs adjust)
        print("  AutoFilling Trade Stats formulas W and Y ...")
        for col in [23, 25]:  # W, Y
            autofill_column(ws_trade_main, prev_row_trade, new_row_trade, col)

        # 3.4 Historical sync: last year Jan → latest−1 (exclude new insert month)
        # (r) in source -> copy source (r) format; no (r) -> sheet 3's own plain format
        print(f"  Historical sync R_China rows {last_year_jan_row}–{prev_row} → Trade Stats "
              f"(H→D, J→F, D→O, F→Q) ...")
        copied = 0
        hist_pairs = [(8, 4), (10, 6), (4, 15), (6, 17)]   # H→D, J→F, D→O, F→Q
        for src_row in range(last_year_jan_row, prev_row + 1):
            month = str(ws_r_china.Cells(src_row, 2).Value or "").strip()
            if month not in MONTHS:
                continue
            try:
                year = get_year_of_row(ws_r_china, src_row)
            except ValueError:
                continue
            dst_row = try_find_month_row(ws_trade_main, year, month)
            if dst_row is None:
                print(f"  Skip {year} {month}: not found on Trade Stats (no row created)")
                continue
            if dst_row == new_row_trade:
                continue  # new month already handled above

            for src_col, dst_col in hist_pairs:
                src_cell = ws_r_china.Cells(src_row, src_col)
                paste_trade_cell(ws_trade_main, dst_row, dst_col, src_cell.Value,
                                 src_cell.NumberFormat, plain_fmt[dst_col])
            copied += 1

        print(f"  Historical sync copied H→D/J→F/D→O/F→Q for {copied} existing month rows")

        # Save
        print("\nSaving workbooks ...")
        wb_china.Save()
        wb_trade.Save()
        print("All done successfully.")

    finally:
        close_all(excel, wb_china, wb_f1, wb_f2, wb_f3, wb_f4, wb_trade)


if __name__ == "__main__":
    main()
