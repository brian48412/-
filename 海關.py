"""
China Customs data update script
================================
Updates China.xls and Chinese Mainland's Trade Statistics.xls
using data from File_1.xls ~ File_4.xls.
"""

import os
import re
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
xlUp            = -4162

# Number formats
FORMAT_WITH_R    = '###0.0 "(r)"'
FORMAT_WITHOUT_R = '#,##0.0  '
FORMAT_HK_VALUE  = '0.00'          # 2 decimal places for R_China(HK) D & H

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# File_1 B2 title, e.g. "（4）2026年8月进出口商品主要国别（地区）总值表（美元值）"
FILE1_MONTH_RE = re.compile(r'(\d{4})\s*年\s*(\d{1,2})\s*月')


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


def to_year(val):
    """Return a year (2000-2100) from a column A value, else None."""
    if val is None:
        return None
    try:
        y = int(float(str(val).strip()))
    except (TypeError, ValueError):
        return None
    return y if 2000 <= y <= 2100 else None


def last_used_row(ws) -> int:
    """Last used row of the sheet (no hard-coded range)."""
    last = 0
    try:
        ur = ws.UsedRange
        last = int(ur.Row) + int(ur.Rows.Count) - 1
    except Exception:
        pass
    try:
        last = max(last, int(ws.Cells(ws.Rows.Count, 2).End(xlUp).Row))
    except Exception:
        pass
    return max(last, 1)


def read_month_rows(ws):
    """
    Scan the whole sheet. Return [(row, year, month), ...] top to bottom for
    every row whose column B is a month name. Year = column A of that row or
    the nearest year above it (column A only has the year on its first month).
    year is None if no year was found above.
    """
    last = last_used_row(ws)
    try:
        block = ws.Range(ws.Cells(1, 1), ws.Cells(last, 2)).Value
        pairs = [(r[0], r[1]) for r in block]
    except Exception:
        pairs = [(ws.Cells(r, 1).Value, ws.Cells(r, 2).Value) for r in range(1, last + 1)]

    out = []
    year = None
    for row, (a_val, b_val) in enumerate(pairs, start=1):
        y = to_year(a_val)
        if y is not None:
            year = y
        month = str(b_val or "").strip()
        if month in MONTHS:
            out.append((row, year, month))
    return out


def find_latest_month_row(ws) -> int:
    rows = read_month_rows(ws)
    if not rows:
        raise ValueError(f"No month found in column B on {ws.Name}")
    return rows[-1][0]


def get_latest_year_month(ws):
    """(row, year, month) of the latest month row on the sheet."""
    rows = read_month_rows(ws)
    if not rows:
        raise ValueError(f"No month found in column B on {ws.Name}")
    return rows[-1]


def prev_year_month(year: int, month: str):
    idx = MONTHS.index(month)
    if idx == 0:
        return year - 1, "Dec"
    return year, MONTHS[idx - 1]


def insert_new_row_after_latest(ws, year: int, month: str, add_star: bool = False):
    """
    Insert a new row after the latest month and label it with the target
    month from File_1 (and the year in column A when the month is Jan).
    If add_star=True (for R_China(HK)), also write "*" into column C.
    """
    latest_row = find_latest_month_row(ws)

    ws.Rows(latest_row + 1).Insert()
    new_row = latest_row + 1
    if month == "Jan":
        ws.Cells(new_row, 1).Value = year      # column A: year on its first month
    ws.Cells(new_row, 2).Value = month

    if add_star:
        ws.Cells(new_row, 3).Value = "*"          # column C

    print(f"  [{ws.Name}] Inserted new_row = {new_row}  ({year} {month})")
    return new_row


def get_year_of_row(ws, row: int) -> int:
    for r in range(row, 0, -1):
        y = to_year(ws.Cells(r, 1).Value)
        if y is not None:
            return y
    raise ValueError(f"Cannot determine year for row {row} on {ws.Name}")


def find_month_row(ws, year: int, month: str) -> int:
    for row, y, m in read_month_rows(ws):
        if y == year and m == month:
            return row
    raise ValueError(f"Cannot find {year} {month} on sheet {ws.Name}")


def try_find_month_row(ws, year: int, month: str):
    """Like find_month_row but returns None when the month is missing."""
    try:
        return find_month_row(ws, year, month)
    except ValueError:
        return None


def parse_file1_year_month(ws_f1):
    """
    Read the target month from File_1 B2 (merged B2:K2), e.g.
    "（4）2026年8月进出口..." -> (2026, "Aug"). Returns None if not found.
    """
    texts = []
    try:
        texts.append(ws_f1.Range("B2").Value)
    except Exception:
        pass
    try:
        texts.append(ws_f1.Range("B2").MergeArea.Cells(1, 1).Value)
    except Exception:
        pass
    for col in range(2, 12):          # B2:K2
        try:
            texts.append(ws_f1.Cells(2, col).Value)
        except Exception:
            pass

    for text in texts:
        if text is None:
            continue
        m = FILE1_MONTH_RE.search(str(text))
        if m:
            year, month_num = int(m.group(1)), int(m.group(2))
            if 1 <= month_num <= 12:
                return year, MONTHS[month_num - 1]
    return None


def check_ready_to_update(sheets, target_year: int, target_month: str) -> bool:
    """
    Stop duplicate runs:
    - target month already on any sheet -> stop
    - latest month on any sheet is not the month before target -> stop
    """
    for ws in sheets:
        if try_find_month_row(ws, target_year, target_month) is not None:
            print(f"{target_year} {target_month} already exists on sheet {ws.Name}. "
                  f"Nothing to update. Stopping.")
            return False

    exp_year, exp_month = prev_year_month(target_year, target_month)
    for ws in sheets:
        rows = read_month_rows(ws)
        if not rows:
            print(f"{ws.Name}: no month rows found in column B. Stopping.")
            return False
        _, last_year, last_month = rows[-1]
        if (last_year, last_month) != (exp_year, exp_month):
            print(f"{ws.Name} latest month is {last_year} {last_month}, expected "
                  f"{exp_year} {exp_month} before {target_year} {target_month}. Stopping.")
            return False
    return True


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
        # PART 0 – Duplicate-run guard (nothing is changed before this passes)
        # ==================================================
        print("\n========== PART 0 : Check target month ==========")
        target = parse_file1_year_month(ws_f1)
        if target is None:
            print("Cannot read the year/month (yyyy年m月) from File_1 cell B2. "
                  "Nothing updated. Stopping.")
            return
        target_year, target_month = target
        print(f"  File_1 month: {target_year} {target_month}")

        if not check_ready_to_update([ws_r_china, ws_r_china_hk, ws_trade_main],
                                     target_year, target_month):
            return          # finally: workbooks closed without saving

        # ==================================================
        # PART 1 – R_China
        # ==================================================
        print("\n========== PART 1 : R_China ==========")

        # 1.1 Insert new_row
        new_row = insert_new_row_after_latest(ws_r_china, target_year, target_month)
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
        curr_year = target_year
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
        new_row_hk = insert_new_row_after_latest(ws_r_china_hk, target_year, target_month,
                                                 add_star=True)
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
        new_row_trade = insert_new_row_after_latest(ws_trade_main, target_year, target_month)
        prev_row_trade = new_row_trade - 1
        month_trade = str(ws_trade_main.Cells(new_row_trade, 2).Value or "").strip()
        print(f"  Trade Stats new_row_trade = {new_row_trade}, prev_row_trade = {prev_row_trade}")

        # 3.2 Copy row formatting from previous month onto the new row
        print("  Copying row formats from prev_row_trade → new_row_trade ...")
        ws_trade_main.Rows(prev_row_trade).Copy()
        ws_trade_main.Rows(new_row_trade).PasteSpecial(Paste=xlPasteFormats)
        excel.CutCopyMode = False
        # Re-write year (Jan only) and month name (PasteFormats may overwrite them)
        if target_month == "Jan":
            ws_trade_main.Cells(new_row_trade, 1).Value = target_year
        ws_trade_main.Cells(new_row_trade, 2).Value = month_trade

        # Sheet 3's own plain (no-(r)) format per column, read before writing
        trade_cols = [4, 6, 11, 15, 17, 19]   # D F K O Q S (M U are formulas)
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
            (ws_r_china_hk, target_new_hk, 4,  19),  # HK D → S
        ]
        for src_ws, src_row, src_col, dst_col in new_row_pairs:
            src_cell = src_ws.Cells(src_row, src_col)
            paste_trade_cell(ws_trade_main, new_row_trade, dst_col, src_cell.Value,
                             src_cell.NumberFormat, plain_fmt[dst_col], force_no_r=True)
        print("  New-month paste done (D/F/O/Q from R_China; K/S from R_China(HK)), no (r)")

        # 3.3b AutoFill formula columns M, U, W, Y (drag prev → new; relative refs adjust)
        print("  AutoFilling Trade Stats formulas M, U, W and Y ...")
        for col in [13, 21, 23, 25]:  # M, U, W, Y
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
