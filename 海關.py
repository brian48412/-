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


def paste_formula(src_cell, dest_cell):
    dest_cell.Formula = src_cell.Formula


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

        # 1.3 Copy formulas L:O
        print("  Copying formulas L:O ...")
        for col in range(12, 16):
            paste_formula(ws_r_china.Cells(prev_row, col), ws_r_china.Cells(new_row, col))

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

        # File_1 → new_row (2 decimal places, no (r))
        print("  Writing File_1 values into new_row ...")
        v = safe_float(ws_f1.Range("E21").Value)
        if v is not None:
            cell = ws_r_china_hk.Cells(new_row_hk, 8)          # H
            cell.Value = v / 1000
            cell.NumberFormat = FORMAT_HK_VALUE

        v = safe_float(ws_f1.Range("G21").Value)
        if v is not None:
            cell = ws_r_china_hk.Cells(new_row_hk, 4)          # D
            cell.Value = v / 1000
            cell.NumberFormat = FORMAT_HK_VALUE

        # AutoFill formulas for F, J, L, N  (all four columns)
        print("  Filling formulas F / J / L / N (AutoFill) ...")
        for col in [6, 10, 12, 14]:          # F, J, L, N
            src = ws_r_china_hk.Cells(prev_row_hk, col)
            dest_range = ws_r_china_hk.Range(
                ws_r_china_hk.Cells(prev_row_hk, col),
                ws_r_china_hk.Cells(new_row_hk, col)
            )
            src.AutoFill(Destination=dest_range)

        # File_2 → new_row(-1) (2 decimal places)
        print("  Writing File_2 values into new_row(-1) ...")
        v = safe_float(ws_f2.Range("E17").Value)
        if v is not None:
            cell = ws_r_china_hk.Cells(prev_row_hk, 8)         # H
            cell.Value = v / 1000000
            cell.NumberFormat = FORMAT_HK_VALUE

        v = safe_float(ws_f2.Range("G17").Value)
        if v is not None:
            cell = ws_r_china_hk.Cells(prev_row_hk, 4)         # D
            cell.Value = v / 1000000
            cell.NumberFormat = FORMAT_HK_VALUE

        # ==================================================
        # PART 3 – Trade Statistics
        # ==================================================
        print("\n========== PART 3 : Trade Statistics ==========")

        new_row_trade = insert_new_row_after_latest(ws_trade_main, 400, 411)
        prev_row_trade = new_row_trade - 1

        # Copy coloured cells (value + number format) from R_China
        print("  Copying coloured cells from R_China ...")
        src_start_row = last_year_jan_row
        src_end_row   = new_row
        for row in range(src_start_row, src_end_row + 1):
            for col in [4, 6, 8, 10]:
                src = ws_r_china.Cells(row, col)
                dst = ws_trade_main.Cells(row, col)
                dst.Value = src.Value
                dst.NumberFormat = src.NumberFormat

        # Four special cells from R_China(HK)
        print("  Copying 4 cells from R_China(HK) ...")
        ws_trade_main.Cells(prev_row_trade, 19).Value = ws_r_china_hk.Cells(prev_row_hk, 4).Value   # S
        ws_trade_main.Cells(new_row_trade, 19).Value  = ws_r_china_hk.Cells(new_row_hk, 4).Value
        ws_trade_main.Cells(prev_row_trade, 11).Value = ws_r_china_hk.Cells(prev_row_hk, 8).Value   # K
        ws_trade_main.Cells(new_row_trade, 11).Value  = ws_r_china_hk.Cells(new_row_hk, 8).Value

        # Save
        print("\nSaving workbooks ...")
        wb_china.Save()
        wb_trade.Save()
        print("All done successfully.")

    finally:
        close_all(excel, wb_china, wb_f1, wb_f2, wb_f3, wb_f4, wb_trade)


if __name__ == "__main__":
    main()