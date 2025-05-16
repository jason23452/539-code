#!/usr/bin/env python3
# backtest_sections.py

import sys
import os
import re
import subprocess
import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.styles import Alignment

def close_excel_workbook(file_path):
    """嘗試關閉已開啟的 Excel 活頁簿（Windows 使用 COM，自 macOS 使用 AppleScript）"""
    if sys.platform.startswith('win'):
        try:
            import win32com.client
            excel = win32com.client.Dispatch('Excel.Application')
            fp = os.path.abspath(file_path).lower()
            for wb in list(excel.Workbooks):
                if wb.FullName.lower() == fp:
                    wb.Close(SaveChanges=True)
                    print(f'已關閉 Excel: {file_path}')
                    break
        except Exception as e:
            print('關閉 Excel 失敗：', e)
    elif sys.platform.startswith('darwin'):
        try:
            name = os.path.basename(os.path.abspath(file_path))
            script = f'tell application "Microsoft Excel" to close workbook "{name}" saving yes'
            subprocess.run(['osascript', '-e', script], check=False)
            print(f'已關閉 Excel: {file_path}')
        except Exception as e:
            print('關閉 Excel 失敗：', e)
    else:
        print('不支援的作業系統，略過關閉。')

def reopen_excel_workbook(file_path):
    """嘗試重新開啟 Excel 活頁簿"""
    if sys.platform.startswith('win'):
        try:
            import win32com.client
            excel = win32com.client.Dispatch('Excel.Application')
            excel.Visible = True
            excel.Workbooks.Open(os.path.abspath(file_path))
            print(f'已重新開啟 Excel: {file_path}')
        except Exception as e:
            print('重新開啟 Excel 失敗：', e)
    elif sys.platform.startswith('darwin'):
        try:
            script = f'tell application "Microsoft Excel" to open POSIX file "{os.path.abspath(file_path)}"'
            subprocess.run(['osascript', '-e', script], check=False)
            print(f'已重新開啟 Excel: {file_path}')
        except Exception as e:
            print('重新開啟 Excel 失敗：', e)
    else:
        print('不支援的作業系統，略過開啟。')

def detect_combo_size(ws):
    """偵測每組號碼長度 M"""
    size = 0
    while ws.cell(1, size+1).value == f"號碼{size+1}":
        size += 1
    return size

def read_draws(ws, col_range):
    """讀取原始開獎資料 draws，支援 B2:F100 或 B:F 格式"""
    m = re.match(r'^([A-Za-z]+)(\d*):([A-Za-z]+)(\d*)$', col_range)
    if not m:
        raise ValueError(f"不支援的欄位範圍格式：{col_range}")
    c1_letter, r1_str, c2_letter, r2_str = m.groups()
    c1 = column_index_from_string(c1_letter)
    c2 = column_index_from_string(c2_letter)
    r1 = int(r1_str) if r1_str else 2
    r2 = int(r2_str) if r2_str else ws.max_row

    draws = []
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2, values_only=True):
        if any(row):
            draws.append([int(v) for v in row if v is not None])
    return draws

def read_section_combos(ws, start_col, combo_size):
    """從「獲獎排列」區段讀取 combos"""
    combos = []
    for row in ws.iter_rows(min_row=2, min_col=start_col,
                            max_col=start_col+combo_size-1,
                            values_only=True):
        if any(row):
            combos.append(tuple(int(v) for v in row))
    return combos

def count_hits(draws, combo, threshold, exact=False):
    """計算符合條件的期數數量"""
    cnt = 0
    for d in draws:
        hits = len(set(combo) & set(d))
        if (exact and hits == threshold) or (not exact and hits >= threshold):
            cnt += 1
    return cnt

def write_section(ws, combo_size, draws, combos, thresholds, start_col):
    """寫回測結果到工作表，不留空列：headers 寫第1列，combos 從第2列開始"""
    headers = [f"號碼{i}" for i in range(1, combo_size+1)] + [name for name,_,_ in thresholds]
    for j, h in enumerate(headers, start=start_col):
        ws.cell(1, j, h).alignment = Alignment(horizontal="center")

    for i, combo in enumerate(combos, start=2):
        for k, num in enumerate(combo, start=start_col):
            ws.cell(i, k, num).alignment = Alignment(horizontal="center")
        for idx, (_, thr, exact) in enumerate(thresholds, start=1):
            cnt = count_hits(draws, combo, thr, exact)
            col = start_col + combo_size + idx - 1
            ws.cell(i, col, cnt).alignment = Alignment(horizontal="center")

def main(path, draws_sheet, col_range, prize_sheet):
    if not os.path.exists(path):
        print("找不到檔案")
        return

    close_excel_workbook(path)

    wb = openpyxl.load_workbook(path, keep_vba=True)
    ws_draws = wb[draws_sheet]
    ws_prize = wb[prize_sheet]

    M = detect_combo_size(ws_prize)
    print(f"偵測到組合長度 M = {M}")

    draws = read_draws(ws_draws, col_range)

    # 自動偵測每個「號碼1」區段起始欄位
    start_cols = [idx for idx, cell in enumerate(ws_prize[1], start=1) if cell.value == '號碼1']
    if len(start_cols) < 3:
        print("無法偵測到三個組合區段起始欄位，請檢查工作表格式。")
        return
    start2, start3, start4 = start_cols[:3]

    combos2 = read_section_combos(ws_prize, start_col=start2, combo_size=M)
    combos3 = read_section_combos(ws_prize, start_col=start3, combo_size=M)
    combos4 = read_section_combos(ws_prize, start_col=start4, combo_size=M)

    if "回測結果" in wb.sheetnames:
        del wb["回測結果"]
    ws = wb.create_sheet("回測結果", index=len(wb.sheetnames))

    write_section(ws, M, draws, combos2,
                  thresholds=[("2星次數",2,False),("3星次數",3,False),("4星次數",4,True)],
                  start_col=start2)
    write_section(ws, M, draws, combos3,
                  thresholds=[("3星次數",3,False),("4星次數",4,True)],
                  start_col=start3)
    write_section(ws, M, draws, combos4,
                  thresholds=[("4星次數",4,True)],
                  start_col=start4)

    wb.save(path)
    print("已將三部分回測結果寫入「回測結果」工作表。")

    reopen_excel_workbook(path)

if __name__ == "__main__":
    if len(sys.argv) != 5:
        print("用法：python backtest_sections.py <檔案> <原始表:球號> <欄位範圍如B2:F100或B:F> <獲獎排列表>")
        sys.exit(1)
    _, path, draws_sheet, col_range, prize_sheet = sys.argv
    main(path, draws_sheet, col_range, prize_sheet)
