import sys
import os
import time
import re
import math
import heapq
import itertools
import threading
from multiprocessing import Pool, cpu_count, freeze_support

import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.styles import Alignment

# -------------------------------------------------
# 全域變數（供多進程子行程使用）
lottery_masks = []
# 適度調大以減少進程間切換；可視電腦規格再行微調
chunk_size_for_combos = 100000

# =================================================
def close_excel_workbook(file_path):
    """嘗試關閉已開啟的 Excel 活頁簿（Windows 使用 COM，自 macOS 則使用 AppleScript）"""
    if sys.platform.startswith("win"):
        try:
            import win32com.client
            excel = win32com.client.Dispatch("Excel.Application")
            fp = os.path.abspath(file_path).lower()
            for wb in list(excel.Workbooks):
                if wb.FullName.lower() == fp:
                    wb.Close(SaveChanges=True)
                    print(f"已關閉 Excel: {file_path}")
                    break
        except Exception as e:
            print("關閉 Excel 失敗：", e)
    elif sys.platform.startswith("darwin"):
        try:
            import subprocess
            abs_path = os.path.abspath(file_path)
            workbook_name = os.path.basename(abs_path)
            script = f'tell application "Microsoft Excel" to close workbook "{workbook_name}" saving yes'
            subprocess.run(["osascript", "-e", script])
            print(f"已關閉 Excel 工作簿: {file_path}")
        except Exception as e:
            print("關閉 Excel 失敗：", e)
    else:
        print("不支援的作業系統，close_excel_workbook() 略過。")


def reopen_excel_workbook(file_path):
    """嘗試重新開啟 Excel 活頁簿（Windows 使用 COM，自 macOS 則使用 AppleScript）"""
    if sys.platform.startswith("win"):
        try:
            import win32com.client
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = True
            excel.Workbooks.Open(os.path.abspath(file_path))
            print(f"已重新開啟 Excel: {file_path}")
        except Exception as e:
            print("重新開啟 Excel 失敗：", e)
    elif sys.platform.startswith("darwin"):
        try:
            import subprocess
            abs_path = os.path.abspath(file_path)
            script = f'tell application "Microsoft Excel" to open POSIX file "{abs_path}"'
            subprocess.run(["osascript", "-e", script])
            print(f"已重新開啟 Excel: {file_path}")
        except Exception as e:
            print("重新開啟 Excel 失敗：", e)
    else:
        print("不支援的作業系統，reopen_excel_workbook() 略過。")

# -------------------------------------------------
def pad_data(data, total_rows, num_columns):
    """將不足列數的資料補空字串，確保行數一致。"""
    while len(data) < total_rows:
        data.append([""] * num_columns)
    return data

# -------------------------------------------------
def init_pool(l_masks):
    """子行程初始化函式：設定全域 lottery_masks。"""
    global lottery_masks
    lottery_masks = l_masks

# -------------------------------------------------
def process_chunk(chunk_of_combos):
    max_size = 200
    heap_cnt2 = []
    heap_cnt3 = []
    heap_e4   = []

    total_draws = len(lottery_masks)

    for combo in chunk_of_combos:
        combo_mask = 0
        for n in combo:
            combo_mask |= 1 << (n - 1)

        cnt2 = cnt3 = cnt4 = cnt_exactly4 = 0
        last2 = last3 = last_exactly4 = -1

        for idx, lmask in enumerate(lottery_masks, start=1):
            match = (combo_mask & lmask).bit_count()
            if match >= 2:
                cnt2 += 1
                last2 = idx
                if match >= 3:
                    cnt3 += 1
                    last3 = idx
                    if match >= 4:
                        cnt4 += 1
                        if match == 4:
                            cnt_exactly4 += 1
                            last_exactly4 = idx

        diff2 = total_draws - last2 if last2 != -1 else total_draws
        diff3 = total_draws - last3 if last3 != -1 else total_draws
        diff_exactly4 = total_draws - last_exactly4 if last_exactly4 != -1 else total_draws

        item = (tuple(combo), cnt2, cnt3, cnt4, cnt_exactly4, diff2, diff3, diff_exactly4)

        key2 = (cnt2, cnt3, cnt4)
        if len(heap_cnt2) < max_size:
            heapq.heappush(heap_cnt2, (key2, item))
        else:
            if key2 > heap_cnt2[0][0]:
                heapq.heapreplace(heap_cnt2, (key2, item))

        key3 = (cnt3, cnt4, cnt2)
        if len(heap_cnt3) < max_size:
            heapq.heappush(heap_cnt3, (key3, item))
        else:
            if key3 > heap_cnt3[0][0]:
                heapq.heapreplace(heap_cnt3, (key3, item))

        key4 = cnt_exactly4
        if len(heap_e4) < max_size:
            heapq.heappush(heap_e4, (key4, item))
        else:
            if key4 > heap_e4[0][0]:
                heapq.heapreplace(heap_e4, (key4, item))

    return (heap_cnt2, heap_cnt3, heap_e4)

# -------------------------------------------------
def merge_heaps(all_partial_heaps):
    def merge_one_category(partial_list):
        max_size = 200
        big_heap = []
        for heap_data in partial_list:
            for (k, item) in heap_data:
                if len(big_heap) < max_size:
                    heapq.heappush(big_heap, (k, item))
                else:
                    if k > big_heap[0][0]:
                        heapq.heapreplace(big_heap, (k, item))
        out = sorted(big_heap, key=lambda x: x[0], reverse=True)
        return [itm for (k, itm) in out]

    cat2_partials = []
    cat3_partials = []
    cat4_partials = []
    for (h2, h3, h4) in all_partial_heaps:
        cat2_partials.append(h2)
        cat3_partials.append(h3)
        cat4_partials.append(h4)

    cat1 = merge_one_category(cat2_partials)
    cat2 = merge_one_category(cat3_partials)
    cat3 = merge_one_category(cat4_partials)

    return cat1, cat2, cat3

# -------------------------------------------------
def main(sheet_range_str, combo_size, file_path):
    freeze_support()

    print(f"處理檔案：{file_path}")
    print(f"讀取範圍：{sheet_range_str}，組合大小：{combo_size}")

    close_excel_workbook(file_path)
    time.sleep(0.2)

    wb = openpyxl.load_workbook(file_path, data_only=True, keep_vba=True)
    ws = wb[wb.sheetnames[0]]

    if "!" in sheet_range_str:
        _, rng = sheet_range_str.split("!", 1)
    else:
        rng = sheet_range_str
    rng = rng.replace("$", "")
    try:
        start_cell, end_cell = rng.split(":")
    except:
        print("範圍格式錯誤")
        sys.exit(1)

    sm = re.match(r"([A-Za-z]+)(\d+)?", start_cell)
    em = re.match(r"([A-Za-z]+)(\d+)?", end_cell)
    if not sm or not em:
        print("解析範圍失敗")
        sys.exit(1)

    start_row = int(sm.group(2)) if sm.group(2) else 1
    end_row   = int(em.group(2)) if em.group(2) else ws.max_row
    min_col   = column_index_from_string(sm.group(1))
    max_col   = column_index_from_string(em.group(1))

    it = ws.iter_rows(min_row=start_row, max_row=end_row,
                      min_col=min_col, max_col=max_col,
                      values_only=True)
    try:
        headers = next(it)
    except StopIteration:
        print("無可用資料")
        sys.exit(1)

    rows = list(it)
    wb.close()

    import pandas as pd
    df = pd.DataFrame(rows, columns=headers).dropna()
    draws = df.iloc[:, :combo_size].values.tolist()

    l_masks = []
    for d in draws:
        m = 0
        for num in d:
            n = int(num)
            if 1 <= n <= 39:
                m |= 1 << (n - 1)
        l_masks.append(m)

    from math import comb
    total_combos = comb(39, combo_size)
    print(f"總組合數: {total_combos}")

    all_combos = itertools.combinations(range(1, 40), combo_size)
    def chunker(iterable, n):
        batch = []
        for x in iterable:
            batch.append(x)
            if len(batch) == n:
                yield batch
                batch = []
        if batch:
            yield batch

    start_time = time.time()
    with Pool(cpu_count(), initializer=init_pool, initargs=(l_masks,)) as pool:
        partial_heaps = pool.map(
            process_chunk,
            chunker(all_combos, chunk_size_for_combos)
        )

    cat1, cat2, cat3 = merge_heaps(partial_heaps)
    end_time = time.time()
    print(f"分散處理耗時: {end_time - start_time:.2f} 秒")

    # ===== 修改排序逻辑：仅依中奖次数排序，不考虑“未开期数” =====
    cat1.sort(key=lambda t: t[1], reverse=True)  # 2星次数
    cat2.sort(key=lambda t: t[2], reverse=True)  # 3星次数
    cat3.sort(key=lambda t: t[4], reverse=True)  # 单独4星次数

    # 组织输出资料
    data_cnt2, data_cnt3, data_e4 = [], [], []
    for (c, c2, c3, c4, cE4, d2, d3, dE4) in cat1:
        data_cnt2.append(list(c) + [c2, c3, c4, d2])
    for (c, c2, c3, c4, cE4, d2, d3, dE4) in cat2:
        data_cnt3.append(list(c) + [c3, c4, d3])
    for (c, c2, c3, c4, cE4, d2, d3, dE4) in cat3:
        data_e4.append(list(c) + [cE4, dE4])

    data_cnt2 = pad_data(data_cnt2, 200, combo_size+4)
    data_cnt3 = pad_data(data_cnt3, 200, combo_size+3)
    data_e4   = pad_data(data_e4,   200, combo_size+2)

    wb2 = openpyxl.load_workbook(file_path, keep_vba=True)
    if "獲獎排列" in wb2.sheetnames:
        del wb2["獲獎排列"]
    ws_out = wb2.create_sheet("獲獎排列", index=1)

    # 写入“2星”区块
    s1 = 1
    hdr2 = [f"號碼{i}" for i in range(1, combo_size+1)] + ["2星", "3星", "4星", "未開"]
    for col, h in enumerate(hdr2, start=s1):
        cell = ws_out.cell(1, col, h)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for i, row in enumerate(data_cnt2, start=2):
        for j, v in enumerate(row, start=s1):
            cell = ws_out.cell(i, j, v)
            cell.alignment = Alignment(horizontal='center', vertical='center')

    # 写入“3星”区块
    s2 = s1 + len(hdr2) + 1
    hdr3 = [f"號碼{i}" for i in range(1, combo_size+1)] + ["3星", "4星", "未開"]
    for col, h in enumerate(hdr3, start=s2):
        cell = ws_out.cell(1, col, h)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for i, row in enumerate(data_cnt3, start=2):
        for j, v in enumerate(row, start=s2):
            cell = ws_out.cell(i, j, v)
            cell.alignment = Alignment(horizontal='center', vertical='center')

    # 写入“单独4星”区块
    s3 = s2 + len(hdr3) + 1
    hdre = [f"號碼{i}" for i in range(1, combo_size+1)] + ["4星", "未開"]
    for col, h in enumerate(hdre, start=s3):
        cell = ws_out.cell(1, col, h)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for i, row in enumerate(data_e4, start=2):
        for j, v in enumerate(row, start=s3):
            cell = ws_out.cell(i, j, v)
            cell.alignment = Alignment(horizontal='center', vertical='center')

    wb2.save(file_path)
    print("已寫入『獲獎排列』工作表（依中獎次數排序），並完成保存。")
    reopen_excel_workbook(file_path)

# =================================================
if __name__ == '__main__':
    freeze_support()
    if len(sys.argv) < 4:
        print("請輸入：<SheetRange> <combo_size> <excel_path>")
        sys.exit(1)

    sheet_range = sys.argv[1]
    try:
        combo_size = int(sys.argv[2])
    except Exception:
        print("combo_size 必須是整數")
        sys.exit(1)
    excel_path = sys.argv[3]

    # 建立簡易 GUI 顯示終端機輸出
    import tkinter as tk
    from tkinter.scrolledtext import ScrolledText

    root = tk.Tk()
    root.title("終端機內容")
    text_area = ScrolledText(root, state='normal', wrap='word', font=('Courier', 10))
    text_area.pack(expand=True, fill='both')

    class Redirector:
        def __init__(self, text_widget): self.text_widget = text_widget
        def write(self, string): self.text_widget.after(0, self.text_widget.insert, tk.END, string); self.text_widget.after(0, self.text_widget.see, tk.END)
        def flush(self): pass

    sys.stdout = Redirector(text_area)
    sys.stderr = Redirector(text_area)

    tk.Button(root, text="退出", command=root.destroy).pack(pady=5)
    def process_thread():
        main(sheet_range, combo_size, excel_path)
        print("\n處理完畢！請點選【退出】鍵結束 GUI。")
    threading.Thread(target=process_thread, daemon=True).start()

    root.mainloop()
