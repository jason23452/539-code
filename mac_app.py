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
# 全域變數
top_n = 200             # 要保留的最佳組合數
max_gap_limit = 1000000  # 最大相鄰中獎期距閾值
lottery_masks = []      # 子行程初始化後存放遮罩列表
chunk_size_for_combos = 100000
# -------------------------------------------------

def close_excel_workbook(file_path):
    '''嘗試關閉已開啟的 Excel 活頁簿（Windows 使用 COM，自 macOS 使用 AppleScript）'''
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
            import subprocess
            name = os.path.basename(os.path.abspath(file_path))
            script = f'tell application "Microsoft Excel" to close workbook "{name}" saving yes'
            subprocess.run(['osascript', '-e', script])
            print(f'已關閉 Excel: {file_path}')
        except Exception as e:
            print('關閉 Excel 失敗：', e)
    else:
        print('不支援的作業系統，略過關閉。')

def reopen_excel_workbook(file_path):
    '''嘗試重新開啟 Excel 活頁簿'''
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
            import subprocess
            script = f'tell application "Microsoft Excel" to open POSIX file "{os.path.abspath(file_path)}"'
            subprocess.run(['osascript', '-e', script])
            print(f'已重新開啟 Excel: {file_path}')
        except Exception as e:
            print('重新開啟 Excel 失敗：', e)
    else:
        print('不支援的作業系統，略過開啟。')

def pad_data(data, total_rows, num_columns):
    '''補足列數以對齊輸出表格'''
    while len(data) < total_rows:
        data.append([''] * num_columns)
    return data

def init_pool(l_masks):
    '''子行程初始化函式：設定全域 lottery_masks。'''
    global lottery_masks
    lottery_masks = l_masks

def process_chunk(chunk_of_combos):
    '''多進程計算：維護 2星／3星／恰4星 的 top_n 小堆'''
    heap2 = []
    heap3 = []
    heap4 = []
    total = len(lottery_masks)

    for combo in chunk_of_combos:
        m = 0
        for n in combo:
            m |= 1 << (n - 1)

        cnt2 = cnt3 = cnt4 = cntE4 = 0
        last2 = last3 = lastE4 = -1

        for idx, lm in enumerate(lottery_masks, start=1):
            matches = (m & lm).bit_count()
            if matches >= 2:
                cnt2 += 1
                last2 = idx
            if matches >= 3:
                cnt3 += 1
                last3 = idx
            if matches >= 4:
                cnt4 += 1
                if matches == 4:
                    cntE4 += 1
                    lastE4 = idx

        diff2 = total - last2 if last2 != -1 else total
        diff3 = total - last3 if last3 != -1 else total
        diffE4 = total - lastE4 if lastE4 != -1 else total

        item = (tuple(combo), cnt2, cnt3, cnt4, cntE4, diff2, diff3, diffE4)

        key2 = (cnt2, cnt3, cnt4)
        if len(heap2) < top_n:
            heapq.heappush(heap2, (key2, item))
        elif key2 > heap2[0][0]:
            heapq.heapreplace(heap2, (key2, item))

        key3 = (cnt3, cnt4, cnt2)
        if len(heap3) < top_n:
            heapq.heappush(heap3, (key3, item))
        elif key3 > heap3[0][0]:
            heapq.heapreplace(heap3, (key3, item))

        key4 = cntE4
        if len(heap4) < top_n:
            heapq.heappush(heap4, (key4, item))
        elif key4 > heap4[0][0]:
            heapq.heapreplace(heap4, (key4, item))

    return heap2, heap3, heap4

def merge_heaps(all_heaps):
    '''合併各子行程回傳的局部堆，並回傳三大類完整排序列表'''
    def merge_one(partials):
        big = []
        for heap_data in partials:
            for k, item in heap_data:
                if len(big) < top_n:
                    heapq.heappush(big, (k, item))
                elif k > big[0][0]:
                    heapq.heapreplace(big, (k, item))
        return [itm for _, itm in sorted(big, key=lambda x: x[0], reverse=True)]

    h2 = [h2 for h2, _h3, _h4 in all_heaps]
    h3 = [_h3 for _h2, _h3, _h4 in all_heaps]
    h4 = [_h4 for _h2, _h3, _h4 in all_heaps]

    return merge_one(h2), merge_one(h3), merge_one(h4)

def compute_max_gap(combo, masks, threshold, exact=False):
    '''計算最大相鄰命中期距'''
    m = 0
    for n in combo:
        m |= 1 << (n - 1)
    prev = 0
    max_gap = 0
    total = len(masks)
    for idx, lm in enumerate(masks, start=1):
        matches = (m & lm).bit_count()
        ok = (matches == threshold) if exact else (matches >= threshold)
        if ok:
            gap = idx - prev
            if gap > max_gap:
                max_gap = gap
            prev = idx
    final_gap = total + 1 - prev
    return final_gap if final_gap > max_gap else max_gap

def main(sheet_range, combo_size, file_path):
    freeze_support()
    print(f'檔案：{file_path}，範圍：{sheet_range}，組合大小：{combo_size}')
    print(f'top_n={top_n}，max_gap_limit={max_gap_limit}')

    close_excel_workbook(file_path)
    time.sleep(0.2)

    wb = openpyxl.load_workbook(file_path, data_only=True, keep_vba=True)
    ws = wb[wb.sheetnames[0]]
    rng = sheet_range.split('!',1)[-1].replace('$','')
    try:
        sc, ec = rng.split(':')
    except:
        print('範圍格式錯誤')
        sys.exit(1)
    sm = re.match(r'([A-Za-z]+)(\d+)?', sc)
    em = re.match(r'([A-Za-z]+)(\d+)?', ec)
    if not sm or not em:
        print('解析範圍失敗')
        sys.exit(1)
    sr = int(sm.group(2) or 1)
    er = int(em.group(2) or ws.max_row)
    c1 = column_index_from_string(sm.group(1))
    c2 = column_index_from_string(em.group(1))
    it = ws.iter_rows(min_row=sr, max_row=er, min_col=c1, max_col=c2, values_only=True)
    try:
        headers = next(it)
    except StopIteration:
        print('無資料')
        sys.exit(1)
    rows = list(it)
    wb.close()

    import pandas as pd
    df = pd.DataFrame(rows, columns=headers).dropna()
    draws = df.iloc[:, :combo_size].astype(int).values.tolist()

    masks = []
    for nums in draws:
        mm = 0
        for v in nums:
            if 1 <= v <= 39:
                mm |= 1 << (v - 1)
        masks.append(mm)

    from math import comb
    total = comb(39, combo_size)
    print(f'總組合數：{total}')
    all_combos = itertools.combinations(range(1,40), combo_size)

    def chunker(it, size):
        buf = []
        for x in it:
            buf.append(x)
            if len(buf) == size:
                yield buf
                buf = []
        if buf:
            yield buf

    t0 = time.time()
    with Pool(cpu_count(), initializer=init_pool, initargs=(masks,)) as pool:
        partials = pool.map(process_chunk, chunker(all_combos, chunk_size_for_combos))
    cat1, cat2, cat3 = merge_heaps(partials)
    print(f'分散處理耗時：{time.time()-t0:.2f}s')

    cat1 = [itm for itm in cat1 if compute_max_gap(itm[0], masks, threshold=2) <= max_gap_limit][:top_n]
    cat2 = [itm for itm in cat2 if compute_max_gap(itm[0], masks, threshold=3) <= max_gap_limit][:top_n]
    cat3 = [itm for itm in cat3 if compute_max_gap(itm[0], masks, threshold=4, exact=True) <= max_gap_limit][:top_n]

    data2 = [list(c)+[c2,c3,c4,d2] for (c,c2,c3,c4,cE4,d2,d3,dE4) in cat1]
    data3 = [list(c)+[c3,c4,d3] for (c,c2,c3,c4,cE4,d2,d3,dE4) in cat2]
    data4 = [list(c)+[cE4,dE4] for (c,c2,c3,c4,cE4,d2,d3,dE4) in cat3]

    data2 = pad_data(data2, top_n, combo_size+4)
    data3 = pad_data(data3, top_n, combo_size+3)
    data4 = pad_data(data4, top_n, combo_size+2)

    wb2 = openpyxl.load_workbook(file_path, keep_vba=True)
    if '獲獎排列' in wb2.sheetnames:
        del wb2['獲獎排列']
    ws_out = wb2.create_sheet('獲獎排列', index=1)

    col = 1
    hdr2 = [f'號碼{i}' for i in range(1, combo_size+1)] + ['2星','3星','4星','未開']
    for j, h in enumerate(hdr2, start=col):
        ws_out.cell(1,j,h).alignment = Alignment('center','center')
    for i, row in enumerate(data2, start=2):
        for j, v in enumerate(row, start=col):
            ws_out.cell(i,j,v).alignment = Alignment('center','center')

    col += len(hdr2) + 1
    hdr3 = [f'號碼{i}' for i in range(1, combo_size+1)] + ['3星','4星','未開']
    for j, h in enumerate(hdr3, start=col):
        ws_out.cell(1,j,h).alignment = Alignment('center','center')
    for i, row in enumerate(data3, start=2):
        for j, v in enumerate(row, start=col):
            ws_out.cell(i,j,v).alignment = Alignment('center','center')

    col += len(hdr3) + 1
    hdr4 = [f'號碼{i}' for i in range(1, combo_size+1)] + ['4星','未開']
    for j, h in enumerate(hdr4, start=col):
        ws_out.cell(1,j,h).alignment = Alignment('center','center')
    for i, row in enumerate(data4, start=2):
        for j, v in enumerate(row, start=col):
            ws_out.cell(i,j,v).alignment = Alignment('center','center')

    wb2.save(file_path)
    print('已寫入「獲獎排列」並保存完成。')
    reopen_excel_workbook(file_path)

if __name__ == '__main__':
    freeze_support()
    if len(sys.argv) < 4:
        print('用法：<SheetRange> <combo_size> <excel_path> [top_n] [max_gap_limit]')
        sys.exit(1)

    sheet_range = sys.argv[1]
    combo_size = int(sys.argv[2])
    excel_path = sys.argv[3]

    if len(sys.argv) >= 5:
        try:
            top_n = int(sys.argv[4])
        except ValueError:
            print('第4个参数 top_n 必须是整数')
            sys.exit(1)

    if len(sys.argv) >= 6:
        try:
            max_gap_limit = int(sys.argv[5])
        except ValueError:
            print('第5个参数 max_gap_limit 必须是整数')
            sys.exit(1)

    print(f'已设定 top_n={top_n}, max_gap_limit={max_gap_limit}')

    import tkinter as tk
    from tkinter.scrolledtext import ScrolledText
    root = tk.Tk()
    root.title('計算終端')
    ta = ScrolledText(root, wrap='word', font=('Consolas',10))
    ta.pack(expand=True, fill='both')
    class R:
        def __init__(self,w): self.w = w
        def write(self,s): self.w.after(0,self.w.insert,tk.END,s); self.w.after(0,self.w.see,tk.END)
        def flush(self): pass
    sys.stdout = R(ta)
    sys.stderr = R(ta)
    threading.Thread(target=lambda: main(sheet_range, combo_size, excel_path), daemon=True).start()
    tk.Button(root, text='退出', command=root.destroy).pack(pady=5)
    root.mainloop()
