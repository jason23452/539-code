import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from openpyxl import Workbook

def fetch_all_td_rows(p1: str, p2: str) -> list[list[str]]:
    """
    抓取查詢結果中每一列的所有 <td> 內容，回傳為 list of list。
    """
    base_url = 'http://www.9800.com.tw'
    entry_url = urljoin(base_url, '/head.asp')

    with requests.Session() as session:
        # Step 1: GET 表單頁面
        resp = session.get(entry_url)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, 'html.parser')

        form = soup.find('form', attrs={'name': 'search'})
        if not form or not form.get('action'):
            raise RuntimeError("❌ 無法找到 name='search' 的 <form> 或其 action")

        post_url = urljoin(base_url, form['action'])

        # Step 2: 準備 POST payload
        payload = {'p1': p1, 'p2': p2}
        for hidden in form.find_all('input', type='hidden'):
            name = hidden.get('name')
            if name and name not in payload:
                payload[name] = hidden.get('value', '')

        # Step 3: 發送 POST 請求
        headers = {'Referer': entry_url}
        post_resp = session.post(post_url, data=payload, headers=headers)
        post_resp.raise_for_status()
        post_soup = BeautifulSoup(post_resp.text, 'html.parser')

        # Step 4: 抓取每一列中的所有 <td>
        rows = []
        for tr in post_soup.find_all('tr'):
            td_values = [td.get_text(strip=True) for td in tr.find_all('td')]
            if td_values:
                rows.append(td_values)

        return rows


def write_rows_to_excel(rows: list[list[str]], filename: str = 'td_results.xlsx') -> None:
    """
    將 2D list 寫入 Excel，每列對應一列資料。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = 'all_td'

    for row_index, row in enumerate(rows, start=1):
        for col_index, value in enumerate(row, start=1):
            ws.cell(row=row_index, column=col_index, value=value)

    wb.save(filename)
    print(f'✅ 所有 <td> 資料已儲存到：{filename}')


if __name__ == '__main__':
    p1_value = '096001'
    p2_value = '114100'

    try:
        all_td_data = fetch_all_td_rows(p1_value, p2_value)
        if all_td_data:
            print("✅ 抓到以下 <td> 資料列：")
            for row in all_td_data:
                print(" -", row)
            write_rows_to_excel(all_td_data)
        else:
            print("⚠️ 沒有抓到任何 <td> 資料")
    except Exception as err:
        print("❌ 發生錯誤：", err)
