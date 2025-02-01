import openpyxl
from openpyxl.styles import Alignment
from datetime import date
from openpyxl import load_workbook
import requests
import time
from openpyxl.utils.exceptions import InvalidFileException

request_url = "https://query1.finance.yahoo.com/v8/finance/chart/"
# request_stock_code = ["%5EHSI", "%5EIXIC", "9988.HK", "1810.HK", "2318.HK", "601318.SS", "002602.SZ", "600276.SS", "000001.SZ", "601127.SS", "002594.SZ", "300750.SZ",
# "002230.SZ", "600436.SS", "000850.SZ", "601111.SS", "000001.SS", "399001.SZ", "399006.SZ"]
request_stock_code = ["%5EHSI", "9988.HK", "1810.HK", "2318.HK", "601318.SS", "002602.SZ", "600276.SS", "000001.SZ",
                      "601127.SS", "002594.SZ", "300750.SZ",
                      "002230.SZ", "600436.SS", "000850.SZ", "601111.SS", "000001.SS", "399001.SZ", "399006.SZ"]

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) ' \
                  'AppleWebKit/537.36 (KHTML, like Gecko) ' \
                  'Chrome/58.0.3029.110 Safari/537.3'
}


def get_stock_data_from_excel(file_path):
    workbook = load_workbook(file_path)
    sheet = workbook.active
    stock_data = {}

    for i in range(1, sheet.max_row + 1, 8):
        stock_name = sheet.cell(row=i, column=4).value  # 股票名在D列
        if not stock_name:
            continue

        closing_price = sheet.cell(row=i, column=4).value
        low_price = sheet.cell(row=i + 4, column=4).value
        high_price = sheet.cell(row=i + 5, column=4).value
        change = sheet.cell(row=i + 7, column=4).value

        if closing_price and low_price and high_price:
            stock_data[stock_name] = {
                'closing_price': round(float(closing_price), 2),
                'low_price': round(float(low_price), 2),
                'high_price': round(float(high_price), 2),
                'change': change
            }

    return stock_data


def get_stock_data_today():
    '''
    需要的参数：regularMarketPrice, regularMarketDayHigh, regularMarketDayLow， regularMarketChangePercent 只要raw格式的数值，regularMarketChangePercent改成百分比格式，保留小数点后两位例如 13.51%
    {
    "chart": {
        "result": [
            {
                "meta": {
                    "currency": "HKD",
                    "symbol": "0020.HK",
                    ...
                    "regularMarketPrice": 1.61,
                    "fiftyTwoWeekHigh": 2.35,
                    "fiftyTwoWeekLow": 0.58,
                    "regularMarketDayHigh": 1.66,
                    "regularMarketDayLow": 1.59,
                    "regularMarketVolume": 257744223,
                    "longName": "SENSETIME-W",
                    "shortName": "SENSETIME-W",
                    "chartPreviousClose": 1.63,
                    "previousClose": 1.63,
                    "scale": 3,
                    "priceHint": 3,
    '''
    stock_data = {}

    for code in request_stock_code:
        if not code:
            continue
        url = f"{request_url}{code}"
        try:
            response = requests.get(url, headers=headers)
            if response.status_code != 200:
                if response.status_code == 429:
                    retryAfter = response.headers.get('Retry-After', 60)
                    print(f"请求股票{code}失败，状态码：{response.status_code}，重试时间：{retryAfter}秒")
                    time.sleep(int(retryAfter))
                    response = requests.get(url, headers=headers)
                    if response.status_code != 200:
                        print(f"请求股票{code}失败，状态码：{response.status_code}")
                        continue
                else:
                    print(f"请求股票{code}失败，状态码：{response.status_code}")
                    continue
            data = response.json()
            try:
                result = data['chart']['result'][0]
                meta = result['meta']
                stock_name = meta.get('longName', meta.get('shortName', code))
                regularMarketPrice = meta['regularMarketPrice']
                regularMarketDayHigh = meta['regularMarketDayHigh']
                regularMarketDayLow = meta['regularMarketDayLow']
                previousClose = meta['previousClose']
                stock_data[stock_name] = {
                    'closing_price': round(regularMarketPrice, 2),
                    'low_price': round(regularMarketDayLow, 2),
                    'high_price': round(regularMarketDayHigh, 2),
                    'previous_close': round(previousClose, 2)
                }
            except (KeyError, IndexError):
                print(f"请求股票{code}失败，数据格式错误")
                continue
        except requests.exceptions.RequestException as e:
            print(f"请求股票{code}时发生错误: {e}")
            continue

    return stock_data


def calculate_change(prev_price, current_price):
    if prev_price:
        return round((current_price - prev_price) / prev_price * 100, 2)
    return ""


def update_excel(filename, stock_data):
    '''
    开头为
    1. 序号：= 行号 - 1
    2. 日期：= 今天日期 格式 `2025/01/01`
    3. 星期：= 今天星期几，如果星期五，则写5
    EXCEL 每个股票一共8列(无需名字，由request_stock_code决定顺序)
    1. 收盘价：表头表示为股票名字，对应 regularMarketPrice
    2. 日k：无需填写
    3. 周k：无需填写
    4. 月k：无需填写
    5. 低点：对应 regularMarketDayLow
    6. 高点：对应 regularMarketDayHigh
    7. 交易数量：无需填写
    8. 涨幅：对应 (previousClose - regularMarketPrice) / previousClose * 100
    '''

    try:
        wb = load_workbook(filename)
        sheet = wb.active
    except (FileNotFoundError, InvalidFileException):
        wb = openpyxl.Workbook()
        sheet = wb.active
        # 创建表头
        headers = ["序号", "日期", "星期"]
        for code in request_stock_code:
            headers.extend([
                f"{code} 收盘价",
                f"{code} 日k",
                f"{code} 周k",
                f"{code} 月k",
                f"{code} 低点",
                f"{code} 高点",
                f"{code} 交易数量",
                f"{code} 涨幅"
            ])
        sheet.append(headers)

    today = date.today()
    weekday_num = today.weekday() + 1
    weekday = 5 if weekday_num == 5 else weekday_num

    last_row = sheet.max_row + 1

    # 写入序号、日期、星期
    sheet.cell(row=last_row, column=1, value=last_row - 1)
    sheet.cell(row=last_row, column=2, value=today.strftime("%Y/%m/%d"))
    sheet.cell(row=last_row, column=3, value=weekday)

    for i, (stock_name, data) in enumerate(stock_data.items(), start=1):
        base_col = 4 + (i - 1) * 8
        # 写入收盘价
        sheet.cell(row=last_row, column=base_col, value=data['closing_price'])
        # 日k留空
        # 周k留空
        # 月k留空
        # 写入低点
        sheet.cell(row=last_row, column=base_col + 4, value=data['low_price'])
        # 写入高点
        sheet.cell(row=last_row, column=base_col + 5, value=data['high_price'])
        # 交易数量留空
        # 计算涨幅
        change = ""
        if data['previous_close']:
            change_value = (data['previous_close'] - data['closing_price']) / data['previous_close'] * 100
            change = f"{change_value:.2f}%"
        sheet.cell(row=last_row, column=base_col + 7, value=change)

    # 调整对齐
    for cell in sheet[last_row]:
        cell.alignment = Alignment(horizontal='center')

    wb.save(filename)
    print(f"数据已成功更新到 {filename}")


def main():
    filename = "test.xlsx"  # 请确保这个文件已经存在，并且有正确的表头
    try:
        stock_data = get_stock_data_today()
        update_excel(filename, stock_data)
    except Exception as e:
        print(f"更新股票数据时发生错误: {e}")


if __name__ == "__main__":
    main()
