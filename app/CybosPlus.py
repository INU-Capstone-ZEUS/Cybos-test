import win32com.client
import pandas as pd
from datetime import datetime, timedelta

def connect_cybos():
    cpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
    stCpCybos = cpCybos.IsConnect
    if stCpCybos == 0:
        print("CYBOS Plus가 연결되어 있지 않습니다.")
        return False
    return True

def get_stock_code(stock_name):
    instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    codeList = instCpCodeMgr.GetStockListByMarket(1)  # 거래소
    codeList2 = instCpCodeMgr.GetStockListByMarket(2)  # 코스닥
    
    for code in codeList + codeList2:
        if stock_name == instCpCodeMgr.CodeToName(code):
            return code
    return None

def get_minute_data(code, start_date, end_date):
    instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
    instStockChart.SetInputValue(0, code)
    instStockChart.SetInputValue(1, ord('1'))  # 기간으로 요청
    instStockChart.SetInputValue(2, end_date)  # 종료일
    instStockChart.SetInputValue(3, start_date)  # 시작일
    instStockChart.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8])  # 날짜, 시간, 시가, 고가, 저가, 종가, 거래량
    instStockChart.SetInputValue(6, ord('m'))  # 분 단위
    instStockChart.SetInputValue(9, ord('1'))  # 수정주가 사용

    data = []
    while True:
        instStockChart.BlockRequest()
        numData = instStockChart.GetHeaderValue(3)
        if numData == 0:
            break

        for i in range(numData):
            date = instStockChart.GetDataValue(0, i)
            time = instStockChart.GetDataValue(1, i)
            open_price = instStockChart.GetDataValue(2, i)
            high_price = instStockChart.GetDataValue(3, i)
            low_price = instStockChart.GetDataValue(4, i)
            close_price = instStockChart.GetDataValue(5, i)
            volume = instStockChart.GetDataValue(6, i)
            data.append([date, time, open_price, high_price, low_price, close_price, volume])

        if instStockChart.Continue == 0:
            break

    df = pd.DataFrame(data, columns=['Date', 'Time', 'Open', 'High', 'Low', 'Close', 'Volume'])
    df['Date'] = pd.to_datetime(df['Date'].astype(str), format='%Y%m%d')
    df['Time'] = pd.to_datetime(df['Time'].astype(str).str.zfill(6), format='%H%M%S').dt.time
    return df

def main():
    if not connect_cybos():
        return

    stock_name = input("종목명을 입력하세요: ")
    code = get_stock_code(stock_name)
    if code is None:
        print("해당 종목을 찾을 수 없습니다.")
        return

    end_date = datetime.now().strftime("%Y%m%d")
    start_date = (datetime.now() - timedelta(days=730)).strftime("%Y%m%d")  # 2년 전

    print(f"{stock_name}({code})의 분봉 데이터를 가져오는 중...")
    df = get_minute_data(code, start_date, end_date)

    file_name = f"{stock_name}_{code}_minute_data.csv"
    df.to_csv(file_name, index=False)
    print(f"데이터가 {file_name}에 저장되었습니다.")

if __name__ == "__main__":
    main()