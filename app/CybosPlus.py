import win32com.client
import pandas as pd
from datetime import datetime
import time

class CybosAPI:
    def __init__(self):
        self.cybos = win32com.client.Dispatch("CpSysDib.StockChart")
        self.objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        self.objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        
        if self.objCpCybos.IsConnect == 0:
            print("CYBOS Plus가 연결되어 있지 않습니다.")
            exit()

    def get_stock_code(self, stock_name):
        instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
        code = instCpStockCode.NameToCode(stock_name)
        if code == "":
            print(f"종목명 '{stock_name}'에 해당하는 종목코드를 찾을 수 없습니다.")
            return None
        return code

    def create_stock_csv(self, stock_name):
        stock_code = self.get_stock_code(stock_name)
        if stock_code is None:
            return

        print(f"종목명: {stock_name}, 종목코드: {stock_code}")

        # 과거 데이터 가져오기
        self.cybos.SetInputValue(0, stock_code)
        self.cybos.SetInputValue(1, ord('2'))
        self.cybos.SetInputValue(2, datetime.now().strftime("%Y%m%d"))
        self.cybos.SetInputValue(3, datetime.now().strftime("%Y%m%d"))
        self.cybos.SetInputValue(4, 1000)
        self.cybos.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8])
        self.cybos.SetInputValue(6, ord('m'))
        self.cybos.SetInputValue(9, ord('1'))

        self.cybos.BlockRequest()

        num_data = self.cybos.GetHeaderValue(3)
        data = []
        for i in range(num_data):
            date = self.cybos.GetDataValue(0, i)
            time = self.cybos.GetDataValue(1, i)
            open_price = self.cybos.GetDataValue(2, i)
            high_price = self.cybos.GetDataValue(3, i)
            low_price = self.cybos.GetDataValue(4, i)
            close_price = self.cybos.GetDataValue(5, i)
            volume = self.cybos.GetDataValue(6, i)
            trading_value = close_price * volume
            
            data.append([date, time, open_price, high_price, low_price, close_price, trading_value])

        df = pd.DataFrame(data, columns=['날짜', '시간', '시가', '고가', '저가', '종가', '거래대금'])
        df['날짜'] = pd.to_datetime(df['날짜'].astype(str), format='%Y%m%d')
        df['시간'] = pd.to_datetime(df['시간'].astype(str).str.zfill(6), format='%H%M%S').dt.time
        df['date'] = df['날짜'].dt.strftime('%Y%m%d') + df['시간'].astype(str)
        df = df[['Date', 'Open', 'High', 'Low', 'End', 'Amount']]

        df.to_csv(f"{stock_name}.csv", index=False)
        print(f"{stock_name}.csv 파일 생성 완료")

        # 실시간 데이터 수신 시작
        self.start_realtime_data(stock_code, stock_name)

    def start_realtime_data(self, stock_code, stock_name):
        handler = RealtimeDataHandler(stock_code, stock_name)
        handler.start()

    def update_csv_files(self, stock_list):
        for stock_name in stock_list:
            self.create_stock_csv(stock_name)

class RealtimeDataHandler:
    def __init__(self, stock_code, stock_name):
        self.stock_code = stock_code
        self.stock_name = stock_name
        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        self.objStockCur.SetInputValue(0, stock_code)

    def start(self):
        handler = win32com.client.WithEvents(self.objStockCur, self)
        self.objStockCur.Subscribe()
        print(f"{self.stock_name} 실시간 데이터 수신 시작")

    def OnReceived(self):
        current_time = datetime.now().strftime("%Y%m%d%H%M%S")
        current_price = self.objStockCur.GetHeaderValue(13)  # 현재가
        open_price = self.objStockCur.GetHeaderValue(4)  # 시가
        high_price = self.objStockCur.GetHeaderValue(5)  # 고가
        low_price = self.objStockCur.GetHeaderValue(6)  # 저가
        volume = self.objStockCur.GetHeaderValue(9)  # 거래량
        trading_value = current_price * volume  # 거래대금 계산

        # CSV 파일에 실시간 데이터 추가
        with open(f"{self.stock_name}.csv", "a") as f:
            f.write(f"{current_time},{open_price},{high_price},{low_price},{current_price},{trading_value}\n")

        print(f"{self.stock_name} 실시간 데이터 추가: {current_time}, 시가: {open_price}, 고가: {high_price}, 저가: {low_price}, 현재가: {current_price}, 거래대금: {trading_value}")

    def stop(self):
        self.objStockCur.Unsubscribe()
        print(f"{self.stock_name} 실시간 데이터 수신 중지")