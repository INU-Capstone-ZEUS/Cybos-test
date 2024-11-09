import win32com.client
import pandas as pd
from datetime import datetime

class CybosAPI:
    def __init__(self):
        self.cybos = win32com.client.Dispatch("CpSysDib.StockChart")
        self.objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        self.objCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
        
        if self.objCpCybos.IsConnect == 0:
            print("CYBOS Plus가 연결되어 있지 않습니다.")
            sys.exit()

    def get_stock_code(self, stock_name):
        code = self.objCpStockCode.NameToCode(stock_name)
        if code == "":
            print(f"종목명 '{stock_name}'에 해당하는 종목코드를 찾을 수 없습니다.")
            return None
        return code

    def create_stock_csv(self, stock_name):
        stock_code = self.get_stock_code(stock_name)
        if stock_code is None:
            return

        print(f"종목명: {stock_name}, 종목코드: {stock_code}")

        self.cybos.SetInputValue(0, stock_code)  # 종목코드
        self.cybos.SetInputValue(1, ord('2'))  # 기간으로 요청
        self.cybos.SetInputValue(2, datetime.now().strftime("%Y%m%d"))  # 요청일자 (오늘)
        self.cybos.SetInputValue(3, datetime.now().strftime("%Y%m%d"))  # 시작일자
        self.cybos.SetInputValue(4, 1000)  # 요청개수
        self.cybos.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8])  # 요청항목 (날짜, 시간, 시가, 고가, 저가, 종가, 거래량)
        self.cybos.SetInputValue(6, ord('m'))  # 분봉
        self.cybos.SetInputValue(9, ord('1'))  # 수정주가 사용

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
            
            data.append([date, time, open_price, high_price, low_price, close_price, volume])

        df = pd.DataFrame(data, columns=['날짜', '시간', '시가', '고가', '저가', '종가', '거래량'])
        df['날짜'] = pd.to_datetime(df['날짜'].astype(str), format='%Y%m%d')
        df['시간'] = pd.to_datetime(df['시간'].astype(str).str.zfill(6), format='%H%M%S').dt.time
        df['일자'] = df['날짜'].dt.strftime('%Y%m%d') + df['시간'].astype(str)
        df = df[['일자', '시가', '고가', '저가', '종가', '거래량']]

        df.to_csv(f"{stock_name}.csv", index=False)
        print(f"{stock_name}.csv 파일 생성 완료")

    def update_csv_files(self, stock_list):
        for stock_name in stock_list:
            self.create_stock_csv(stock_name)