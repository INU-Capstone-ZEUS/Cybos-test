import win32com.client
import pandas as pd
from datetime import datetime
import time
import threading
import boto3
from botocore.exceptions import NoCredentialsError


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
        self.cybos.SetInputValue(4, 30)
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

        df = pd.DataFrame(data, columns=['날짜', '시간', 'Open', 'High', 'Low', 'End', 'Amount'])
        df['날짜'] = pd.to_datetime(df['날짜'].astype(str), format='%Y%m%d')
        df['시간'] = pd.to_datetime(df['시간'].astype(str).str.zfill(6), format='%H%M%S').dt.time
        df['Date'] = df['날짜'].dt.strftime('%Y%m%d') + df['시간'].astype(str)
        df = df[['Date', 'Open', 'High', 'Low', 'End', 'Amount']]
        #날짜 정렬
        df = df.sort_values('Date', ascending=True)

        df.to_csv(f"{stock_name}.csv", index=False)
        print(f"{stock_name}.csv 파일 생성 완료")

        # 분마다 데이터 업데이트 시작
        self.start_minute_update(stock_code, stock_name)

    def start_minute_update(self, stock_code, stock_name):
        updater = MinuteDataUpdater(self, stock_code, stock_name)
        updater.start()

    def update_csv_files(self, stock_list):
        for stock_name in stock_list:
            self.create_stock_csv(stock_name)

    def get_current_data(self, stock_code):
        self.objStockCur.SetInputValue(0, stock_code)
    
        # Subscribe 메서드를 사용하여 실시간 데이터 구독
        self.objStockCur.Subscribe()

        # 잠시 대기하여 데이터를 받을 시간을 줍니다
        time.sleep(1)

        current_time = datetime.now().strftime("%Y%m%d%H%M%S")
        current_price = self.objStockCur.GetHeaderValue(13)  # 현재가
        open_price = self.objStockCur.GetHeaderValue(4)  # 시가
        high_price = self.objStockCur.GetHeaderValue(5)  # 고가
        low_price = self.objStockCur.GetHeaderValue(6)  # 저가
        volume = self.objStockCur.GetHeaderValue(9)  # 거래량
        trading_value = current_price * volume  # 거래대금 계산
        # 구독 해제
        self.objStockCur.Unsubscribe()

        return current_time, open_price, high_price, low_price, current_price, trading_value

def upload_to_s3(local_file, bucket, s3_file):
    s3 = boto3.client('s3',
        aws_access_key_id=ACCESS_KEY,
        aws_secret_access_key=SECRET_KEY)
    try:
        s3.upload_file(local_file, bucket, s3_file)
        print(f"Upload Successful: {local_file} to {s3_file}")
        return True
    except FileNotFoundError:
        print(f"The file {local_file} was not found")
        return False
    except NoCredentialsError:
        print("Credentials not available")
        return False
    
class MinuteDataUpdater:
    def __init__(self, cybos_api, stock_code, stock_name):
        self.cybos_api = cybos_api
        self.stock_code = stock_code
        self.stock_name = stock_name
        self.running = False

    def start(self):
        self.running = True
        thread = threading.Thread(target=self.update_loop)
        thread.start()

    def stop(self):
        self.running = False

    def update_loop(self):
        while self.running:
            current_time, open_price, high_price, low_price, current_price, trading_value = self.cybos_api.get_current_data(self.stock_code)

            # CSV 파일에 데이터 추가
            csv_file_name = f"{self.stock_name}.csv"
            with open(csv_file_name, "a") as f:
                f.write(f"{current_time},{open_price},{high_price},{low_price},{current_price},{trading_value}\n")

            print(f"{self.stock_name} 데이터 추가: {current_time}, 시가: {open_price}, 고가: {high_price}, 저가: {low_price}, 현재가: {current_price}, 거래대금: {trading_value}")

            # S3에 업로드
            bucket_name = 'dev-jeus-bucket'  # S3 버킷 이름
            s3_file_name = f"{self.stock_name}.csv"  # S3에 저장될 파일 이름
            upload_to_s3(csv_file_name, bucket_name, s3_file_name)

            # 1분 대기
            time.sleep(60)