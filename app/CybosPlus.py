import win32com.client
import pandas as pd
from datetime import datetime, timedelta
import os
import time
import threading
import boto3
import json
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

    def create_stock_data_json(self, stock_name):
    
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

            data.append({
                "Date": f"{date}{time:04d}",
                "Open": open_price,
                "High": high_price,
                "Low": low_price,
                "Close": close_price,
                "Volume": volume,
                "Amount": trading_value
            })
        data.sort(key=lambda x: x["Date"])
        # JSON 파일로 저장
        with open(f"{stock_name}_data.json", "w") as f:
            json.dump(data, f, indent=4)

        print(f"{stock_name}_data.json 파일 생성 완료")

        # 분마다 데이터 업데이트 시작
        self.start_minute_update(stock_code, stock_name)
    
    def get_stock_info(self, stock_name,id):
        stock_code = self.get_stock_code(stock_name)
        if stock_code is None:
            return None

        self.objStockMst.SetInputValue(0, stock_code)
        self.objStockMst.BlockRequest()
        _id = id
        return {
            "_id":_id,
            "code": stock_code,
            "name": self.objStockMst.GetHeaderValue(1),  # 종목명
            "price": self.objStockMst.GetHeaderValue(13),  # 시가
            "rate": self.objStockMst.GetHeaderValue(12),  # 등락률
            "status" : "random"
        }


    def start_minute_update(self, stock_code, stock_name):
        updater = MinuteDataUpdater(self, stock_code, stock_name)
        updater.start()

    def update_json_files(self, stock_list):
        for stock_name in stock_list:
            self.create_stock_data_json(stock_name)

    def get_current_data(self, stock_code):
        self.objStockCur.SetInputValue(0, stock_code)
        self.objStockCur.Subscribe()
        time.sleep(1)

        current_time = datetime.now().strftime("%Y%m%d%H:%M")
        current_price = self.objStockCur.GetHeaderValue(13)
        open_price = self.objStockCur.GetHeaderValue(4)
        high_price = self.objStockCur.GetHeaderValue(5)
        low_price = self.objStockCur.GetHeaderValue(6)
        volume = self.objStockCur.GetHeaderValue(9)
        trading_value = current_price * volume

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
            # 최신 30개의 데이터를 가져옵니다
            self.cybos_api.cybos.SetInputValue(0, self.stock_code)
            self.cybos_api.cybos.SetInputValue(1, ord('2'))
            self.cybos_api.cybos.SetInputValue(2, datetime.now().strftime("%Y%m%d"))
            self.cybos_api.cybos.SetInputValue(3, datetime.now().strftime("%Y%m%d"))
            self.cybos_api.cybos.SetInputValue(4, 30)
            self.cybos_api.cybos.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8])
            self.cybos_api.cybos.SetInputValue(6, ord('m'))
            self.cybos_api.cybos.SetInputValue(9, ord('1'))
            self.cybos_api.cybos.BlockRequest()

            num_data = self.cybos_api.cybos.GetHeaderValue(3)
            data = []

            for i in range(num_data):
                date = self.cybos_api.cybos.GetDataValue(0, i)
                time = self.cybos_api.cybos.GetDataValue(1, i)
                open_price = self.cybos_api.cybos.GetDataValue(2, i)
                high_price = self.cybos_api.cybos.GetDataValue(3, i)
                low_price = self.cybos_api.cybos.GetDataValue(4, i)
                close_price = self.cybos_api.cybos.GetDataValue(5, i)
                volume = self.cybos_api.cybos.GetDataValue(6, i)
                trading_value = close_price * volume

                data.append({
                    "Date": f"{date}{time:04d}",
                    "Open": open_price,
                    "High": high_price,
                    "Low": low_price,
                    "Close": close_price,
                    "Volume": volume,
                    "Amount": trading_value
                })
            data.sort(key=lambda x: x["Date"])
            # JSON 파일로 저장
            json_file_name = f"{self.stock_name}.json"
            with open(json_file_name, "w") as f:
                json.dump(data, f, indent=4)

            print(f"{self.stock_name} 데이터 업데이트 완료: {len(data)} 개의 데이터")

            # S3에 업로드
            bucket_name = 'dev-jeus-bucket'  # S3 버킷 이름
            s3_file_name = f"{self.stock_name}.json"  # S3에 저장될 파일 이름
            upload_to_s3(json_file_name, bucket_name, s3_file_name)

            # 다음 분의 00초까지 대기
            now = datetime.now()
            next_minute = (now + timedelta(minutes=1)).replace(second=0, microsecond=0)
            time_to_wait = (next_minute - now).total_seconds()
            time.sleep(time_to_wait)