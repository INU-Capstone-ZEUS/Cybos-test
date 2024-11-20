import win32com.client
import pandas as pd
import threading
from datetime import datetime, timedelta
import threading
import boto3
import json
from botocore.exceptions import NoCredentialsError
from dotenv import load_dotenv
import requests
import pythoncom


class CybosAPI:
    def __init__(self):
        self.objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        if self.objCpCybos.IsConnect == 0:
            print("CYBOS Plus가 연결되어 있지 않습니다.")
            exit()
        self.updaters = {}

    def get_stock_code(self, stock_name):
        instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
        code = instCpStockCode.NameToCode(stock_name)
        if code == "":
            print(f"종목명 '{stock_name}'에 해당하는 종목코드를 찾을 수 없습니다.")
            return None
        return code

    def remove_A(self, stock_code):
        return stock_code[1:]

    
    
    def get_stock_info(self, stock_name, id):
        pythoncom.CoInitialize()
        try:
            cybos = win32com.client.Dispatch("CpSysDib.StockChart")
            objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
            stock_code = self.get_stock_code(stock_name)
            if stock_code is None:
                return None

            # 현재가 정보 요청
            print(stock_code)
            objStockMst.SetInputValue(0, stock_code)
            print("종목정보 검색")
            objStockMst.BlockRequest()
            print("완료")

            # 현재가
            current_price = objStockMst.GetHeaderValue(13)
            print(current_price)

            # 1분 전 종가 구하기
            cybos.SetInputValue(0, stock_code)
            cybos.SetInputValue(1, ord('2'))
            cybos.SetInputValue(2, datetime.now().strftime("%Y%m%d"))
            cybos.SetInputValue(3, datetime.now().strftime("%Y%m%d"))
            cybos.SetInputValue(4, 2)
            cybos.SetInputValue(5, [5])
            cybos.SetInputValue(6, ord('m'))
            cybos.SetInputValue(9, ord('1'))

            print("check1")
            cybos.BlockRequest()

            if cybos.GetHeaderValue(3) >= 2:
                prev_price = cybos.GetDataValue(0, 1)
                rate = ((current_price - prev_price) / prev_price * 100)
                rate = round(rate, 2)
            else:
                rate = 0.0

            _id = id
            stock_code = self.remove_A(stock_code)

            return {
                "_id": _id,
                "code": stock_code,
                "name": objStockMst.GetHeaderValue(1),
                "price": current_price,
                "rate": rate,
                "status": "random"
            }
        except pythoncom.com_error as e:
            print(f"COM 오류 발생: {e}")
            return None
        except Exception as e:
            print(f"Error in get_stock_info: {e}")
            return None
        finally:
            pythoncom.CoUninitialize()

    def update_json_files(self, stock_list):
        for stock_name in stock_list:
            self.create_stock_data_json(stock_name)

    def create_stock_data_json(self, stock_name):
        pythoncom.CoInitialize()
        try:
            cybos = win32com.client.Dispatch("CpSysDib.StockChart")
            stock_code = self.get_stock_code(stock_name)
            if stock_code is None:
                return

            print(f"종목명: {stock_name}, 종목코드: {stock_code}")

            cybos.SetInputValue(0, stock_code)
            cybos.SetInputValue(1, ord('2'))
            cybos.SetInputValue(2, datetime.now().strftime("%Y%m%d"))
            cybos.SetInputValue(3, datetime.now().strftime("%Y%m%d"))
            cybos.SetInputValue(4, 30)
            cybos.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8])
            cybos.SetInputValue(6, ord('m'))
            cybos.SetInputValue(9, ord('1'))
            cybos.BlockRequest()

            num_data = cybos.GetHeaderValue(3)
            data = []

            for i in range(num_data):
                date = cybos.GetDataValue(0, i)
                time = cybos.GetDataValue(1, i)
                open_price = str(cybos.GetDataValue(2, i))
                high_price = str(cybos.GetDataValue(3, i))
                low_price = str(cybos.GetDataValue(4, i))
                end_price = str(cybos.GetDataValue(5, i))
                volume = cybos.GetDataValue(6, i)
                amount = str(volume * int(end_price))
                hour = str(time // 100).zfill(2)
                minute = str(time % 100).zfill(2)
                formatted_date = f"{date}{hour}:{minute}"

                data.append({
                    "Date": formatted_date,
                    "Open": open_price,
                    "High": high_price,
                    "Low": low_price,
                    "End": end_price,
                    "Amount": amount
                })
            
            data.sort(key=lambda x: x["Date"])

            with open(f"{stock_name}데이타.json", "w") as f:
                json.dump(data, f, indent=4)
            print(f"{stock_name}데이타.json 파일 생성 완료")

            if stock_name not in self.updaters:
                updater = MinuteDataUpdater(stock_code, stock_name)
                updater.run()
                self.updaters[stock_name] = updater

        except Exception as e:
            print(f"오류 발생: {e}")
        finally:
            pythoncom.CoUninitialize()

    def stop_all_updaters(self):
        for updater in self.updaters.values():
            updater.stop()
        for updater in self.updaters.values():
            updater.join()
        self.updaters.clear()

class MinuteDataUpdater(threading.Thread):
    def __init__(self, stock_code, stock_name):
        threading.Thread.__init__(self)
        self.stock_code = stock_code
        self.stock_name = stock_name
        self.running = False
        self.timer = None

    def run(self):
        self.running = True
        self.schedule_next_update()
    
    def schedule_next_update(self):
        print("업데이트 시작")
        if not self.running:
            return

        now = datetime.now()
        next_minute = (now + timedelta(minutes=1)).replace(second=0, microsecond=0)
        delay = (next_minute - now).total_seconds()

        self.timer = threading.Timer(delay, self.perform_update)
        self.timer.start()

    def perform_update(self):
        print("update start")
        self.update_data()
        self.schedule_next_update()

    def stop(self):
        self.running = False
        if self.timer:
            self.timer.cancel()

    def update_data(self):
        pythoncom.CoInitialize()
        try:
            cybos = win32com.client.Dispatch("CpSysDib.StockChart")
            cybos.SetInputValue(0, self.stock_code)
            cybos.SetInputValue(1, ord('2'))
            cybos.SetInputValue(2, datetime.now().strftime("%Y%m%d"))
            cybos.SetInputValue(3, datetime.now().strftime("%Y%m%d"))
            cybos.SetInputValue(4, 30)
            cybos.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8])
            cybos.SetInputValue(6, ord('m'))
            cybos.SetInputValue(9, ord('1'))
            print("정보 요청 시작")
            cybos.BlockRequest()
            num_data = cybos.GetHeaderValue(3)
            data = []
            for i in range(num_data):
                date = cybos.GetDataValue(0, i)
                time = cybos.GetDataValue(1, i)
                open_price = str(cybos.GetDataValue(2, i))
                high_price = str(cybos.GetDataValue(3, i))
                low_price = str(cybos.GetDataValue(4, i))
                end_price = str(cybos.GetDataValue(5, i))
                volume = cybos.GetDataValue(6, i)
                amount = str(volume * int(end_price))
                hour = str(time // 100).zfill(2)
                minute = str(time % 100).zfill(2)
                formatted_date = f"{date}{hour}:{minute}"

                data.append({
                    "Date": formatted_date,
                    "Open": open_price,
                    "High": high_price,
                    "Low": low_price,
                    "End": end_price,
                    "Amount": amount
                })
        
            data.sort(key=lambda x: x["Date"])
            json_file_name = f"{self.stock_name}데이타.json"
            with open(json_file_name, "w", encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            print(f"{self.stock_name} 데이터 업데이트 완료: {len(data)} 개의 데이터")

            bucket_name = 'dev-jeus-bucket'
            s3_file_name = f"{self.stock_name}데이타.json"
            upload_to_s3(json_file_name, bucket_name, s3_file_name)

        except Exception as e:
            print(f"Error updating data for {self.stock_name}: {e}")
        finally:
            pythoncom.CoUninitialize()

def upload_to_s3(local_file, bucket, s3_file):
    s3 = boto3.client('s3', aws_access_key_id=ACCESS_KEY, aws_secret_access_key=SECRET_KEY)
    try:
        s3.upload_file(local_file, bucket, s3_file)
        print(f"Upload Successful: {local_file} to {s3_file}")
        data_to_fastapi(local_file)
        alert_chart()
        alert_list()
        return True
    except FileNotFoundError:
        print(f"The file {local_file} was not found")
        return False
    except NoCredentialsError:
        print("Credentials not available")
        return False

base_url = "https://jeus.site:8080"

def data_to_fastapi(json_file_name):
    url = f"{base_url}/predict"
    headers = {"accept": "application/json"}
    with open(json_file_name,'rb') as file:
        file = file.read()
    files = {
        'file': (json_file_name,file,'application/json')
    }
    response = requests.post(url, headers=headers,files=files)
    if response.status_code == 200:
        return response.json()
    print(json_file_name, response.json())

def alert_list():
    url = f"{base_url}/aler_list"
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()

def alert_chart():
    url = f"{base_url}/alert_chart"
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()