import json
import sys
from PyQt5.QtWidgets import QApplication
import requests
from hero4 import Kiwoom
from CybosPlus import CybosAPI
import boto3
from botocore.exceptions import NoCredentialsError
import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from dotenv import load_dotenv


base_url = "https://jeus.site:8080"

def alert_list():
    url = f"{base_url}/alert_list"
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()
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

def update_json_files(kiwoom, cybos):
    
    try:
        with open("condition_search_results.txt", 'r', encoding='utf-8') as file:
            stock_list = [line.strip() for line in file if line.strip()]
    except FileNotFoundError:
        print("condition_search_results.txt 파일을 찾을 수 없습니다.")
        stock_list = []
    except Exception as e:
        print(f"파일 읽기 중 오류 발생: {str(e)}")
        stock_list = []
    
    stock_list_data = cybos.get_stock_info(stock_list)
    

    with open("주도주리스트.json", "w", encoding='utf-8') as f:
        json.dump(stock_list_data, f, ensure_ascii=False, indent=4)
    print("주도주리스트.json 파일 생성 완료")
    
    
    # S3에 업로드
    upload_to_s3("주도주리스트.json", 'dev-jeus-bucket', "주도주리스트.json")

    cybos.update_json_files(stock_list)


class FileHandler(FileSystemEventHandler):
    def __init__(self, kiwoom, cybos):
        self.kiwoom = kiwoom
        self.cybos = cybos

    def on_modified(self, event):
        if not event.is_directory and event.src_path.endswith("condition_search_results.txt"):
            print(f"File {event.src_path} has been modified")
            update_json_files(self.kiwoom, self.cybos)

def initialize_files():
    # TXT 파일 초기화
    with open("condition_search_results.txt", "w") as f:
        f.write("")
    print("condition_search_result.txt 파일이 초기화되었습니다.")

    # json 파일들 삭제
    for file in os.listdir():
        if file.endswith(".json"):
            os.remove(file)
            print(f"{file} 파일이 삭제되었습니다.")

def main():
    app = QApplication(sys.argv)
    initialize_files()

    # Kiwoom API 초기화 및 로그인
    kiwoom = Kiwoom()
    kiwoom.show()
    kiwoom.comm_connect()

    # CybosPlus API 초기화
    cybos = CybosAPI()

    # 조건검색 결과 업데이트 시 json 파일 생성/업데이트 및 S3 업로드
    # kiwoom.kiwoom.OnReceiveTrCondition.connect(lambda *args: update_json_files(kiwoom, cybos))
    # kiwoom.kiwoom.OnReceiveRealCondition.connect(lambda *args: update_json_files(kiwoom, cybos))
    # 파일 변경 감지 설정
    event_handler = FileHandler(kiwoom,cybos)
    observer = Observer()
    observer.schedule(event_handler, path='.', recursive=False)
    observer.start()

    try:
        app.exec_()
    except KeyboardInterrupt:
        print("프로그램을 종료합니다.")
    finally:
        cybos.stop_all_updaters()
        observer.stop()
        observer.join()

if __name__ == "__main__":
    main()