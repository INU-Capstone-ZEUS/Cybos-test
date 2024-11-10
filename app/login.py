import sys
from PyQt5.QtWidgets import QApplication
from hero4 import Kiwoom
from CybosPlus import CybosAPI
import boto3
from botocore.exceptions import NoCredentialsError

def upload_to_s3(local_file, bucket, s3_file):
    s3 = boto3.client('s3')

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

def update_csv_files(kiwoom, cybos):
    stock_list = kiwoom.get_condition_search_results()
    cybos.update_csv_files(stock_list)
    
    s3_bucket_name = "dev-jeus-bucket"
    stock_list_file = "condition_search_results.txt"
    # S3에 업로드
    for stock_name in stock_list:
        local_file = f"{stock_name}.csv"
        s3_file = f"stock_data/{stock_name}.csv"
        upload_to_s3(local_file, 'YOUR_BUCKET_NAME', s3_file)

def main():
    app = QApplication(sys.argv)
    
    # Kiwoom API 초기화 및 로그인
    kiwoom = Kiwoom()
    kiwoom.show()
    kiwoom.comm_connect()

    # CybosPlus API 초기화
    cybos = CybosAPI()

    # 조건검색 결과 업데이트 시 CSV 파일 생성/업데이트 및 S3 업로드
    kiwoom.kiwoom.OnReceiveTrCondition.connect(lambda *args: update_csv_files(kiwoom, cybos))
    kiwoom.kiwoom.OnReceiveRealCondition.connect(lambda *args: update_csv_files(kiwoom, cybos))

    sys.exit(app.exec_())

if __name__ == "__main__":
    main()