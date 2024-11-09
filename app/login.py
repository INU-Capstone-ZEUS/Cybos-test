import sys
from PyQt5.QtWidgets import QApplication
from hero4 import Kiwoom
from CybosPlus import CybosAPI

def main():
    app = QApplication(sys.argv)
    
    # Kiwoom API 초기화 및 로그인
    kiwoom = Kiwoom()
    kiwoom.show()
    kiwoom.comm_connect()

    # CybosPlus API 초기화
    cybos = CybosAPI()

    # 조건검색 결과 업데이트 시 CSV 파일 생성/업데이트
    def update_csv_files():
        stock_list = kiwoom.get_condition_search_results()
        cybos.update_csv_files(stock_list)

    # Kiwoom API의 조건검색 결과 업데이트 시그널에 CSV 업데이트 함수 연결
    kiwoom.kiwoom.OnReceiveTrCondition.connect(lambda *args: update_csv_files())
    kiwoom.kiwoom.OnReceiveRealCondition.connect(lambda *args: update_csv_files())

    sys.exit(app.exec_())

if __name__ == "__main__":
    main()