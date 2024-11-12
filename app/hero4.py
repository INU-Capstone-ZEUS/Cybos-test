import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *

class Kiwoom(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("실시간 조건검색 결과")
        self.setGeometry(300, 300, 500, 600)

        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.kiwoom.OnEventConnect.connect(self._event_connect)
        self.kiwoom.OnReceiveConditionVer.connect(self._receive_condition_ver)
        self.kiwoom.OnReceiveTrCondition.connect(self._receive_tr_condition)
        self.kiwoom.OnReceiveRealCondition.connect(self._receive_real_condition)

        self.login_event_loop = None
        self.condition_event_loop = None
        self.screen_no = "0101"  # 실시간 조건검색 화면 번호

        self.init_ui()

        QTimer.singleShot(1000, self.auto_start_condition_search)
    def init_ui(self):
        self.condition_combo = QComboBox(self)
        self.condition_combo.move(20, 20)
        self.condition_combo.resize(200, 30)

        self.condition_button = QPushButton("조건검색 실행", self)
        self.condition_button.move(240, 20)
        self.condition_button.clicked.connect(self.start_condition_search)

        self.stock_table = QTableWidget(self)
        self.stock_table.setColumnCount(3)
        self.stock_table.setHorizontalHeaderLabels(["종목코드", "종목명", "상태"])
        self.stock_table.setColumnWidth(0, 80)
        self.stock_table.setColumnWidth(1, 160)
        self.stock_table.setColumnWidth(2, 80)
        self.stock_table.move(20, 70)
        self.stock_table.resize(460, 500)

    def auto_start_condition_search(self):
        if self.condition_combo.currentIndex() >= 0:
            self.start_condition_search()

    def start_condition_search(self):
        if self.condition_combo.currentIndex() >= 0:
            condition_index = self.condition_combo.currentData()
            condition_name = self.condition_combo.currentText()
            ret = self.kiwoom.dynamicCall("SendCondition(QString, QString, int, int)",
                                          self.screen_no, condition_name, condition_index, 0)
            if ret == 1:
                print(f"조건검색 '{condition_name}' 요청 성공")
            else:
                print(f"조건검색 '{condition_name}' 요청 실패")

    def comm_connect(self):
        self.kiwoom.dynamicCall("CommConnect()")
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def _event_connect(self, err_code):
        if err_code == 0:
            print("로그인 성공")
            self.get_condition_load()
        else:
            print("로그인 실패")
        self.login_event_loop.exit()

    def get_condition_load(self):
        ret = self.kiwoom.dynamicCall("GetConditionLoad()")
        if ret == 1:
            print("조건식 불러오기 요청 성공")
            self.condition_event_loop = QEventLoop()
            self.condition_event_loop.exec_()
        else:
            print("조건식 불러오기 요청 실패")

    def _receive_condition_ver(self, ret, msg):
        if ret == 1:
            print("조건식 불러오기 성공")
            condition_list = self.kiwoom.dynamicCall("GetConditionNameList()")
            conditions = condition_list.split(';')[:-1]
            for condition in conditions:
                index, name = condition.split('^')
                self.condition_combo.addItem(name, int(index))
            
            # 조건식 로드 완료 후 자동으로 첫 번째 조건 선택 및 검색 시작
            if self.condition_combo.count() > 0:
                self.condition_combo.setCurrentIndex(0)
                self.auto_start_condition_search()
        else:
            print(f"조건식 불러오기 실패: {msg}")
        self.condition_event_loop.exit()

    def _receive_tr_condition(self, screen_no, codes, condition_name, condition_index, next):
        print(f"조건검색 '{condition_name}' 결과:")
        self.stock_table.setRowCount(0)
        if codes:
            code_list = codes.split(';')
            for code in code_list:
                name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", code)
                row = self.stock_table.rowCount()
                self.stock_table.insertRow(row)
                self.stock_table.setItem(row, 0, QTableWidgetItem(code))
                self.stock_table.setItem(row, 1, QTableWidgetItem(name))
                self.stock_table.setItem(row, 2, QTableWidgetItem("편입"))
                print(f"종목코드: {code}, 종목명: {name}")
        else:
            print("해당 조건을 만족하는 종목이 없습니다.")
        self.update_txt_file()

    def _receive_real_condition(self, code, event_type, condition_name, condition_index):
        name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", code)
        if event_type == "I":
            print(f"편입: 종목코드: {code}, 종목명: {name}")
            row = self.stock_table.rowCount()
            self.stock_table.insertRow(row)
            self.stock_table.setItem(row, 0, QTableWidgetItem(code))
            self.stock_table.setItem(row, 1, QTableWidgetItem(name))
            self.stock_table.setItem(row, 2, QTableWidgetItem("편입"))
        elif event_type == "D":
            print(f"이탈: 종목코드: {code}, 종목명: {name}")
            for row in range(self.stock_table.rowCount()):
                if self.stock_table.item(row, 0).text() == code:
                    self.stock_table.setItem(row, 2, QTableWidgetItem("이탈"))
                    break
        self.update_txt_file()

    def update_txt_file(self):
        try:
            with open("condition_search_results.txt", 'w', encoding='utf-8') as file:
                for row in range(self.stock_table.rowCount()):
                    name = self.stock_table.item(row, 1).text()
                    status = self.stock_table.item(row, 2).text()
                    if status == "편입":
                        file.write(f"{name}\n")
            print("텍스트 파일 업데이트 완료")
        except Exception as e:
            print(f"텍스트 파일 업데이트 중 오류 발생: {str(e)}")

    def get_condition_search_results(self):
        results = []
        for row in range(self.stock_table.rowCount()):
            name = self.stock_table.item(row, 1).text()  # 종목명은 두 번째 열(인덱스 1)에 있습니다.
            status = self.stock_table.item(row, 2).text()
            if status == "편입" and name.strip():  # 종목명이 비어있지 않은 경우만 추가
                results.append(name)
                print(f"조건검색 결과: 종목명 {name}")  # 디버깅용 출력
        return results