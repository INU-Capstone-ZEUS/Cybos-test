from flask import Flask, jsonify
from pykiwoom.kiwoom import *
import logging
import sys
from PyQt5.QtWidgets import QApplication

app = Flask(__name__)

logging.basicConfig(level=logging.DEBUG)

class KiwoomAPI:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.kiwoom = Kiwoom()

    def login(self):
        if not self.kiwoom.GetConnectState():
            self.kiwoom.CommConnect()
            self.app.exec_()
        
        login_status = self.kiwoom.GetConnectState()
        if login_status == 1:
            return {"status": "success", "message": "로그인 성공"}
        else:
            return {"status": "error", "message": "로그인 실패"}

kiwoom_api = KiwoomAPI()

@app.route('/login', methods=['GET', 'POST'])
def login():
    try:
        result = kiwoom_api.login()
        return jsonify(result)
    except Exception as e:
        logging.error(f"로그인 중 오류 발생: {str(e)}")
        return jsonify({"status": "error", "message": str(e)}), 500

# 주식 종목 데이터를 가져오는 엔드포인트 추가
@app.route('/stock_data/<code>', methods=['GET'])
def get_stock_data(code):
    try:
        kiwoom_api.kiwoom.SetInputValue("종목코드", code)
        kiwoom_api.kiwoom.CommRqData("주식기본정보", "opt10001", 0, "0101")
        
        name = kiwoom_api.kiwoom.GetCommData("주식기본정보", "종목명", 0).strip()
        price = kiwoom_api.kiwoom.GetCommData("주식기본정보", "현재가", 0).strip()
        
        return jsonify({
            "status": "success",
            "data": {
                "code": code,
                "name": name,
                "price": price
            }
        })
    except Exception as e:
        logging.error(f"주식 데이터 조회 중 오류 발생: {str(e)}")
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)