# 32bit_kiwoom_server.py (32비트 환경에서 실행)
from flask import Flask, jsonify
from pykiwoom.kiwoom import *
import pythoncom

app = Flask(__name__)
kiwoom = Kiwoom()

@app.route('/login', methods=['POST'])
def login():
    pythoncom.CoInitialize()  # COM 객체 초기화
    
    # 로그인 창 표시
    kiwoom.CommConnect(block=True)
    
    if kiwoom.GetConnectState() == 1:
        login_info = {
            "status": "logged_in",
            "message": "Login successful"
        }
    else:
        login_info = {
            "status": "login_failed",
            "message": "Login failed"
        }
    
    pythoncom.CoUninitialize()  # COM 객체 해제
    
    return jsonify(login_info)

@app.route('/user_info')
def user_info():
    if kiwoom.GetConnectState() == 0:
        return jsonify({"error": "Not connected"}), 400
    
    account_num = kiwoom.GetLoginInfo("ACCOUNT_CNT")
    accounts = kiwoom.GetLoginInfo("ACCNO")
    user_id = kiwoom.GetLoginInfo("USER_ID")
    user_name = kiwoom.GetLoginInfo("USER_NAME")
    return jsonify({
        "account_num": account_num,
        "accounts": accounts,
        "user_id": user_id,
        "user_name": user_name
    })

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000)