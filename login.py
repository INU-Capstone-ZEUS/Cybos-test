from flask import Flask, jsonify
from pykiwoom.kiwoom import *
import time

app = Flask(__name__)
kiwoom = Kiwoom()

def maintain_connection():
    while True:
        if kiwoom.GetConnectState() == 0:
            kiwoom.CommConnect(block=True)
        time.sleep(60)  # 1분마다 연결 상태 확인

import threading
threading.Thread(target=maintain_connection, daemon=True).start()

@app.route('/login')
def login():
    if kiwoom.GetConnectState() == 0:
        kiwoom.CommConnect(block=True)
    return jsonify({"status": "logged_in"})

@app.route('/user_info')
def user_info():
    if kiwoom.GetConnectState() == 0:
        return jsonify({"error": "Not connected"}), 400
    
    account_num = kiwoom.GetLoginInfo('ACCOUNT_CNT')
    accounts = kiwoom.GetLoginInfo('ACCNO')
    user_id = kiwoom.GetLoginInfo('USER_ID')
    user_name = kiwoom.GetLoginInfo('USER_NAME')
    return jsonify({
        "account_num": account_num,
        "accounts": accounts,
        "user_id": user_id,
        "user_name": user_name
    })

if __name__ == '__main__':
    app.run(port=5000)