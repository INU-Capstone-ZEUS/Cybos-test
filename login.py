from flask import Flask, jsonify
from pykiwoom.kiwoom import *

app = Flask(__name__)
kiwoom = Kiwoom()
kiwoom.CommConnect(block=True)

@app.route('/login')
def login():
    return jsonify({"status": "logged_in"})

@app.route('/user_info')
def user_info():
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