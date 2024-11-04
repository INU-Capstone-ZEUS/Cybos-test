from flask import Flask, jsonify, request
from pykiwoom.kiwoom import *
import pythoncom
import os

app = Flask(__name__)
kiwoom = Kiwoom()

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'GET':
        return jsonify({"message": "Please use POST method for login"}), 200

    pythoncom.CoInitialize()
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
    
    pythoncom.CoUninitialize()
    return jsonify(login_info)

if __name__ == '__main__':
    host = os.environ.get('FLASK_HOST', '0.0.0.0')
    port = int(os.environ.get('FLASK_PORT', 5000))
    app.run(host=host, port=port)