from flask import Flask, jsonify, request
from DesinAPI import DesinAPI
import pandas as pd
import pythoncom
from flask_cors import CORS


app = Flask(__name__)
CORS(app)
desin_api = DesinAPI()
@app.route('/get_stock_code/<name>', methods=['GET'])
def get_stock_code(name):
    pythoncom.CoInitialize()
    try:
        code = desin_api.GetStockCode(name)
        if code:
            return jsonify({"name": name, "code": code})
        else:
            return jsonify({"error": f"Stock code not found for {name}"}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally :
        pythoncom.CoUninitialize()


@app.route('/search_name_list_by_name', methods=['POST'])
def search_name_list_by_name():
    pythoncom.CoInitialize()
    name = request.json['name']
    result = desin_api.SearchNameListByName(name)
    pythoncom.CoUninitialize()
    return jsonify(result)

@app.route('/search_name_list_by_code', methods=['POST'])
def search_name_list_by_code():
    pythoncom.CoInitialize()
    code = request.json['code']
    result = desin_api.SearchNameListByCode(code)
    pythoncom.CoUninitialize()
    return jsonify(result)

@app.route('/get_recent_data', methods=['POST'])
def get_recent_data():
    pythoncom.CoInitialize()
    data = request.json
    result = desin_api.GetRecentData(data['codeName'], data['count'], data['tick_range'], None, data['mT'])
    pythoncom.CoInitialize()
    return jsonify(result.to_dict(orient='records'))

@app.route('/get_day_data', methods=['POST'])
def get_day_data():
    pythoncom.CoInitialize()
    data = request.json
    result = desin_api.GetDayData(data['codeName'], data['count'], data['tick_range'], None)
    pythoncom.CoInitialize()
    return jsonify(result.to_dict(orient='records'))

@app.route('/get_minute_or_tick_data', methods=['POST'])
def get_minute_or_tick_data():
    pythoncom.CoInitialize()
    data = request.json
    result = desin_api.GetMinuteOrTickData(data['codeName'], data['count'], data['tick_range'], None, data['mT'])
    pythoncom.CoInitialize()
    return jsonify(result.to_dict(orient='records'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
    