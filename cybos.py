import win32com.client
import pandas as pd

# 대신증권 API 초기화
def init_cybos():
    cybos = win32com.client.Dispatch("CpUtil.CpCybos")
    if not cybos.IsConnect:
        print("대신증권에 연결되지 않았습니다.")
        return False
    print("대신증권에 정상적으로 연결되었습니다.")
    return True

# 과거 데이터 조회 함수
def get_minute_data(stock_code, start_date, end_date):
    chart = win32com.client.Dispatch("CpSysDib.StockChart")
    chart.SetInputValue(4, 1000)
    # 대신증권 API에서 필요한 만큼 충분히 데이터를 가져오기
    chart.SetInputValue(0, stock_code)  # 종목 코드
    chart.SetInputValue(1, ord('1'))    # 1: 기간으로 설정
    chart.SetInputValue(2, end_date)    # 요청 끝 날짜
    chart.SetInputValue(3, start_date)  # 요청 시작 날짜
    chart.SetInputValue(5, [0, 1, 2, 3, 4, 5])  # 0: 날짜, 1: 시간, 2: 시가, 3: 고가, 4: 저가, 5: 종가
    chart.SetInputValue(6, ord('m'))    # 'm': 분봉 데이터 요청
    chart.SetInputValue(9, ord('1'))    # 수정주가 사용 여부 (1: 사용)

    chart.BlockRequest()

    num_data = chart.GetHeaderValue(3)  # 데이터 개수 확인

    data = []
    for i in range(num_data):
        date = chart.GetDataValue(0, i)
        time = chart.GetDataValue(1, i)
        open_price = chart.GetDataValue(2, i)
        high_price = chart.GetDataValue(3, i)
        low_price = chart.GetDataValue(4, i)
        close_price = chart.GetDataValue(5, i)
        data.append([date, time, open_price, high_price, low_price, close_price])

    # Pandas DataFrame으로 변환
    columns = ['Date', 'Time', 'Open', 'High', 'Low', 'Close']
    df = pd.DataFrame(data, columns=columns)

    # 원하는 날짜 범위만 필터링
    df_filtered = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]

    return df_filtered

if __name__ == "__main__":
    # 대신증권 연결 초기화
    if not init_cybos():
        exit()

    # 종목 코드 및 날짜 설정
    stock_code = 'A005930'  # 삼성전자 예시
    start_date = 20231231  # 시작 날짜 (2024년 1월 1일)
    end_date = 20240102    # 끝 날짜 (2024년 1월 2일)

    # 과거 분봉 데이터 가져오기
    df = get_minute_data(stock_code, start_date, end_date)

    # 결과 출력
    print(df)
