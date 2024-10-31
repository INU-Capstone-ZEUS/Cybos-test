import pandas as pd
from pykiwoom.kiwoom import Kiwoom
from datetime import datetime, timedelta

def login_to_kiwoom():
    # 키움 API 인스턴스 생성
    kiwoom = Kiwoom()
    
    # 로그인 시도
    try:
        kiwoom.CommConnect(block=True)  # 블록 방식으로 로그인 대기
        print("로그인 성공!")
    except Exception as e:
        print("로그인 실패:", str(e))
        return None
    
    return kiwoom  # 로그인 후 Kiwoom 인스턴스 반환

def get_stock_code_by_name(kiwoom, stock_name):
    # 종목 코드와 종목 이름을 매핑할 딕셔너리
    stock_list = kiwoom.GetCodeListByMarket('0')  # 0은 코스피, 10은 코스닥
    stock_info = {}
    
    for code in stock_list:
        name = kiwoom.GetMasterCodeName(code)
        stock_info[name] = code  # 종목명과 종목코드를 매핑

    # 종목명으로 코드 검색
    return stock_info.get(stock_name)

def fetch_minute_data(kiwoom, stock_code, num_data=1000):
    # 오늘부터 3일 전까지의 분봉 데이터 조회
    today = datetime.now().strftime('%Y%m%d')  # 오늘 날짜
    start_date = (datetime.now() - timedelta(days=3)).strftime('%Y%m%d')  # 3일 전 날짜

    df_all = pd.DataFrame()
    current_page = 0

    while True:
        # 주가 데이터를 요청
        data = kiwoom.block_request("opt10080", 
                                     종목코드=stock_code, 
                                     틱범위=1,  # 1분봉
                                     시작일자=start_date,
                                     종료일자=today,  # 종료일은 오늘
                                     페이지=current_page,
                                     output="분봉정보",
                                     next=0)
        
        df = pd.DataFrame(data)
        
        # 열 이름 출력
        print("현재 데이터의 열 이름:", df.columns.tolist())
        
        if df.empty:
            break
        
        df_all = pd.concat([df_all, df], ignore_index=True)

        # 1000개 데이터 가져오기
        if len(df_all) >= num_data:
            break
        
        current_page += 1  # 다음 페이지로 이동

    return df_all.head(num_data)  # 최대 num_data개만 반환

def save_minute_data_to_csv(df, stock_name):
    # 필요한 열만 선택
    try:
        df = df[["일자", "시간", "시가", "고가", "저가", "현재가", "거래량", "거래대금", 
                  "누적체결매도수량", "누적체결매수수량"]]
    except KeyError as e:
        print(f"오류: 선택한 열이 존재하지 않습니다. {e}")
        return
    
    # CSV 파일로 저장
    file_name = f"{stock_name}_minute_data.csv"
    df.to_csv(file_name, index=False, encoding='utf-8-sig')  # UTF-8 인코딩으로 저장
    print(f"{file_name} 파일로 저장되었습니다.")


if __name__ == "__main__":
    # 키움증권 로그인
    kiwoom_instance = login_to_kiwoom()
    if kiwoom_instance is not None:
        while True:
            # 사용자로부터 종목명 입력 받기
            stock_name = input("조회할 종목명을 입력하세요 (예: 삼성전자) 또는 '종료' 입력: ")
            
            if stock_name.lower() == '종료':  # '종료'를 입력하면 루프 종료
                print("종료합니다.")
                break
            
            # 종목명으로 종목 코드 검색
            stock_code = get_stock_code_by_name(kiwoom_instance, stock_name)
            
            if stock_code:
                # 오늘부터 3일 전까지의 분봉 데이터 조회
                minute_data = fetch_minute_data(kiwoom_instance, stock_code, num_data=1000)
                
                if not minute_data.empty:
                    save_minute_data_to_csv(minute_data, stock_name)  # CSV 파일로 저장
                else:
                    print("해당 종목의 분봉 데이터가 없습니다.")
            else:
                print("해당 종목명이 존재하지 않습니다.")
