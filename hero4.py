from pykiwoom.kiwoom import Kiwoom
import pandas as pd
import time

# 키움 API 인스턴스 생성 및 로그인
kiwoom = Kiwoom()
kiwoom.CommConnect(block=True)

# 조건식 로드
kiwoom.GetConditionLoad()

# 조건식 목록 가져오기
conditions = kiwoom.GetConditionNameList()
print("조건식 목록:")
for condition_id, condition_name in conditions:
    print(f"ID: {condition_id}, 이름: {condition_name}")

# 조건검색으로 종목 리스트 가져오기
condition_name = "장밴"  # 사용자가 원하는 조건식 이름
condition_index = 2  # 조건식 ID (예시)

# 조건검색으로 종목 리스트 가져오기
stock_list = kiwoom.SendCondition("0101", condition_name, condition_index, 0)
print("조건검색 결과:", stock_list)

# 주가 데이터를 저장할 DataFrame 생성
price_data = pd.DataFrame()

# 주가 데이터를 가져오기
for code in stock_list:
    name = kiwoom.GetMasterCodeName(code)
    # 일봉 데이터 조회
    df = kiwoom.block_request("opt10081",
                              종목코드=code,
                              기준일자="20231001",  # 예시로 2023년 10월 1일부터 조회
                              수정주가구분=1,
                              output="주가",
                              next=0)

    # 종목 이름을 DataFrame에 추가
    df['종목명'] = name
    price_data = price_data.append(df, ignore_index=True)

    time.sleep(0.2)  # 요청 과부하 방지를 위해 약간의 딜레이 추가

# 최종 주가 데이터 출력
print(price_data)
