from pykiwoom.kiwoom import Kiwoom
import pandas as pd
import time

# 키움 API 인스턴스 생성 및 로그인
# kiwoom = Kiwoom()
# kiwoom.login()
# kiwoom.CommConnect(block=True)
#
# # 조건식 로드
# kiwoom.GetConditionLoad()
#
# # 조건식 목록 가져오기
# conditions = kiwoom.GetConditionNameList()
# print("조건식 목록:")
# for condition_id, condition_name in conditions:
#     print(f"ID: {condition_id}, 이름: {condition_name}")
#
# # 조건검색으로 종목 리스트 가져오기
# condition_name = "장밴"  # 사용자가 원하는 조건식 이름
# condition_index = 2  # 조건식 ID (예시)
#
# # 조건검색으로 종목 리스트 가져오기
# stock_list = kiwoom.SendCondition("0101", condition_name, condition_index, 0)
# print("조건검색 결과:", stock_list)
#
# # 주가 데이터를 저장할 DataFrame 생성
# price_data = pd.DataFrame()
#
# # 주가 데이터를 가져오기
# for code in stock_list:
#     name = kiwoom.GetMasterCodeName(code)
#     # 일봉 데이터 조회
#     df = kiwoom.block_request("opt10081",
#                               종목코드=code,
#                               기준일자="20231001",  # 예시로 2023년 10월 1일부터 조회
#                               수정주가구분=1,
#                               output="주가",
#                               next=0)
#
#     # 종목 이름을 DataFrame에 추가
#     df['종목명'] = name
#     price_data = price_data.append(df, ignore_index=True)
#
#     time.sleep(0.2)  # 요청 과부하 방지를 위해 약간의 딜레이 추가
#
# # 최종 주가 데이터 출력
# print(price_data)
def login():
    kiwoom = Kiwoom()
    try:
        kiwoom.CommConnect(block=True)
        print("succece!")
    except Exception as e:
        print("Fail",str(e))
        return None

    return kiwoom


def get_stock_code_by_name(kiwoom, stock_name):
    # 종목 코드와 종목 이름을 매핑할 딕셔너리
    stock_list = kiwoom.GetCodeListByMarket('0')  # 0은 코스피, 10은 코스닥
    stock_info = {}

    for code in stock_list:
        name = kiwoom.GetMasterCodeName(code)
        stock_info[name] = code  # 종목명과 종목코드를 매핑

    # 종목명으로 코드 검색
    return stock_info.get(stock_name)

def search_info(kiwoom, stock_code):
    # 종목 정보 조회
    try:
        # Opt10001을 사용하여 주식 정보를 요청
        stock_info = kiwoom.block_request("opt10001",
                                            종목코드=stock_code,
                                            output="주식정보",
                                            next=0)
        return stock_info  # 주식 정보 반환
    except Exception as e:
        print("종목 정보 조회 실패:", str(e))
        return None

if __name__ == "__main__":
    kiwoom_instance = login()
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
                # 특정 종목 검색
                stock_info = search_info(kiwoom_instance, stock_code)

                if stock_info is not None:
                    print("종목 정보:", stock_info)
                else:
                    print("종목 정보를 가져올 수 없습니다.")
            else:
                print("해당 종목명이 존재하지 않습니다.")
