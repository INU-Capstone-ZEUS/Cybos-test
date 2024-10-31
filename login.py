from pykiwoom.kiwoom import Kiwoom
import time

def login_and_keep_alive():
    kiwoom = Kiwoom()
    kiwoom.CommConnect(block=True)
    print("로그인 완료")

    # 주기적으로 서버와 통신하여 세션 유지
    while True:
        state = kiwoom.GetConnectState()
        if state == 0:
            kiwoom.CommConnect(block=True)
        
        # 예시로 현재 시간을 요청하여 세션 유지
        server_time = kiwoom.GetServerGubun()
        
        #print(f"서버 시간: {server_time}")
        
        # 5분마다 서버와 통신
        time.sleep(300)

        return kiwoom

if __name__ == "__main__":
    login_and_keep_alive()
