import pywinauto
import time
from win32api import GetSystemMetrics

def logIn():
    try:
        app = pywinauto.application.Application()
        # 실행
        app.start(r"C:\DAISHIN\STARTER\ncStarter.exe /prj:cp")
        # 이미 연결 되어있을 수 있다. 예외 처리해야함 ;

        # (1) 키보드 보안경고창 예 버튼 클릭
        flag = 0
        while flag == 0:
            try:
                #title="대신증권 CYBOS FAMILY"
                #dlg = app.connect(title=title).Dialog
                #dlg['예(&Y)Button'].click()
                app.window(title_re='대신증권_CYBOS_FAMILY').window(title='예(Y)Button').click()

                print('대신증권 CYBOS FAMILY 키보드보안')

            except Exception as e:
                 print('Cybos Family : ', e)
                 break
        # (2) 비밀번호, 인증번호, 접속
        title = 'CYBOS Starter'
        dlg = pywinauto.timings.wait_until_passes(120, 30, lambda: app.window(title=title))

        # 통신암호
        flag = 0
        while flag == 0:
            try:
                dlg.Edit2.type_keys('s6m1')
                dlg.모의투자Button.click()  # 모의투자
                dlg.Button.click()

            except Exception as e:
                print('통신암호에러', e)
                break
            else:
                flag = 1
        # (3) 모의 투자 팝업
        time.sleep(10)  # 시스템 사양에 따라 달라짐
        # app.dlg.print_control_identifiers()
        title = "CYBOS Starter"
        new_dlg = app.connect(title=title).Dialog
        new_dlg.Button.click()
        time.sleep(10)  # 시스템 사양에 따라 달라짐
        # (4) 공지사항 팝업
        monitor_x = GetSystemMetrics(0)
        monitor_y = GetSystemMetrics(1)
        print("Width =", monitor_x, "Height =", monitor_y)
        if (monitor_x == 1920) and (monitor_y == 1080):
            pywinauto.mouse.click(button='left', coords=(1450, 160))  # 1920 x 1080
        elif (monitor_x == 1920) and (monitor_y == 1200):
            pywinauto.mouse.click(button='left', coords=(1450, 220))  # 1920 x 1200
    except Exception:
        raise
    else:
        pass

if __name__ == '__main__':
    logIn()