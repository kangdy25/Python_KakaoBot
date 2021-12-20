import time, win32con, win32api, win32gui
from io import BytesIO
import win32clipboard 
from PIL import Image
from datetime import datetime, timedelta
from pynput.keyboard import Key, Controller

# # 카톡창 이름, (활성화 상태의 열려있는 창)
kakao_opentalk_name = '💒 행복한교회 청년부 공지방 📢'

# # 채팅방에 메시지 전송
def kakao_sendtext(chatroom_name, text):
    # # 핸들 _ 채팅방
    hwndMain = win32gui.FindWindow( None, chatroom_name)
    hwndEdit = win32gui.FindWindowEx( hwndMain, None, "RICHEDIT50W", None)
    # hwndListControl = win32gui.FindWindowEx( hwndMain, None, "EVA_VH_ListControl_Dblclk", None)

    win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, text)
    pressPaste()
    time.sleep(0.01)
    returnPaste()

    print("sent") 

    returnPaste()

# # 붙여넣기
def pressPaste():
    keyboard = Controller()

    keyboard.press(Key.ctrl)
    keyboard.press('v')
    time.sleep(0.01)
    keyboard.release(Key.ctrl)
    keyboard.release('v')

# # 엔터
def returnPaste():
    keyboard = Controller()
    keyboard.press(Key.enter)
    time.sleep(0.01)
    keyboard.release(Key.enter)

# # 엔터
def SendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

# # 채팅방 열기
def open_chatroom(chatroom_name):
    # # 채팅방 목록 검색하는 Edit (채팅방이 열려있지 않아도 전송 가능하기 위하여)
    hwndkakao = win32gui.FindWindow(None, "카카오톡")
    hwndkakao_edit1 = win32gui.FindWindowEx( hwndkakao, None, "EVA_ChildWindow", None)
    hwndkakao_edit2_1 = win32gui.FindWindowEx( hwndkakao_edit1, None, "EVA_Window", None)
    hwndkakao_edit2_2 = win32gui.FindWindowEx( hwndkakao_edit1, hwndkakao_edit2_1, "EVA_Window", None)
    hwndkakao_edit3 = win32gui.FindWindowEx( hwndkakao_edit2_2, None, "Edit", None)

    # # Edit에 검색 _ 입력되어있는 텍스트가 있어도 덮어쓰기됨
    win32api.SendMessage(hwndkakao_edit3, win32con.WM_SETTEXT, 0, chatroom_name)
    time.sleep(1)   # 안정성 위해 필요
    SendReturn(hwndkakao_edit3)
    time.sleep(1)

# # 클립보드에 저장하기
def send_to_clipboard(clip_type, data): 
    win32clipboard.OpenClipboard() 
    win32clipboard.EmptyClipboard() 
    win32clipboard.SetClipboardData(clip_type, data) 
    win32clipboard.CloseClipboard() 

# 말씀 요약 파일 불러오기
def fileCopy():
    yesterday = datetime.today() - timedelta(1)
    Correctyesterday= yesterday.strftime("%Y_%m_%d")

    # image = Image.open("/ToSend_img" + str(Correctyesterday) + '.jpg') 
    image = Image.open("C:/Programming/Python/Python_AutoWork/ToSend_img" + "/" + str(Correctyesterday) + ".jpg") 

    output = BytesIO() 
    image.convert("RGB").save(output, "BMP") 
    data = output.getvalue()[14:] # The file header off-set of BMP is 14 bytes. 
    output.close() 
    send_to_clipboard(win32clipboard.CF_DIB, data)

def main():
    open_chatroom(kakao_opentalk_name)  # 채팅방 열기
    text = "이번 주 말씀 요약입니다~~ \n청년부 모든 사람이 한 주도 말씀 기억하시면서 승리하시길 소망합니다"
    fileCopy() # 이미지 클립보드 복사
    kakao_sendtext(kakao_opentalk_name, text)    # 메시지 전송

if __name__ == '__main__':
    main()
    