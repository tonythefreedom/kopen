import time
from re import template
import win32con
import win32api
import win32gui
import ctypes
import gspread
from google.oauth2.gdch_credentials import ServiceAccountCredentials
from pywinauto import clipboard
import os
import openai
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import ctypes
import pyautogui

kakao_opentalk_name = '우리동네국민상회 안산초지점'

PBYTE256 = ctypes.c_ubyte * 256
_user32 = ctypes.WinDLL("user32")
GetKeyboardState = _user32.GetKeyboardState
SetKeyboardState = _user32.SetKeyboardState
PostMessage = win32api.PostMessage
SendMessage = win32gui.SendMessage
FindWindow = win32gui.FindWindow
IsWindow = win32gui.IsWindow
GetCurrentThreadId = win32api.GetCurrentThreadId
GetWindowThreadProcessId = _user32.GetWindowThreadProcessId
AttachThreadInput = _user32.AttachThreadInput

MapVirtualKeyA = _user32.MapVirtualKeyA
MapVirtualKeyW = _user32.MapVirtualKeyW

MakeLong = win32api.MAKELONG
w = win32con

api_key = os.environ['API_KEY']
credentials_json = "C:\\top-opus-433400-m0-4f3b1de4a55f.json"

input_path = "C:\\kopen\\buy"


# 디렉토리가 존재하지 않으면 새로 생성
if not os.path.exists(input_path):
    os.makedirs(input_path)
    print(f"디렉토리 '{input_path}'가 생성되었습니다.")
else:
    print(f"디렉토리 '{input_path}'가 이미 존재합니다.")

if not api_key:
    print("환경 변수 'OPENAI_API_KEY'를 설정해주세요.")
    exit(1)

def PostKeyEx(hwnd, key, shift, specialkey):
    if IsWindow(hwnd):

        ThreadId = GetWindowThreadProcessId(hwnd, None)

        lparam = MakeLong(0, MapVirtualKeyA(key, 0))
        msg_down = w.WM_KEYDOWN
        msg_up = w.WM_KEYUP

        if specialkey:
            lparam = lparam | 0x1000000

        if len(shift) > 0:
            pKeyBuffers = PBYTE256()
            pKeyBuffers_old = PBYTE256()

            SendMessage(hwnd, w.WM_ACTIVATE, w.WA_ACTIVE, 0)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, True)
            GetKeyboardState(ctypes.byref(pKeyBuffers_old))

            for modkey in shift:
                if modkey == w.VK_MENU:
                    lparam = lparam | 0x20000000
                    msg_down = w.WM_SYSKEYDOWN
                    msg_up = w.WM_SYSKEYUP
                pKeyBuffers[modkey] |= 128

            SetKeyboardState(ctypes.byref(pKeyBuffers))
            time.sleep(0.01)
            PostMessage(hwnd, msg_down, key, lparam)
            time.sleep(0.01)
            PostMessage(hwnd, msg_up, key, lparam | 0xC0000000)
            time.sleep(0.01)
            SetKeyboardState(ctypes.byref(pKeyBuffers_old))
            time.sleep(0.01)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, False)

        else:
            SendMessage(hwnd, msg_down, key, lparam)
            SendMessage(hwnd, msg_up, key, lparam | 0xC0000000)

def SendReturn(hwnd):
    time.sleep(1)
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

def open_chatroom(chatroom_name):
    hwndkakao = win32gui.FindWindow(None, "카카오톡")
    handkakao_edit1 = win32gui.FindWindowEx(hwndkakao, None, "EVA_ChildWindow", None)
    handkakao_edit2_1 = win32gui.FindWindowEx(handkakao_edit1, None, "EVA_ChildWindow", None)
    handkakao_edit2_2 = win32gui.FindWindowEx(handkakao_edit1, handkakao_edit2_1, "EVA_ChildWindow", None)
    handkakao_edit3 = win32gui.FindWindowEx(handkakao_edit2_2, handkakao_edit2_1, "Edit", None)
    time.sleep(2)
    win32api.SendMessage(handkakao_edit3, win32con.WM_SETTEXT, 0, chatroom_name)
    time.sleep(2)
    SendReturn(handkakao_edit3)
    time.sleep(2)


def chat_with_gpt(api_key: str, message: str, context: str) -> str:
    """
    OpenAI의 ChatGPT API를 호출하여 사용자의 메시지와 컨텍스트에 따라 응답합니다.

    :param api_key: OpenAI API 키
    :param message: 사용자 입력 메시지
    :param context: ChatGPT에게 전달할 시스템 컨텍스트
    :return: ChatGPT의 응답 메시지
    """
    openai.api_key = api_key

    try:
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": context},
                {"role": "user", "content": message}
            ]
        )
        # 응답 메시지 반환
        return response['choices'][0]['message']['content']
    except Exception as e:
        return f"Error: {e}"

def authenticate_google_sheets(credentials_json: str):
    """
    Google Sheets API 인증 함수

    :param credentials_json: 인증 파일 경로
    :return: 인증된 서비스 객체
    """
    # 인증 범위 설정
    scope = ['https://www.googleapis.com/auth/spreadsheets']

    # 인증 객체 생성
    creds = Credentials.from_service_account_file(credentials_json, scopes=scope)

    # 인증된 서비스 객체 반환
    service = build('sheets', 'v4', credentials=creds)
    return service

def parse_text_to_list(text: str) -> list:
    """
    텍스트를 2차원 리스트로 변환하는 함수

    :param text: 입력 텍스트 (줄바꿈으로 행 구분, 쉼표로 열 구분)
    :return: 2차원 리스트
    """
    return [line.split(",") for line in text.strip().split("\n")]

def write_to_google_sheets(spreadsheet_url: str, raw_data: str, text_data: str):
    """
    구글 스프레드시트에 데이터를 작성하는 함수

    :param spreadsheet_url: 스프레드시트 URL
    :param raw_data: Sheet1에 입력할 데이터 (텍스트 형태)
    :param text_data: Sheet2에 입력할 데이터 (텍스트 형태)
    :param credentials_json: 인증 파일 경로
    """
    # 인증된 서비스 객체 생성
    service = authenticate_google_sheets(credentials_json)

    # 스프레드시트 ID 추출
    spreadsheet_id = spreadsheet_url.split('/d/')[1].split('/')[0]

    # Sheet1에 raw_data 작성
    raw_data_list = parse_text_to_list(raw_data)
    sheet1_range = 'Sheet1!A1'  # Sheet1 시작 범위
    raw_body = {
        'values': raw_data_list
    }
    service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=sheet1_range,
        valueInputOption="RAW",
        body=raw_body
    ).execute()

    # Sheet2에 text_data 작성
    text_data_list = parse_text_to_list(text_data)
    sheet2_range = 'Sheet2!A1'  # Sheet2 시작 범위
    text_body = {
        'values': text_data_list
    }
    service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=sheet2_range,
        valueInputOption="RAW",
        body=text_body
    ).execute()


def remove_quotes(input_string):
    """
    문자열에서 모든 " 를 제거합니다.

    Parameters:
        input_string (str): 입력 문자열

    Returns:
        str: "가 제거된 문자열
    """
    return input_string.replace('"', '')


def main():
    open_chatroom(kakao_opentalk_name)
    time.sleep(2)
    hwndMain = win32gui.FindWindow(None, kakao_opentalk_name)
    hwndListControl = win32gui.FindWindowEx(hwndMain, None, "EVA_VH_ListControl_Dblclk", None)

    # PostKeyEx(hwndListControl, ord('A'), [w.VK_CONTROL], False)
    # time.sleep(2)
    # PostKeyEx(hwndListControl, ord('C'), [w.VK_CONTROL], False)
    # time.sleep(5)

    # Ctrl + S 눌러서 저장 시도
    PostKeyEx(hwndListControl, ord('S'), [w.VK_CONTROL], False)
    time.sleep(1)

    # 파일 내용 정리 후 다시 저장, 기존 임시 파일 삭제
    if(filter_text_lines(input_path, input_path + "\\" + kakao_opentalk_name)) :
        delete_kakaotalk_files(input_path)



    # 저장 닫기
    PostKeyEx(hwndListControl, 0x0D, [w.VK_RETURN], False)

    time.sleep(2)
    
    # 닫기 버튼 찾아 닫기
    # pyautogui.moveTo(find_close_button(hwndListControl))
    pyautogui.click(find_close_button(hwndListControl))

    time.sleep(2)

    # 대화 버튼 찾기
    pyautogui.moveTo(find_dialog_button(hwndListControl))
    pyautogui.click(find_dialog_button(hwndListControl))

    time.sleep(1)

    # 대화내용 메뉴 찾기
    pyautogui.moveTo(find_clean_dialog_menu(hwndListControl))

    time.sleep(1)

    # 대화내용 지우기 버튼 찾기
    pyautogui.moveTo(find_clean_dialog_button(hwndListControl))
    pyautogui.click(find_clean_dialog_button(hwndListControl))

    time.sleep(1)

    # 대화내용 지우기 확인 버튼 찾기
    pyautogui.moveTo(find_clean_dialog_confirm(hwndListControl))
    # pyautogui.click(find_clean_dialog_confirm(hwndListControl))


    # ctext = clipboard.GetData()
    # print(ctext)
    # response = chat_with_gpt(api_key, "system 으로 보낸 content 의 내용을 보고 주문내역을 구글 스프레드시트에 붙여 넣을 수 있게 csv 파일 타입으로 정리해줘. 주문자, 상품명, 수량 을 테이블 헤더로. 한글로만 답변해. 주문내역 외 기타 다른 대답은 생성 하지마", ctext)
    #
    # print(remove_quotes(response))
    # if response:
    #     write_to_google_sheets("https://docs.google.com/spreadsheets/d/1P-4rqijcsLK7O2DoL2kuEj4tMdP32dC6hpankFKnWv0/edit?usp=sharing",ctext, remove_quotes(response))


user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

# 활성화된 창의 핸들을 가져오는 함수
def get_foreground_window():
    return user32.GetForegroundWindow()

# 윈도우 창의 제목을 가져오는 함수
def get_window_title(hwnd):
    length = user32.GetWindowTextLengthW(hwnd) + 1  # 제목의 길이를 구한 후 +1
    buffer = ctypes.create_unicode_buffer(length)  # 버퍼 생성
    user32.GetWindowTextW(hwnd, buffer, length)  # 창 제목을 버퍼에 저장
    return buffer.value


# 마우스 좌표를 저장할 구조체 정의
class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

# 마우스 좌표를 가져오는 함수
def get_mouse_position():
    point = POINT()
    user32.GetCursorPos(ctypes.byref(point))  # GetCursorPos 호출
    return point.x, point.y


class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long),
                ("top", ctypes.c_long),
                ("right", ctypes.c_long),
                ("bottom", ctypes.c_long)]

# 닫기 버튼 찾기
def find_close_button(hwnd):

    # 윈도우 RECT 가져오기
    rect = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))

    width = rect.right - rect.left
    height = rect.bottom - rect.top
    x = rect.left + width // 2 + 55
    y = rect.top + height // 4 * 3.4

    return x, y


# 대화내용 버튼 찾기
def find_dialog_button(hwnd):

    # 윈도우 RECT 가져오기
    rect = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))

    width = rect.right - rect.left
    x = rect.left + width - 25
    y = rect.top - 30

    return x, y

# 대화내용 메뉴 찾기
def find_clean_dialog_menu(hwnd):

    # 윈도우 RECT 가져오기
    rect = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))

    width = rect.right - rect.left
    x = rect.left + width - 25
    y = rect.top + 242

    return x, y


# 대화지우기 버튼 찾기
def find_clean_dialog_button(hwnd):

    # 윈도우 RECT 가져오기
    rect = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))

    width = rect.right - rect.left
    x = rect.left + width + 200
    y = rect.top + 290

    return x, y


# 대화지우기 확인 버튼 찾기
def find_clean_dialog_confirm(hwnd):

    # 윈도우 RECT 가져오기
    rect = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))

    width = rect.right - rect.left
    height = rect.bottom - rect.top
    x = rect.left + width // 2 - 80
    y = rect.top + height // 4 * 3.2

    return x, y


def filter_text_lines(input_file_path, output_file_path):
    """
    입력 파일에서 '['로 시작하는 행을 제외한 모든 텍스트를 삭제하고 결과를 새로운 파일에 저장합니다.

    :param input_file_path: 원본 파일 경로 (str)
    :param output_file_path: 출력 파일 경로 (str)
    """
    try:
        # 출력 디렉토리 생성
        output_dir = os.path.dirname(output_file_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # 파일 읽기
        with open(input_file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()

        # '['로 시작하는 행만 필터링
        filtered_lines = [line for line in lines if line.lstrip().startswith('[')]

        # 결과를 새로운 파일에 저장
        with open(output_file_path, 'w', encoding='utf-8') as file:
            file.writelines(filtered_lines)

        print(f"Filtered text saved to {output_file_path}")
    except Exception as e:
        print(f"An error occurred: {e}")


def delete_kakaotalk_files(directory_path):
    """
    지정된 디렉토리에서 'KakaoTalk'로 시작하는 모든 파일을 삭제합니다.

    :param directory_path: 디렉토리 경로 (str)
    """
    try:
        if not os.path.exists(directory_path):
            print(f"Directory does not exist: {directory_path}")
            return

        for filename in os.listdir(directory_path):
            if filename.startswith("KakaoTalk"):
                file_path = os.path.join(directory_path, filename)
                if os.path.isfile(file_path):
                    os.remove(file_path)
                    print(f"Deleted file: {file_path}")
    except Exception as e:
        print(f"An error occurred while deleting files: {e}")

if __name__ == "__main__":
    print("Starting script...")
    main()