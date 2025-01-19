import ctypes

user32 = ctypes.WinDLL("user32")

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


# 메인 메뉴 버튼 찾기
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
    y = rect.top + 50

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
