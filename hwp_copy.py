import pyperclip as cb
import win32com.client as win32  # 모듈 임포트
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import win32com.client as win32
from time import sleep
# root = Tk()
# filename = askopenfilename()
# root.destroy()

hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")  # 한/글 실행하기
# hwp.XHwpWindows.Item(0).Visible = False  # 백그라운드 숨김 해제
# hwp.Open(filename)
# hwp.HAction.Run("FileFullScreen")



def en_location(hwp):
    en_list = []
    ctrl = hwp.HeadCtrl  # 첫 번째 컨트롤(HeadCtrl)부터 탐색 시작.
    while ctrl != None:  # 끝까지 탐색을 마치면 ctrl이 None을 리턴하므로.
        nextctrl = ctrl.Next  # 미리 nextctrl을 지정해 두고,
        if ctrl.CtrlID == "en":  # 현재 컨트롤이 "미주en"인 경우
            position = ctrl.GetAnchorPos(0)  # 해당 컨트롤의 좌표를 position 변수에 저장
            position = (position.Item("List"), position.Item("Para"), position.Item("Pos"))
            en_list.append(position)
            
        ctrl = nextctrl  # 다음 컨트롤 탐색
    hwp.MovePos(3)
    position = hwp.GetPos()
    en_list.append(position)
    return en_list

def content_copy(hwp, location_tuple):
    start, end = location_tuple
    hwp.SetPos(*start)
    hwp.Run("Select")
    hwp.SetPos(*end)
    hwp.Run("Copy")
    
    
def total_content_copy(hwp, location):
    n = len(location)
    content = []
    if n==0:
        print("0 endnote.")
        return
    for i in range(n-1):
        hwp.SetPos(*location[i])
        hwp.Run("Select")
        hwp.SetPos(*location[i+1])
        hwp.Run("Copy")
        text = cb.paste()
        content.append(text)
           
    return content
    
# hwp.Run("Cancel")  # 완료했으면 선택해제

# location_list = en_location(hwp)
# content_list = total_content_copy(hwp, location_list)

# print("location_num :", len(location_list))
# for i in location_list:
#     print("location :", i)

# print("content_num:", len(content_list))
# for i in content_list:
#     print(i)
#     print("------------------------")

def location_pair_generate(location_list):
    n = len(location_list)
    if n<2:
        print("not enough locations. You have",n,"EndNotes")
        return None
    return [(location_list[i], location_list[i+1]) for i in range(n-1)]