
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import win32com.client as win32
from time import sleep
root = Tk()
filename = askopenfilename()
root.destroy()

hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.RegisterModule('FileAuto') # 보안 승인창 뜨지 않도록 모듈 등록
hwp.Open(filename)
hwp.XHwpWindows.Item(0).Visible=True
hwp.HAction.Run("FileFullScreen")

Action = hwp.CreateAction("EquationModify")
Set = Action.CreateSet()
ItemSet = Set.CreateItemSet("EqEdit", "EqEdit")
Action.GetDefault(ItemSet)
extracted_equation = ItemSet.Item("String")
print(extracted_equation)

"""모든 수식 텍스트 차례로 dict로 얻기.
키는 (List, Para, Pos), 값은 eqn_string"""


def extract_eqn(hwp):  # 이전 포스팅에서 소개한, 수식 추출방법을 함수로 정의
    Act = hwp.CreateAction("EquationModify")
    Set = Act.CreateSet()
    Pset = Set.CreateItemSet("EqEdit", "EqEdit")
    Act.GetDefault(Pset)
    return Pset.Item("String")
    

eqn_dict = {}  # 사전 형식의 자료 생성 예정
ctrl = hwp.HeadCtrl  # 첫 번째 컨트롤(HeadCtrl)부터 탐색 시작.


while ctrl != None:  # 끝까지 탐색을 마치면 ctrl이 None을 리턴하므로.
    nextctrl = ctrl.Next  # 미리 nextctrl을 지정해 두고,
    if ctrl.CtrlID == "eqed":  # 현재 컨트롤이 "수식eqed"인 경우
        position = ctrl.GetAnchorPos(0)  # 해당 컨트롤의 좌표를 position 변수에 저장
        position = position.Item("List"), position.Item("Para"), position.Item("Pos")
        hwp.SetPos(*position)  # 해당 컨트롤 앞으로 캐럿(커서)을 옮김
        hwp.FindCtrl()  # 해당 컨트롤 선택
        eqn_string = extract_eqn(hwp)  # 문자열 추출
        eqn_dict[position] = eqn_string  # 좌표가 key이고, 수식문자열이 value인 사전 생성
    ctrl = nextctrl  # 다음 컨트롤 탐색
hwp.Run("Cancel")  # 완료했으면 선택해제

for key, val in eqn_dict.items():
    print("key:", key)
    print("val:", val)