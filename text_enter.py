# https://martinii.fun/238

import pyperclip as cb
import win32com.client as win32  # 모듈 임포트
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import win32com.client as win32
from time import sleep
root = Tk()
filename = askopenfilename()
root.destroy()

hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")  # 한/글 실행하기
hwp.XHwpWindows.Item(0).Visible = True  # 백그라운드 숨김 해제
hwp.Open(filename)
hwp.HAction.Run("FileFullScreen")

def insert_text(text):
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

cb.copy("Hello World")
hwp.Run("Paste")

#누름틀(필드) 이용하여 텍스트 입력
# ① 필드 생성
hwp.CreateField(
    Direction="입력칸",
    memo="텍스트 입력",
    name="textarea")

# ② 텍스트 입력
hwp.PutFieldText(
    "textarea",
    "Hello World")

# # ③ 필드 삭제
# hwp.Run("DeleteField")


# 여러 개의 누름틀에 동시에 입력하기
필드리스트 = ["국어점수", "영어점수", "수학점수", "과학점수"]

점수리스트 = ["90", "95", "80", "60"]

hwp.PutFieldText(
    "\x02".join(필드리스트),
    "\x02".join(점수리스트))