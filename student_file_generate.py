import os
import pyperclip as cb
import win32com.client as win32  # 모듈 임포트
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import win32com.client as win32
from time import sleep
from hwp_copy import en_location, content_copy, location_pair_generate
from hwp_writing import find_text, insert_text

def student_file_generate(student_name, file_location):
    file_name = '오답노트_'+str(student_name)+'.hwp'
    hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")  # 한/글 실행하기
    hwp.RegisterModule("FilePathCheckDLL", 'SecurityModule')
    hwp.Open('C:/Users/mye39/Desktop/mooho/AI/hwp_project/pyhwp/documents/오답노트/오답노트_예시.hwp')
    # hwp.XHwpWindows.Item(0).Visible = True  # 백그라운드 숨김 해제
    hwp.SaveAs(os.path.join(file_location, file_name))
    hwp.Run("FileSave")
    hwp.Run("FileQuit")
    
# file_location = 'C:/Users/mye39/Desktop/mooho/AI/hwp_project/pyhwp/documents/오답노트'
# student_file_generate('무호', file_location)