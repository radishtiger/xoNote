import pyperclip as cb
import win32com.client as win32  # 모듈 임포트
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import win32com.client as win32
from time import sleep
import os
from hwp_copy import en_location, content_copy, location_pair_generate



def find_text(hwp, text):
    # 해당 text 찾고 Esc 누르는 것까지.
    # hwp.MovePos(2)
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    option = hwp.HParameterSet.HFindReplace
    option.FindString = text
    option.IgnoreMessage = 1
    hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    # print("test searching end")
    
    hwp.Run("Cancel")
    # print("cancel end")

    
    
def insert_text(hwp, text):
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    
# root = Tk()
# root.destroy()

# test_path = 'C:/Users/mye39/Desktop/mooho/AI/hwp_project/pyhwp/documents'
# test_files = os.listdir(test_path)
# test_index = str(6) + '회'
# for file in test_files:
#     if test_index in file:
#         test_file_name = file
#         break
# test_file_name = os.path.join(test_path, test_file_name)
# print("test file_num : ", test_file_name)

# hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")  # 한/글 실행하기
# hwp.RegisterModule("FilePathCheckDLL", 'SecurityModule')

# hwp.XHwpWindows.Item(0).Visible = False  # 백그라운드 숨김 해제
# hwp.Open(test_file_name)

# en_list = en_location(hwp)
# en_pair = location_pair_generate(en_list)
# content_copy(hwp, en_pair[3])
# print("copy complete")
# print("content:\n", cb.paste())
# # hwp.Run("FileClose")
# # print("file close")
# sleep(1)

# print("new_hwp execute")
# # hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")  # 한/글 실행하기
# # hwp.RegisterModule("FilePathCheckDLL", 'SecurityModule')
# # hwp.HAction.Run("FileFullScreen")

# student_path = 'C:/Users/mye39/Desktop/mooho/AI/hwp_project/pyhwp/documents/오답노트'
# student_file_path = os.path.join(student_path, '오답노트_송무호 - 복사본.hwp')
# print("student_file_name :", student_file_path)

# # hwp.HAction.Run("FileFullScreen")
# hwp.Open(student_file_path)
# hwp.XHwpWindows.Item(0).Visible = True  # 백그라운드 숨김 해제

# find_text(hwp, '[5회]')

# # hwp.Run("FileClose")
# # print("file close")



