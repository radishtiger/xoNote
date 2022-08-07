import os
from time import sleep
import pyperclip as cb

import win32com.client as win32  # 모듈 임포트

from hwp_copy import en_location, content_copy, location_pair_generate
from hwp_writing import find_text, insert_text

hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")  # 한/글 실행하기
hwp.RegisterModule("FilePathCheckDLL", 'SecurityModule')
hwp.XHwpWindows.Item(0).Visible = True  # 백그라운드 숨김/해제
# hwp.HAction.Run("FileFullScreen")

def change_font_size(hwp, size):
    """
    change the font size of its line
    """
    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    option = hwp.HParameterSet.HCharShape
    option.Height = size
    hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
    
def go_up(hwp):
    hwp.Run("MoveLineUp")
    
def go_down(hwp):
    hwp.Run("MoveLineDown")
    
    
# en_file = 'C:/Users/mye39/Desktop/mooho/AI/hwp_project/pyhwp/documents/Endnote_exercise_another.hwp'
# hwp.Open(en_file)
# en_list = en_location(hwp)

# print("en_list :", en_list)


def column_head_format(hwp):
    hwp.MovePos(2)
    prev_pos = hwp.GetPos()
    testPos = []
    testnum = 0
    while True:
        testnum +=1
        find_text(hwp, '['+str(testnum)+'회]')
        cur_pos = hwp.GetPos()
        if prev_pos[1]!=cur_pos[1]:
            testPos.append(cur_pos)
        else:
            break
    
    
        
    
def fig_size_control(hwp):
    raise NotImplementedError