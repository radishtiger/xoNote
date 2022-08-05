# Basic setting
import os
import pyperclip as cb
import win32com.client as win32  # 모듈 임포트
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import win32com.client as win32
from time import sleep
from hwp_copy import en_location, content_copy, location_pair_generate
from hwp_writing import find_text, insert_text


hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")  # 한/글 실행하기
hwp.RegisterModule("FilePathCheckDLL", 'SecurityModule')
hwp.XHwpWindows.Item(0).Visible = True  # 백그라운드 숨김 해제
hwp.HAction.Run("FileFullScreen")

# 시험지 정보 받아오기
test_path = 'C:/Users/mye39/Desktop/mooho/AI/hwp_project/pyhwp/documents'
test_files = os.listdir(test_path)

# 학생 오답노트 파일 정보 받아오기
student_path = 'C:/Users/mye39/Desktop/mooho/AI/hwp_project/pyhwp/documents/오답노트'




## 틀렸다고 가정한 문항
xoDict = {'송무호' : {6 : (1,2), 7:(2, 5)}, '이재욱' : {8:(5,10)}}
# 의미 = 송무호 : 6회차 1,2번, 
#               7회차 2,5번 문항
# 이재욱 : 8회차 5,10번 문항










student_files = os.listdir(student_path)

# 중복되는 학생이 있는지 체크

forgotten_student_list = []

# 학생별 오답노트 문제배치를 반복
for name, xo_info in xoDict.items():
    find_student_file = 0
    
    # 학생 오답노트 파일 이름 찾기 
    for file in student_files:
        if name in file:
            student_file = file
            find_student_file = 1
            print(f"{name} 학생 오답파일 찾기 완료.     파일 이름:{student_file}")
            break
    if find_student_file == 0:
        forgotten_student_list.append(name)
        print(f"{name} 학생의 오답노트 파일은 존재하지 않습니다.\n다음 학생으로 넘어갑니다.")
        continue
    student_file_path = os.path.join(student_path, student_file)
    
    
    # 시험지별 문항 옮겨넣기
    for testnum, incorrect_probs_tuple in xo_info.items():
        test_index = str(testnum) + '회' # 테스트 회차
        for test in test_files:
            if test_index in test:
                test_file = test
                print(f"{test_index} 시험지 open.      파일 이름:", test_file, "\n")
                break
        test_file_path = os.path.join(test_path, test_file)
        hwp.Open(test_file_path) # 테스트지 열기 
        en_list = en_location(hwp) # refer to hwp_copy.py
        print("en_location done. \ten_list num:",len(en_list))
        location_pair = location_pair_generate(en_list)
        for i, probs in enumerate(incorrect_probs_tuple):
            hwp.Open(test_file_path) # 테스트지 열기 
            content_copy(hwp, location_pair[probs-1]) # 틀린 문제 복사
            hwp.Run('FileClose')
            hwp.Open(student_file_path) # 학생 파일 열기
            find_text(hwp, test_index) # 회차 찾고 엔터
            hwp.Run("BreakPara") # 엔터 또 친 다음에
            hwp.Run('Paste') # 복사한거 붙여넣기
            # if i!=0:
            #     insert_text(hwp, 'breakcolumn Start')
            #     hwp.Run("BreakColumn")
            #     insert_text(hwp, 'breakcolumn End')
            hwp.Run("FileSave")
            hwp.Run('FileClose')
        
        
        