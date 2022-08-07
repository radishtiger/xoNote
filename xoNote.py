# Basic setting
import os
from time import sleep
import pyperclip as cb

import win32com.client as win32  # 모듈 임포트

from hwp_copy import en_location, content_copy, location_pair_generate
from hwp_writing import find_text, insert_text
from FormatChange import change_font_size

from openpyxl import load_workbook
from excel_read import student_info_dict

from student_file_generate import student_file_generate

grade = 1
if grade == 1:
    file_location = 'C:/Users/mye39/Desktop/mooho/AI/hwp_project/pyhwp/documents/오답노트/student_info/(22년)고1 오전클리닉 TEST.xlsx'
    test_path = 'C:/Users/mye39/Desktop/mooho/AI/hwp_project/pyhwp/documents/오답노트/[21년 여름] 수학(하) Daily test'
    student_path = 'C:/Users/mye39/Desktop/mooho/AI/hwp_project/pyhwp/documents/오답노트/고1_automation'
elif grade==2:
    file_location = 'C:/Users/mye39/Desktop/mooho/AI/hwp_project/pyhwp/documents/오답노트/student_info/(22년)고2 오전클리닉 TEST.xlsx'
    test_path = 'C:/Users/mye39/Desktop/mooho/AI/hwp_project/pyhwp/documents/오답노트/[21년 여름]수2 Daily test'
    student_path = 'C:/Users/mye39/Desktop/mooho/AI/hwp_project/pyhwp/documents/오답노트/고2_automation'
    
xoDict = {}
for i in range(1, 16):
    xoDict = student_info_dict(file_location, i, xoDict)
    if i==1:
        print("학생 정/오답 정보 로드 완료 회차:")
    print(i,"", end=' ')
print("\nload complete\n")    

# for key, val in xoDict.items():
#     print("name:", key)
#     print("\tincorrect info:")
#     for testnum, probs in val.items():
#         print("\t\t", testnum, " : ",probs)
#     print()


# 시험지 정보 받아오기
test_files = os.listdir(test_path)

# 학생 오답노트 폴더 정보 받아오기

student_files = os.listdir(student_path)

# 오답노트 없는 학생에 대하여 파일 생성하기

hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")  # 한/글 실행하기
hwp.RegisterModule("FilePathCheckDLL", 'SecurityModule')
hwp.XHwpWindows.Item(0).Visible = True  # 백그라운드 숨김/해제
# hwp.HAction.Run("FileFullScreen")

forgotten_student_list = []

# 학생별 오답노트 문제배치를 반복
for name, xo_info in xoDict.items():
    find_student_file = 0
    if sum([name in stu_file for stu_file in student_files])==0:
        print(name,"학생 오답노트 파일 생성.")
        student_file_generate(name, student_path)
    
    student_files = os.listdir(student_path)
    # 학생 오답노트 파일 이름 찾기 
    for file in student_files:
        if name in file:
            student_file = file
            find_student_file = 1
            print(f"\n {name} 학생 오답파일 찾기 완료.")
            student_xo_info = xoDict[name]
            for testnum, incorrect_probs in student_xo_info.items():
                print(testnum,"회차", end=' ')
                print(":\t", incorrect_probs)
                
            break
    if find_student_file == 0:
        forgotten_student_list.append(name)
        print(f"{name} 학생의 오답노트 파일은 존재하지 않습니다. 다음 학생으로 넘어갑니다.\n")
        continue
    student_file_path = os.path.join(student_path, student_file)
    
    
    print("문항배치 완료된 회차:")
    # 시험지별 문항 옮겨넣기
    for testnum, incorrect_probs_tuple in xo_info.items():
        test_index = ' ' + str(testnum) + '회' # 테스트 회차
        for test in test_files:
            if test_index in test:
                test_file = test
                break
        test_file_path = os.path.join(test_path, test_file)
        hwp.Open(test_file_path) # 테스트지 열기 
        en_list = en_location(hwp) # refer to hwp_copy.py
        # print("en_location done. \ten_list num:",len(en_list))
        location_pair = location_pair_generate(en_list)
        for i, probs in enumerate(incorrect_probs_tuple):
            hwp.Open(test_file_path) # 테스트지 열기 
            content_copy(hwp, location_pair[probs-1]) # 틀린 문제 복사
            hwp.Open(student_file_path) # 학생 파일 열기
            find_text(hwp, '['+str(testnum) + '회]') # 회차 찾고 엔터
            hwp.Run("MoveLineEnd") # 해당 줄 끝으로 이동
            hwp.Run("BreakPara") # 엔터 또 친 다음에
            hwp.Run('Paste') # 복사한거 붙여넣기
            # hwp.Run('DeleteBack')
            change_font_size(hwp, 10)
            if i==0:
                hwp.Run('DeleteBack')
            if i==len(incorrect_probs_tuple):
                hwp.Run('BreakPara')
            hwp.Run("FileSave")
        print(f"{testnum}",end=" ")
        hwp.Open(test_file_path)
        hwp.Run("FileClose")
    hwp.Open(student_file_path)
    hwp.Run("FileSave")
    hwp.Run("FileClose")
    print("\n\n")
print("forgotten_student_name :\n",forgotten_student_list)
        
        
        