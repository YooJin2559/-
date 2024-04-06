from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pyautogui
import os

try:
    # 4-0. word 버튼 클릭
    word_button = driver.find_element_by_xpath('word 버튼 Xpath')  # 프론트가 완성되면 word버튼 Xpath 주소 넣기
    word_button.click()

    # 4-1. 페이지에서 요약 상자 클릭
    summary_box = driver.find_element_by_xpath('요약 상자 Xpath')  # 요약 상자 안의 내용 Xpath 경로 넣기
    summary_box.click()

    # 4-2. 요약 상자 내용 전체 선택
    summary_box.send_keys(Keys.CONTROL + 'a')

    # 4-3. 요약 상자 내용 전체 복사
    summary_box.send_keys(Keys.CONTROL + 'c')

    # 4-4. Word 프로그램 실행 (5초 대기)
    pyautogui.hotkey('win', 'r') #win+r 실행 단축키 자동클릭
    time.sleep(1) #1초 대기
    pyautogui.typewrite('winword') 
    pyautogui.press('enter') #word 실행
    time.sleep(5) #자동화 속도가 워드 컴퓨터 실행 속도보다 빠를 수 있으니 5초정도 대기시간 부여(유연하게 조정 가능)


    # 4-5. Word에 요약 상자 내용 붙여넣기
    pyautogui.press('enter') #기본 템플릿(새문서)실행
    time.sleep(3)
    pyautogui.hotkey('ctrl', 'v')
    
    # 파일 이름 설정(file1,2,3,4 ... 순차적으로 숫자가 늘어나면서 파일이름 저장)
    file_name = 'file'
    file_number = 1
    while os.path.exists(file_name + str(file_number) + '.docx'):
        file_number += 1

    # 중복되지 않는 파일 이름 설정
    file_name = file_name + str(file_number)
    
    # 4-6. 요약 상자 내용 저장
    pyautogui.hotkey('ctrl', 's')
    time.sleep(1)  # 저장 다이얼로그가 나타날 때까지 대기
    pyautogui.typewrite(file_name)  # 파일 이름으로 저장
    pyautogui.press('enter')

except Exception as e:
    print("에러 발생:", e)

finally:
    # 웹 드라이버 종료
    # driver.quit()  -> 끝나고 브라우저 꺼질거면 살리면 됨.