from selenium import webdriver  # 셀레니움 라이브러리
from bs4 import BeautifulSoup  # 뷰티풀숲 라이브러리
from base64 import b64decode  # base64 decode를 이용하여 파일 저장을 함
from selenium.webdriver.edge.options import Options as EdgeOptions
import time  # time 슬립용 라이브러리
import re  # 정규식 사용을 위한 라이브러리
from tkinter import *
from tkinter import messagebox as mb
from tkinter import filedialog as fd
from tkinter import ttk  # progress bar용
from threading import Thread
from selenium.webdriver.common.by import By  #deprecated 해결용

options = ''
driver = ''
page_bar = ''
pages = ''
owner = ''  # 블로그 주인 아이디 저장
directory = ''  # 경로 설정

w = Tk()

def make_pdf(h_driver, post_num, date, title):  # print to pdf는 headless 모드에서만 실행된다
    global owner  # 블로그 주인장 참조를 위한 global 변수
    global directory

    # 헤더와 푸터..
    h_driver.get(f"https://blog.naver.com/PostPrint.naver?blogId={owner}&logNo={post_num}")
    pdf = b64decode(h_driver.print_page())
    with open(fr"{directory}/{date} {title} {owner}[{post_num}].pdf", "wb") as f:  # f설정과 r설정을 동시에 넣음
        f.write(pdf)


def find_post_list():
    global options, driver, page_bar, pages, owner
    options = webdriver.EdgeOptions()
    options.add_argument('headless')
    driver = r"C:\Users\jay\Desktop\edgedriver_win64\msedgedriver.exe" ################### 여기 수정해야됨
    driver = webdriver.Edge(driver, options=options)
    driver.get(f"https://blog.naver.com/PostList.naver?blogId={owner}&categoryNo=0&from=postList")
    list_open = driver.find_element(by=By.XPATH, value='//*[@id="category-name"]/div/table[2]/tbody/tr/td[2]/div/a') # 목록열기 닫기 버튼 선택
    print(list_open)
    '''
    if list_open.text == '목록열기': #목록열기로 떠있으면 목록을 열어준다
        list_open.send_keys()
        list_open = driver.find_element(by=By.XPATH, value='//*[@id="category-name"]/div/table[2]/tbody/tr/td[2]/div/a') # 목록열기 닫기 버튼 선택
        print(list_open.text)
    '''
    page_bar = driver.find_element_by_class_name('blog2_paginate')  # 페이지번호 틀
    pages = page_bar.find_elements_by_css_selector('a')  # 각 페이지 숫자 버튼
    print('페이지 숫자버튼들', pages)


def print_page_export():
    global options, driver, page_bar, pages
    cnt = 0  # 페이지 넘김을 제어하기 위한 cnt 변수
    find_post_list()  # 일단 게시글의 글 목록 부분의 DOM을 얻음.
    #return  ######################################### 여기도 수정
    while True:
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        li = soup.select('#listTopForm > table > tbody > tr > td.title > div > span')  # 뷰티풀숲으로 리스트상의 모든 제목들 선택
        li_date = soup.select('#listTopForm > table > tbody > tr > td.date > div > span')  # 리스트상의 모든 날짜들 선택
        headless_driver = r"C:\Users\jay\Desktop\edgedriver_win64\msedgedriver.exe"
        # 또 다른 웹드라이버를 만들어줘야지 초반에 게시글 번호 따려고 한 드라이버를 똑같이 써서 다시 get을 해버리면 dom이 날아가버림
        headless_driver = webdriver.Edge(headless_driver, options=options)

        for i in range(len(li)):
            link = li[i].find('a')['href'] # 게시글 링크 확인
            post_num = re.search('logNo=(.+?)&category', link).group(1)  # 정규표현식 라이브러리 사용 # 게시글 번호 확인
            title = li[i].text # 게시글 제목 확인
            title = re.sub('[\/:*?"<>|]','',title)  # 파일에 들어갈 수 없는 특수문자 제거
            date = li_date[i].text  # 게시글 작성 날짜 확인
            print(date, post_num, title)
            make_pdf(headless_driver, post_num, date, title)
        print()

        #if len(pages) == 0:  # 페이지가 하나 뿐이라면 0인 이유는 페이지를 a 링크로 찾는데 하나 일 때는 a 링크 설정이 안걸린다.
            #break
        if pages[-1].text != '다음' and cnt == len(pages):  # 조건문에서 pages를 건들기때문에 dom 갱신 순서 조심
            break
        if pages[cnt].text == '다음':
            pages[cnt].send_keys('\n')
            cnt = 1
        else:
            pages[cnt].send_keys('\n')
            cnt += 1
        time.sleep(2)
        page_bar = driver.find_element_by_class_name('blog2_paginate')  # for 문 돌때마다 dom이 날라가므로 항상 다시 구해줘야됨
        pages = page_bar.find_elements_by_css_selector('a')
    mb.showinfo('Completed', '완료')
    driver.close()  #다 끝났으니 웹드라이버를 종료해줌

def stop():
    global driver
    driver.close()

def press():
    global owner
    text = entry1.get()
    if text == '':
        mb.showerror('오류', '아이디를 입력해주세요.')
        new_window = Toplevel(w)
        new_window.geometry('300x100')
        new_window.title('변환 중')
        pb = ttk.Progressbar(new_window, orient='horizontal', mode='determinate', length=150)
        pb.pack(pady=20)
        btn_stop = Button(new_window, text='중단', command= lambda: [new_window.destroy(), stop()])
        btn_stop.pack()

    else:
        if mb.askyesno('확인', f'입력하신 아이디는 {text} 입니다. 변환하시려면 OK를 누르세요.'):
            owner = text
            th1 = Thread(target=print_page_export)
            th1.setDaemon(True)
            th1.start()
            #print_page_export()
            return


def on_close():
    if mb.askyesno('종료', '종료하시겠습니까?'):
        w.destroy()


def ask_dir():
    global directory
    directory = fd.askdirectory(parent=w, initialdir="/", title='저장 경로 설정')
    print(directory)

w.title("블로그 PDF변환")
w.geometry('400x200')
label1 = Label(w, text='블로그 모든 게시글 PDF 변환', background='white')
label1.pack(pady=20)
label2 = Label(w, text='블로그 주인 아이디를 입력하세요.', background='white')
label2.pack(pady=5)
entry1 = Entry(w)
entry1.pack()
btn = Button(w, text='변환', command=press)
btn.pack(pady=20)
info = Label(w, text='ohito@naver.com')
info.pack()
btnSave = Button(w, text='저장 경로', command=ask_dir)
btnSave.place(x=230, y=131)
w.protocol('WM_DELETE_WINDOW', on_close)
w.mainloop()