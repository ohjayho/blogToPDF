'''
제작자는 오재호이며, 이메일 주소는 ohjayhoi@gmail.com 입니다.
사회복무요원 복무 중 업무를 돕기 위해 작성한 프로그램입니다.
무단 수정과 배포를 금지합니다.
'''

from selenium import webdriver  # 셀레니움 라이브러리
from selenium.webdriver.edge.options import Options as EdgeOptions  # Edge에 headless 옵션을 추가하기 위한 라이브러리
from selenium.webdriver.common.by import By  # deprecated 해결용
from selenium import webdriver  # 웹드라이버
from selenium.webdriver.edge.service import Service  # 웹드라이버 다운로드 및 웹드라이버 실행 시 cmd 창 안뜨게 service 사용
from subprocess import CREATE_NO_WINDOW  # 이 flag는 윈도우즈에서만 작동된다고 한다
from webdriver_manager.microsoft import EdgeChromiumDriverManager  # 웹드라이버 알아서 실행할때마다 다운로드
from bs4 import BeautifulSoup  # 뷰티풀숲 라이브러리
from base64 import b64decode  # base64 decode를 이용하여 파일 저장을 함
import time  # time 슬립용 라이브러리
import re  # 정규식 사용을 위한 라이브러리
from tkinter import *
from tkinter import messagebox as mb
from tkinter import filedialog as fd
from tkinter import ttk  # progress bar용
from threading import Thread  # 멀티쓰레드용
import math  # 올림 용
from pathlib import Path  # tkinter Designer용 asset 불러올 때 Path 씀

# 여기서부터 tkinter designer 용
# from tkinter import *
# Explicit imports to satisfy Flake8
from tkinter import Tk, Canvas, Entry, Button, PhotoImage  # tkinter import *을 했어도 따로 명시를 해줘야한다.


OUTPUT_PATH = Path(__file__).parent  # tkinter를 위한 이미지를 불러오는 용도
ASSETS_PATH = OUTPUT_PATH / Path("./assets")  # 상동


def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)
#여기까지 tkinter designer 끝


options = ''
driver = ''
page_bar = ''
pages = ''
owner = ''  # 블로그 주인 아이디 저장
dir = ''  # 경로 설정
web_dir = ''  # 웹드라이버 경로
pb = '' # Progress Bar 업데이트용 변수
total_post_num = 0  # 전체 게시글 수
new_window = ''  # 로딩 윈도우용 변수 ( 작업이 끝나면 닫아줘야하는데 다른 method에서 닫아주려니 참조해야됨 )
headless_driver = ''  # 새로 만들어준 두번째 웹드라이버를 참조하기 위한 변수
edge_service = ''  # 다운로드 없이 자동 다운로드 설치 Service를 참조하기 위한 변수

w = Tk()

def make_pdf(h_driver, post_num, date, title):  # print to pdf는 headless 모드에서만 실행된다
    global owner  # 블로그 주인장 참조를 위한 global 변수
    global dir

    # 헤더와 푸터..
    h_driver.get(f"https://blog.naver.com/PostPrint.naver?blogId={owner}&logNo={post_num}")
    pdf = b64decode(h_driver.print_page())
    with open(fr"{dir}/{date} {title} [작성자-{owner}][게시글 번호-{post_num}].pdf", "wb") as f:  # f설정과 r설정을 동시에 넣음
        f.write(pdf)


def find_post_list():
    global options, driver, page_bar, pages, owner, total_post_num, edge_service
    options = webdriver.EdgeOptions()
    options.add_argument('headless')
    #print('web_dir',web_dir)
    #driver = web_dir #  driver에 웹드라이버 경로 저장
    #driver = webdriver.Edge(driver, options=options) 웹드라이버 경로지정
    edge_service = Service(EdgeChromiumDriverManager().install())
    # 원래 위 Service parameter에 'msedgedriver' 설정해야되는데 기본값이 edge 라서 설정안함, 알아서 웹드라이버 설치하는 함수
    edge_service.creationflags = CREATE_NO_WINDOW
    driver = webdriver.Edge(options=options, service=edge_service)  # 웹드라이버 다운로드 없이 진행
    driver.get(f"https://blog.naver.com/PostList.naver?blogId={owner}&categoryNo=0&from=postList")
    total_post_num = driver.find_element(by=By.CLASS_NAME, value='category_title').text  # 전체 게시글 수를 구하고
    print(total_post_num, '토탈', re.search('전체보기 (.+?)개의 글', total_post_num).group(1))
    total_post_num = math.ceil(100/int(re.sub(',', '', re.search('전체보기 (.+?)개의 글', total_post_num).group(1))) * 100)/100
    print('나눈 토탈', total_post_num)
    # re.search로 숫자를 찾고, re.sub으로 쉼표를 없애주고 100을 전체 게시글 수로 나눠주고 셋째자리에서 올림
    #print('전체게시글수', total_post_num)
    list_open = driver.find_element(by=By.CLASS_NAME, value='btn_openlist')  # 목록열기 닫기 버튼 선택
    list_open_status = driver.find_element(by=By.ID, value='toplistSpanBlind') # 목록열기 버튼 텍스트 상태 확인
    if list_open_status.text == '목록열기':  # 목록열기로 떠있으면 목록을 열어준다
        list_open.send_keys('\n')
        print('its open ya')

    page_bar = driver.find_element(by=By.CLASS_NAME, value='blog2_paginate')  # 페이지번호 틀
    pages = page_bar.find_elements(by=By.CSS_SELECTOR, value='a')  # 각 페이지 숫자 버튼


def print_page_export():
    global options, driver, headless_driver, page_bar, pages, total_post_num, new_window, edge_service
    cnt = 0  # 페이지 넘김을 제어하기 위한 cnt 변수
    try:  # 잘못된 아이디를 입력했을 때를 대비하여 try except 문을 작성했음.
        find_post_list()  # 일단 게시글의 글 목록 부분의 DOM을 얻음.
    except:
        mb.showerror('오류', '없는 아이디입니다.\n아이디를 다시 입력해주세요.')
        return
    make_loading_window()  # 로딩 창을 띄운다.
    #return  테스트용
    while True:
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        li = soup.select('#listTopForm > table > tbody > tr > td.title > div > span')  # 뷰티풀숲으로 리스트상의 모든 제목들 선택
        li_date = soup.select('#listTopForm > table > tbody > tr > td.date > div > span')  # 리스트상의 모든 날짜들 선택
        headless_driver = r"C:\Users\jay\Desktop\edgedriver_win64\msedgedriver.exe"
        # 또 다른 웹드라이버를 만들어줘야지 초반에 게시글 번호 따려고 한 드라이버를 똑같이 써서 다시 get을 해버리면 dom이 날아가버림
        headless_driver = webdriver.Edge(options=options, service=edge_service)  #service=edge_service 다시 **********
        print(li, li_date,'안녕')

        for i in range(len(li)):
            link = li[i].find('a')['href'] # 게시글 링크 확인
            post_num = re.search('logNo=(.+?)&category', link).group(1)  # 정규표현식 라이브러리 사용 # 게시글 번호 확인
            title = li[i].text # 게시글 제목 확인
            title = re.sub('[\/:*?"<>|]','',title)  # 파일에 들어갈 수 없는 특수문자 제거
            date = li_date[i].text  # 게시글 작성 날짜 확인
            #print(date, post_num, title)
            print('테스트라뇽',link,post_num,title)
            make_pdf(headless_driver, post_num, date, title)
            pb['value'] += total_post_num  # Progress Bar 가산
            w.update_idletasks()  # Progress bar 업데이트
        #print()

        if len(pages) == 0:  # 페이지가 하나 뿐이라면 0인 이유는 페이지를 a 링크로 찾는데 하나 일 때는 a 링크 설정이 안걸린다.
            break
        if pages[-1].text != '다음' and cnt == len(pages):  # 조건문에서 pages를 건들기때문에 dom 갱신 순서 조심
            break
        if pages[cnt].text == '다음':
            pages[cnt].send_keys('\n')
            cnt = 1
        else:
            pages[cnt].send_keys('\n')
            cnt += 1
        time.sleep(2)
        page_bar = driver.find_element(by=By.CLASS_NAME, value='blog2_paginate')  # for 문 돌때마다 dom이 날라가므로 항상 다시 구해줘야됨
        pages = page_bar.find_elements(by=By.CSS_SELECTOR, value='a')
    new_window.destroy()
    mb.showinfo('Completed', '완료')
    driver.close()  # 다 끝났으니 웹드라이버를 종료해줌
    headless_driver.close()  # 두번째 웹드라이버도 종료해줘야지


def stop():
    global driver
    driver.close()  # 변환을 중단 했으니 웹드라이버도 종료해줘야지


def make_loading_window():
    global pb, new_window
    new_window = Toplevel(w)
    new_window.geometry('300x100')
    new_window.title('변환 중')
    pb = ttk.Progressbar(new_window, orient='horizontal', mode='determinate', length=150)
    pb.pack(pady=20)
    btn_stop = Button(new_window, text='중단', command=lambda: [new_window.destroy(), stop()])
    btn_stop.pack()


def press():
    global owner
    text = entry1.get()
    if text == '':
        mb.showerror('오류', '아이디를 입력해주세요.')

    #elif web_dir == '':
        #mb.showerror('오류', '웹드라이버 경로를 설정해주세요.')

    else:
        if mb.askyesno('확인', f'입력하신 아이디는 {text} 입니다. 변환하시려면 OK를 누르세요.'):
            owner = text
            th1 = Thread(target=print_page_export)  # 멀티쓰레딩 타겟 설정
            th1.setDaemon(True)  # 멀티쓰레딩 데몬 활성화
            th1.start()  # 멀티쓰레드 시작
            #print_page_export()
            return


def on_close():
    global driver, headless_driver
    if mb.askyesno('종료', '종료하시겠습니까?'):
        w.destroy()
        driver.close()  # 닫을 때 혹시모르니 살아있는 headless 프로세스를 종료하기 위해 클로즈를 해줌
        headless_driver.close()  # 두번째 웹드라이버도 종료해줘야지



def ask_dir():
    global dir
    dir = fd.askdirectory(parent=w, initialdir="C:/", title='저장 경로 설정')
    label1.config(text=dir + '/')
    print(dir)

'''
def ask_web_dir():
    global web_dir
    web_dir = fd.askopenfilename(parent=w, initialdir="C:/", title='웹드라이버 경로 설정')
    print(web_dir)
'''



#  tkinter GUI 만드는 부분
w.title("블로그 PDF변환")

# 창 크기
w.geometry('400x300')

# 창 배경화면
'''
w.configure(background="grey15")
label1 = Label(w, text='블로그 모든 게시글 PDF 변환', background='white')
label1.pack(pady=20)
label2 = Label(w, text='블로그 주인 아이디를 입력하세요.', background='white')
label2.pack(pady=5)
entry1 = Entry(w)
entry1.pack()
btn = Button(w, text='변환', command=press, padx=20, pady=20)
btn.pack(pady=20)
info = Label(w, text='ohito@naver.com')
info.pack()
btnSave = Button(w, text='저장 경로', command=ask_dir, padx=20, pady=5)
btnSave.place(x=230, y=131)
w.protocol('WM_DELETE_WINDOW', on_close)
w.mainloop()
'''
#  tkinter GUI 끝

w.title("블로그 PDF변환")
w.geometry("600x500")
w.configure(bg ="#FFFFFF")


canvas = Canvas(
    w,
    bg = "#FFFFFF",
    height = 500,
    width = 600,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge"
)

canvas.place(x = 0, y = 0)
canvas.create_rectangle(
    0.0,
    0.0,
    360.0,
    500.0,
    fill="#207BE5",
    outline="")

canvas.create_text(
    22.0,
    200.0,
    anchor="nw",
    text="사용 방법\n\n반드시 관리자 권한으로 실행하셔야 합니다.\n\n1. 블로그 주인 아이디 입력\n\n2. 저장 경로 지정\n\n(지정 안할 시 C드라이브에 자동 저장)\n\n4. 변환 버튼 클릭",
    fill="#FFFFFF",
    font=("MalgunGothic", 14 * -1)
)

canvas.create_text(
    20.0,
    40.0,
    anchor="nw",
    text="블로그 모든 게시글 PDF 변환 프로그램",
    fill="#FFFFFF",
    font=("MalgunGothic", 18 * -1)
)

canvas.create_rectangle(
    20.0,
    69.0,
    94.0,
    72.0,
    fill="#FFFFFF",
    outline="")

canvas.create_text(
    20.0,
    96.0,
    anchor="nw",
    text="이 프로그램은 블로그의 모든 게시글을\nPDF로 변환하여 지정된 경로에\n[작성날짜] [게시글 제목] 순으로\n저장해주는 프로그램입니다. ",
    fill="#FFFFFF",
    font=("MalgunGothic", 16 * -1)
)

canvas.create_text(
    20.0,
    385.0,
    anchor="nw",
    text="※ 블로그 주인 아이디는\n블로그 주소 blog.naver.com/\n뒤에 붙는 아이디를 확인하여 입력해주시기 바랍니다.",
    fill="#E6E6E6",
    font=("MalgunGothic", 14 * -1)
)

canvas.create_text(
    20.0,
    460.0,
    anchor="nw",
    text="제작자:oh jay ho (2022 당시 서울 강서구청 사회복무요원)\n문의: ohjayhoi@gmail.com",
    fill="#FFFFFF",
    font=("MalgunGothic", 12 * -1)
)


'''
button_image_3 = PhotoImage(
    file=relative_to_assets("button_3.png"))
btn_web_dir = Button(
    image=button_image_3,
    borderwidth=0,
    highlightthickness=0,
    command=ask_web_dir,
    relief="flat"
)
btn_web_dir.place(
    x=395.0,
    y=191.0,
    width=154.0,
    height=39.0
)
'''

canvas.create_text(
    385.0,
    59.0,
    anchor="nw",
    text="아이디를 입력해주세요",
    fill="#1D5CA6",
    font=("MalgunGothic", 18 * -1)
)

entry_image_1 = PhotoImage(
    file=relative_to_assets("entry_1.png"))
entry_bg_1 = canvas.create_image(
    472.0,
    128.5,
    image=entry_image_1
)
entry1 = Entry(
    bd=0,
    bg="#ECECEC",
    highlightthickness=0
)
entry1.place(
    x=391.0,
    y=130.0,
    width=160.0,
    height=20.0
)

canvas.create_text(
    390.0,
    108.0,
    anchor="nw",
    text="네이버 블로그 아이디",
    fill="#1D5CA7",
    font=("MalgunGothic", 12 * -1)
)

button_image_2 = PhotoImage(
    file=relative_to_assets("button_2.png"))
btn_dir = Button(
    image=button_image_2,
    borderwidth=0,
    highlightthickness=0,
    command=ask_dir,
    relief="flat"
)
btn_dir.place(
    x=395.0,
    y=200.0,
    width=154.0,
    height=39.0
)

canvas.create_text(
    374.0,
    280.0,
    anchor="nw",
    text="현재 경로: ",
    fill="#1D5CA7",
    font=("MalgunGothic", 15 * -1)
)

label1 = Label(w, text='C:/', background='white')
label1.place(x=374, y=305)

button_image_1 = PhotoImage(
    file=relative_to_assets("button_1.png"))

btn_convert = Button(
    image=button_image_1,
    borderwidth=0,
    highlightthickness=0,
    command=press,
    relief="flat"
)
btn_convert.place(
    x=374.0,
    y=390.0,
    width=197.0,
    height=46.0
)

w.resizable(False, False)
w.protocol('WM_DELETE_WINDOW', on_close)
w.mainloop()