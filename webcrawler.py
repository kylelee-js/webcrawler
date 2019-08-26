from bs4 import BeautifulSoup
# from urllib import parse
# import openpyxl
# from openpyxl import Workbook
import xlrd, xlwt
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

# 웹드라이버 불러오기
driver = webdriver.Chrome(executable_path=r'C:\\Users\\관산센터\\Downloads\\chromedriver_win32\\chromedriver.exe')

# 엑셀 시트 불러오기
wd = xlrd.open_workbook('아파트.xlsx')
names = wd.sheet_by_index(0)

# 실험을 위한 셀 하나만 불러오기
value = names.cell_value(7,16)


# 엑셀에 저장 위한 워크북 생성
wbwt = xlwt.Workbook(encoding='utf-8')
print('가동중...')

# 실제 포문

# max_rw = names.nrows
# max_cl = names.ncols

# for i in range(max_rw):
    # asd = names.cell_vlaue(i, 16)
    # address = bs(asd)  # bs()는 아래에 함수로 작성
    # ws = wbwt.add_sheet('Sheet2', cell_overwrite_ok=True)
    # max_rw2 = ws[1].nrows
    # for z in range(max_rw2):
        # ws.write(z,1,address)

# 마무리 저장  
# wbwt.save("2018년 공동주택(아파트) 현황.xlsx")




# 웹드라이버를 통한 크롬 브라우져 조작
def bs(address):
    driver.get('https://www.google.co.kr/maps/@37.9467592,126.6870721,5.08z?hl=ko')

    # 검색 버튼 클릭
    driver.find_element_by_name('q').send_keys(address)
    driver.find_element_by_name('q').send_keys(Keys.ENTER)    
    
    # 브라우져 로딩을 기다리는 타임 함수
    driver.implicitly_wait(10)
    time.sleep(5)

    print("1차 완료")
    
    html = driver.page_source
    bsObj = BeautifulSoup(html, "html.parser")

    span = bsObj.find_all('h2') #예시로 둔 태그, 실제 검색창 태그는 따로 있음
    for i in span:
        print(i)
    
    #asd = driver.find_element_by_css_selector('h1.section-hero-header-title')
    print("asdasd")
    print(span) #확인용
    
    return span[0].text
    

def zipcode():
    
    #address = wd.cell(row=6, column=16).value
    code = bs(value)
    # code = bs(asd)
    ws = wbwt.add_sheet('Sheet2', cell_overwrite_ok=True)
    ws.write(1,1,code)

    wbwt.save("2018년 공동주택(아파트) 현황.xlsx")
    print("작동이 완료!")
    

zipcode()
print('DONE!')
