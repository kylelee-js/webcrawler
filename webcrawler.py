from bs4 import BeautifulSoup
from urllib import parse
import openpyxl
from openpyxl import Workbook
import xlrd, xlwt
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time


driver = webdriver.Chrome(executable_path=r'C:\\Users\\관산센터\\Downloads\\chromedriver_win32\\chromedriver.exe')

wd = xlrd.open_workbook('아파트.xlsx')
names = wd.sheet_by_index(0)
value = names.cell_value(7,16)

wbwt = xlwt.Workbook(encoding='utf-8')
print('가동중...')


def bs(address):
    driver.get('https://www.google.co.kr/maps/@37.9467592,126.6870721,5.08z?hl=ko')

    driver.find_element_by_name('q').send_keys(address)
    driver.find_element_by_name('q').send_keys(Keys.ENTER)
    
    
    driver.implicitly_wait(10)
    time.sleep(5)

    print("1차 완료")
    
    html = driver.page_source
    bsObj = BeautifulSoup(html, "html.parser")

    span = bsObj.find_all('h2')
    for i in span:
        print(i)
    
    #asd = driver.find_element_by_css_selector('h1.section-hero-header-title')
    print("asdasd")
    print(span) #확인용
    
    return span[0].text
    

def zipcode():
    
    #address = wd.cell(row=6, column=16).value
    code = bs(value)
    ws = wbwt.add_sheet('Sheet2', cell_overwrite_ok=True)
    ws.write(1,1,code)

    wbwt.save("2018년 공동주택(아파트) 현황.xlsx")
    print("작동이 완료!")
    

zipcode()
print('DONE!')
