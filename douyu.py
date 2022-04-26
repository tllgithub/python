
# -*- coding: utf-8 -*-
import xlrd #导入xlrd模块
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import math
import sys

reload(sys)
sys.setdefaultencoding('utf8')
bok = xlrd.open_workbook(r'C:/Users/xgm/Desktop/MCC-0329.xlsx') #excel表格路径

sht = bok.sheets()[0]

row =sht.nrows
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
driver = webdriver.Chrome(options=chrome_options)
url = "https://www.wjx.cn/vj/rXCuxtS.aspx"
driver.get(url)
for i in range(row):
    rowdate = sht.row_values(i)#i行的list
    if i==293:
        for a, b in enumerate(rowdate):
            x = str(b)
            y = "#divquestion"+str(a+1)+" > ul > li:nth-child("+x[0]+")"
            time.sleep(1.5)
            if a<3:
                print(y)
                driver.find_element_by_css_selector(y).click()
            if a>=3:
                if a<=4:
                    driver.find_element_by_id("select2-q"+str(a+1)+"-container").click()
                    time.sleep(1.5)
                    driver.find_element_by_css_selector("#select2-q"+str(a+1)+"-results > li:nth-child("+str(b+1)[0]+")").click()
            if a>4:
                if a<7:
                    driver.find_element_by_css_selector(y).click()
            if a>6:
                max = 3 if (a-6)%3 ==0 else (a-6)%3
                z = "#divquestion"+str(int(math.ceil(float(a-6)/3)+7))+" > table > tbody > tr:nth-child("+str(max)+") > td:nth-child("+str(b+1)[0]+")"
                print(z)
                driver.find_element_by_css_selector(z).click()
        driver.find_element_by_id("submit_button").click()







