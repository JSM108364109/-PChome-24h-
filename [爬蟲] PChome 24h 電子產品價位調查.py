# -*- coding: utf-8 -*-
"""
Created on Mon Feb 24 21:39:04 2020

@author: user
"""

import pandas as pd
import sqlite3
import os
from time import sleep
from selenium import webdriver
chromepath=r"C:\\Users\user\chromedriver.exe" #chromediver路徑
driver = webdriver.Chrome(executable_path=chromepath)  #執行chromedriver


################################################ 匯入欲查詢商品之Prod_Number
df=pd.read_excel(r'C:\Users\user\Desktop\Price_List.xlsx')
Prod_Number=df['Prod_Number'].tolist()
Name=[]
Price=[]
Price_str=[]

################################################ 抓取
for i in Prod_Number:
    print(i)
    URL="https://ecshweb.pchome.com.tw/search/v3.3/?q="+i
    driver.get(URL)
    driver.find_elements_by_class_name("Scope")[1].click() #選Scope裡的PChome24H到貨分類
    sleep(0.5)

    try:
        tem_pn = WebDriverWait(driver, 5).until( EC.presence_of_element_located((By.CLASS_NAME, "prod_name")) ) #抓商品名稱
        tem_price = WebDriverWait(driver, 5).until( EC.presence_of_element_located((By.CLASS_NAME, "price")) ) #抓價格
     
    except:
        print(i+"找不到商品")
        continue

    Name.append(tem_pn.text) #name結果加入創建的list
    Price_str.append(tem_price.text) #Price_str結果加入創建的list
    
################################################ 整理&匯出 
for x in Price_str:      
    x=x.split('$') ###字串切割
    x=int(x[1]) ###字串轉數值
    Price.append(x)

result_dict={"Name":Name,"Price":Price}
result=pd.DataFrame(result_dict)
result.to_excel("./20200303_result.xlsx",index=False, header=True)

