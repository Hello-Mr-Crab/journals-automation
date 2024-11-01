##############################################################################
'''依赖环境'''
import re
import os
import fitz
import base64
import time
import shutil
import random
import zipfile
import ddddocr
import requests
import pyautogui
import threading
from pandas import DataFrame



from time import sleep
from time import strftime
from datetime import datetime
from collections import OrderedDict
from bs4 import BeautifulSoup
#所有图片
from pictures.Backgrounds import 主界面背景
from pictures.Backgrounds import 微信名片
from pictures.Icons import IEEE_icon  
from pictures.Icons import Google_icon
from pictures.Icons import 中国知网_icon
from pictures.Icons import 主界面_icon
from pictures.Icons import 注意_icon
from pictures.Icons import 中国科学_icon
from pictures.Icons import LightScience_icon
from pictures.Icons import Springer_icon
from pictures.Icons import 电子学报_icon
from pictures.Icons import 通信学报_icon
from pictures.Icons import 电子测量与仪器学报_icon
from pictures.Icons import 项目原码_icon
from pictures.Icons import PDF_icon
#selenium环境
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
#matplotlib环境
import catppuccin
import matplotlib.pyplot as plt
import matplotlib as mpl
mpl.style.use(catppuccin.PALETTE.mocha.identifier)
#tkinter环境
import tkinter as tk
from tkinter import ttk
from tkinter import BOTH
from tkinter import messagebox,filedialog
#utils中的一些函数
from utils.account import account_and_pasword
from utils.Systemsettings import set_Volume_to_100
from utils.Systemsettings import open_Listening_mode
from utils.Systemsettings import close_Listening_mode
from utils.Tools import Convert_digitstring
from utils.Tools import Base64_to_Image
from utils.Tools import Speak
from utils.Tools import Drawsystem
from utils.Tools import get_windows_scaling_factor
plt.rcParams['font.sans-serif']=['SimHei']
IEEExplore_url='https://ieeexplore.ieee.org/Xplore/home.jsp'#IEEEexplore的网址
Googleschloar_url='https://scholar.lanfanshu.cn/'
SUES='https://www.sues.edu.cn/'#学校官网
Cnki='https://www.cnki.net/' #知网网址   


##############################################################################
'''自动化搜索结果统计'''
def search_cnki(search_element,flag): 
    open_Listening_mode()
    journal_result=[0]*len(search_element)
    essay_result=[0]*len(search_element)
    options=Options()
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument('--ignore-ssl-errosr')
    options.add_argument('--ignore-certificate-errors')
    options.add_experimental_option('excludeSwitches',['enable-automation'])
    if flag:
        options.add_argument("-headless=new")
        options.add_argument("--disable-gpu")
        options.add_argument('--window-position=-2400,-2400')
    
    browser=webdriver.ChromiumEdge(options)
    browser.get(Cnki)
    browser.maximize_window()
    input_element=browser.find_element(By.ID,'txt_SearchText')
    input_element.send_keys(search_element[0])
    win=browser.window_handles
    browser.switch_to.window(win[-1])
    sleep(2)
    try:
        none=browser.find_element(By.XPATH,'//*[@id="briefBox"]/p')
        if none:
            journal_result[0]="未查询到内容"
            essay_result[0]="未查询到内容"
    except NoSuchElementException:
        学术期刊=browser.find_element(By.XPATH,'//*[@id="ModuleSearch"]/div[2]/div/div/ul/li[1]/a/em')
        学位论文=browser.find_element(By.XPATH,'//*[@id="ModuleSearch"]/div[2]/div/div/ul/li[2]/a/em')
        sleep(4)
        journal_result[0]=学术期刊.text
        essay_result[0]=学位论文.text
        input1=browser.find_element(By.ID,'txt_search')
        input1.clear()
   # 在第一个查找元素查找后的界面内循环搜索
    for i in range(1,len(search_element)):
        input1.send_keys(search_element[i])
        input1.send_keys(Keys.ENTER)
        win=browser.window_handles
        browser.switch_to.window(win[-1])
        try:
            none=browser.find_element(By.XPATH,'//*[@id="briefBox"]/p')
            if none:
                journal_result[i]="未查询到内容"
                essay_result[i]="未查询到内容"
            input1.clear()
        except NoSuchElementException:
            学术期刊=browser.find_element(By.XPATH,'//*[@id="ModuleSearch"]/div[2]/div/div/ul/li[1]/a/em')
            学位论文=browser.find_element(By.XPATH,'//*[@id="ModuleSearch"]/div[2]/div/div/ul/li[2]/a/em')
            sleep(4)
            journal_result[i]=学术期刊.text
            essay_result[i]=学位论文.text
            input1.clear()
    browser.quit()
    final_result=DataFrame([journal_result,essay_result],index=['学术期刊数量','学位论文数量'],columns=search_element)
    final_result=final_result.transpose()
    close_Listening_mode()
    return final_result

     
def search_IEEE(search_element):   
    open_Listening_mode()
    journal_result=[0]*len(search_element)
    total_result=[0]*len(search_element)
    conferences_result=[0]*len(search_element)
    options=Options()
    options.add_argument('--ignore-ssl-errosr')
    options.add_argument('--ignore-certificate-errors')
    options.add_experimental_option('excludeSwitches',['enable-automation'])
    options.add_argument('--disable-blink-features=AutomationControlled')
    browser=webdriver.ChromiumEdge(options)
    browser.maximize_window() 
    try:
       browser.get(IEEExplore_url)
       WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/div[2]/div[2]/xpl-search-bar-migr/div/form/div[2]/div/div[1]/xpl-typeahead-migr/div/input')))
       input_element=browser.find_element(By.XPATH, '//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/div[2]/div[2]/xpl-search-bar-migr/div/form/div[2]/div/div[1]/xpl-typeahead-migr/div/input')
       input_element.send_keys(search_element[0])
       input_element.send_keys(Keys.ENTER)
       win=browser.window_handles
       browser.switch_to.window(win[-1])
       sleep(3)
       try:
            total=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/h1/span[1]/span[2]')
            if total.text:
                total_result[0]=total.text
                try:
                    journals=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/div[1]/xpl-facet-content-type-migr/div/div/div[2]/label')
                    if journals.text:
                        journal_result[0]=journals.text[10:len(journals.text)-1]
                    try:
                        conferences=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/div[1]/xpl-facet-content-type-migr/div/div/div[1]/label')
                        if conferences.text:
                            conferences_result[0]=conferences.text[13:len(conferences.text)-1]
                    except NoSuchElementException:
                        conferences_result[0]='未查询到内容'   
                except NoSuchElementException:
                    journal_result[0]='未查询到内容'
                    try:
                        conferences=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/div[1]/xpl-facet-content-type-migr/div/div/div[1]/label')
                        if conferences.text:
                            conferences_result[0]=conferences.text[13:len(conferences.text)-1]
                    except NoSuchElementException:
                        conferences_result[0]='未查询到内容'
       except NoSuchElementException:
            total_result[0]='未查询到内容'
            journal_result[0]="未查询到内容"
            conferences_result[0]="未查询到内容"
       input1=browser.find_element(By.XPATH,'//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/div/div/xpl-search-bar-migr/div/form/div[2]/div/div[1]/xpl-typeahead-migr/div/input')
       # 在第一个查找元素查找后的界面内循环搜索
       for i in range(1,len(search_element)):
            input1=browser.find_element(By.XPATH,'//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/div/div/xpl-search-bar-migr/div/form/div[2]/div/div[1]/xpl-typeahead-migr/div/input')
            input1.send_keys(search_element[i])
            input1.send_keys(Keys.ENTER)
            win=browser.window_handles
            browser.switch_to.window(win[-1])
            browser.implicitly_wait(10)
            try:
                total=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/h1/span[1]/span[2]')
                if total.text:
                    total_result[i]=total.text
                    try:
                        journals=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/div[1]/xpl-facet-content-type-migr/div/div/div[2]/label')
                        if journals.text:
                            journal_result[i]=journals.text[10:len(journals.text)-1]
                        try:
                            conferences=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/div[1]/xpl-facet-content-type-migr/div/div/div[1]/label')
                            if conferences.text:
                                conferences_result[i]=conferences.text[13:len(conferences.text)-1]
                        except NoSuchElementException:
                            conferences_result[i]='未查询到内容'
                    except NoSuchElementException:
                        journal_result[i]='未查询到内容'
                        try:
                            conferences=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/div[1]/xpl-facet-content-type-migr/div/div/div[1]/label')
                            if conferences.text:
                                conferences_result[i]=conferences.text[13:len(conferences.text)-1]
                        except NoSuchElementException:
                            conferences_result[i]='未查询到内容'
            except NoSuchElementException:
                total_result[i]='未查询到内容'
                journal_result[i]="未查询到内容"
                conferences_result[i]="未查询到内容"
       browser.quit()
       final_result=DataFrame([total_result,journal_result,conferences_result],index=["Total",'Journals','Conferences'],columns=search_element)
       final_result=final_result.transpose()
       close_Listening_mode()
       return final_result
    except TimeoutException:
          browser.get(SUES)
          browser.maximize_window()
          actions=ActionChains(browser)
          #进入学校官网的图书馆
          library=browser.find_element(By.XPATH,'/html/body/div[5]/div/div[2]/div[2]/div/div/ul/li[6]/a').click()
          win=browser.window_handles
          browser.switch_to.window(win[-1])#切换到搜索后当前的页面
          #在图书馆内查找数字资源
          resource=browser.find_element(By.XPATH,'/html/body/div[3]/div/div[2]/div[3]/div/div/div/div[2]/div[1]/div[1]')
          actions.move_to_element(resource).perform()
          #在数字资源处找到外文数据库
          外文数据库=browser.find_element(By.LINK_TEXT,"外文数据库")
          外文数据库.click()#点击外文数据库
          browser.implicitly_wait(4)
          browser.find_element(By.XPATH,'//*[@id="wp_news_w9"]/ul/li[8]/span/a').click()
          win=browser.window_handles
          browser.switch_to.window(win[-1])#切换到搜索后当前的页面
          sleep(2)
          input_element=browser.find_element(By.XPATH, '//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/div[2]/div[2]/xpl-search-bar-migr/div/form/div[2]/div/div[1]/xpl-typeahead-migr/div/input')
          input_element.send_keys(search_element[0])
          input_element.send_keys(Keys.ENTER)
          win=browser.window_handles
          browser.switch_to.window(win[-1])
          sleep(3)
          try:
                total=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/h1/span[1]/span[2]')
                if total.text:
                    total_result[0]=total.text
                    try:
                        journals=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/div[1]/xpl-facet-content-type-migr/div/div/div[2]/label')
                        if journals.text:
                            journal_result[0]=journals.text[10:len(journals.text)-1]
                        try:
                            conferences=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/div[1]/xpl-facet-content-type-migr/div/div/div[1]/label')
                            if conferences.text:
                                conferences_result[0]=conferences.text[13:len(conferences.text)-1]
                        except NoSuchElementException:
                            conferences_result[0]='未查询到内容'   
                    except NoSuchElementException:
                        journal_result[0]='未查询到内容'
                        try:
                            conferences=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/div[1]/xpl-facet-content-type-migr/div/div/div[1]/label')
                            if conferences.text:
                                conferences_result[0]=conferences.text[13:len(conferences.text)-1]
                        except NoSuchElementException:
                            conferences_result[0]='未查询到内容'
          except NoSuchElementException:
                total_result[0]='未查询到内容'
                journal_result[0]="未查询到内容"
                conferences_result[0]="未查询到内容"
          input1=browser.find_element(By.XPATH,'//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/div/div/xpl-search-bar-migr/div/form/div[2]/div/div[1]/xpl-typeahead-migr/div/input')
           # 在第一个查找元素查找后的界面内循环搜索
          for i in range(1,len(search_element)):
                input1=browser.find_element(By.XPATH,'//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/div/div/xpl-search-bar-migr/div/form/div[2]/div/div[1]/xpl-typeahead-migr/div/input')
                input1.send_keys(search_element[i])
                input1.send_keys(Keys.ENTER)
                win=browser.window_handles
                browser.switch_to.window(win[-1])
                browser.implicitly_wait(10)
                try:
                    total=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/h1/span[1]/span[2]')
                    if total.text:
                        total_result[i]=total.text
                        try:
                            journals=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/div[1]/xpl-facet-content-type-migr/div/div/div[2]/label')
                            if journals.text:
                                journal_result[i]=journals.text[10:len(journals.text)-1]
                            try:
                                conferences=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/div[1]/xpl-facet-content-type-migr/div/div/div[1]/label')
                                if conferences.text:
                                    conferences_result[i]=conferences.text[13:len(conferences.text)-1]
                            except NoSuchElementException:
                                conferences_result[i]='未查询到内容'
                        except NoSuchElementException:
                            journal_result[i]='未查询到内容'
                            try:
                                conferences=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[1]/div[2]/xpl-search-dashboard/section/div/div[1]/xpl-facet-content-type-migr/div/div/div[1]/label')
                                if conferences.text:
                                    conferences_result[i]=conferences.text[13:len(conferences.text)-1]
                            except NoSuchElementException:
                                conferences_result[i]='未查询到内容'
                except NoSuchElementException:
                    total_result[i]='未查询到内容'
                    journal_result[i]="未查询到内容"
                    conferences_result[i]="未查询到内容"
          browser.quit()
          final_result=DataFrame([total_result,journal_result,conferences_result],index=["Total",'Journals','conferences'],columns=search_element)
          final_result=final_result.transpose()
          close_Listening_mode()
          return final_result     

def search_google(search_element,flag=0):
    open_Listening_mode()
    #打开edge，打开google学术官网
    final_result=[0]*len(search_element)
    options=Options()
    options.add_argument("--disable-blink-features=AutomationControlled")#欺骗网页，隐藏自动化操作的表头
    options.add_argument('--ignore-ssl-errosr')
    options.add_argument('--ignore-certificate-errors')
    options.add_experimental_option('excludeSwitches',['enable-automation'])
    if flag:
        options.add_argument("-headless=new")
        options.add_argument("--disable-gpu")
        options.add_argument('--window-position=-2400,-2400')
    else:
           pass
    browser=webdriver.ChromiumEdge(options)
    browser.get(Googleschloar_url)
    browser.maximize_window()
    #在输入框输入传入的search_element
    input_element=browser.find_element(By.ID,'gs_hdr_tsi')
    input_element.send_keys(search_element[0])
    input_element.send_keys(Keys.ENTER)#模拟点击enter健
    try :
        #等待1秒，看是否有广告出现，如果没有广告出现,webdriverwait会抛出Timeoutexception
        WebDriverWait(browser, 2).until(EC.presence_of_element_located((By.LINK_TEXT,"广告")))
        #关闭广告元素
        browser.switch_to.alert.dismiss()
        #关闭后继续查找
        win=browser.window_handles
        browser.switch_to.window(win[-1])#切换到关闭广告后当前的页面
        sleep(1)
       
        element = browser.find_element(By.XPATH, '//*[@id="gs_ab_md"]/div')#定位搜索结果和时间戳所在页面
        final_result[0]=element.text
    except TimeoutException:#没有广告时接受到Timeoutexception，直接查找
        win=browser.window_handles
        browser.switch_to.window(win[-1])#切换到关闭广告后当前的页面
        sleep(2)
        element = browser.find_element(By.XPATH, '//*[@id="gs_ab_md"]/div')#定位搜索结果和时间戳所在页面
        final_result[0]=element.text
        input_element1=browser.find_element(By.ID,'gs_hdr_tsi')
        input_element1.clear()
    for i in range(1,len(search_element)):
        input_element1.send_keys(search_element[i])
        input_element1.send_keys(Keys.ENTER)#模拟点击enter健
        try :
            #等待2秒，看是否有广告出现，如果没有广告出现,webdriverwait会抛出Timeoutexception
            WebDriverWait(browser, 2).until(EC.presence_of_element_located((By.LINK_TEXT,"广告")))
            #关闭广告元素
            browser.switch_to.alert.dismiss()
            #关闭后继续查找
            win=browser.window_handles
            browser.switch_to.window(win[-1])#切换到关闭广告后当前的页面
            sleep(1)

            element = browser.find_element(By.XPATH, '//*[@id="gs_ab_md"]/div')#定位搜索结果和时间戳所在页面
            final_result[i]=element.text
            browser.refresh()
            input_element1=browser.find_element(By.ID,'gs_hdr_tsi')
            input_element1.clear()
        except TimeoutException:#没有广告时接受到Timeoutexception，直接查找
            win=browser.window_handles
            browser.switch_to.window(win[-1])#切换到关闭广告后当前的页面
            sleep(2)
            element = browser.find_element(By.XPATH, '//*[@id="gs_ab_md"]/div')#定位搜索结果和时间戳所在页面
            final_result[i]=element.text
            input_element1=browser.find_element(By.ID,'gs_hdr_tsi')
            input_element1.clear()
    browser.quit()
    final_result=DataFrame([final_result],columns=search_element,index=["搜索结果"])
    final_result=final_result.transpose()
    close_Listening_mode()
    return final_result

def get_time():
    year=strftime("%Y")
    month=strftime("%m")
    week=int(strftime("%w"))
    day=strftime("%d")
    hour=strftime("%H")
    minute=strftime("%M")
    second=strftime("%S") 
    week_dic={1:'一',2:'二',3:'三',4:'四',5:'五',6:'六',0:'日'}
    time_now=f'{year}年 {month}月{day}日 周{week_dic.get(week)} {hour}:{minute}:{second}'
    clock=tk.Label(main_canvas,text=time_now,font=15,fg="#000000",bg="#F3C5FF")
    clock.grid(row=1,column=1,pady=20)
    if root.winfo_exists():
        root.after(1000,get_time)
    
def IEEE():
    Iroot=tk.Toplevel()
    Iroot.geometry("1410x1100")
    Iroot.resizable(True,True)
    Iroot.title("IEEE自动化搜索结果统计")
 
    Iroot.iconphoto(True,Base64_to_Image(IEEE_icon.img,100,100))
    frame=tk.Frame(Iroot,bg='#d7e8f0')
    frame.pack(fill="both",expand=True)
    tree=ttk.Treeview(frame,columns=("第一列","第二列","第三列","第四列"),height=20,show="headings")
    tree.heading("第一列",text="待搜索内容")
    tree.heading("第二列",text="Total")
    tree.heading("第三列",text="Journals")
    tree.heading("第四列",text="Conferences")
    tree.column("第一列",width=500)
    tree.column("第二列",width=200)
    tree.column("第三列",width=200)
    tree.column("第四列",width=200)
    tree.grid(row=3,column=3,sticky="nsew")
    scrollbar=ttk.Scrollbar(frame,orient="vertical",command=tree.yview)
    scrollbar.grid(row=3,column=4,sticky="ns")
    tree.configure(yscrollcommand=scrollbar.set)
    def file():
        search_list=[]
        file_path=filedialog.askopenfilename(title='选择存放关键词的txt文件')
        if  os.path.exists(file_path) and os.path.isfile(file_path):
            with open(file_path, 'r', encoding='utf-8') as infile:
                lines=infile.readlines()
                for line in lines:
                    search_list.append(line)
            cleaned_lines=[line.replace(" ","") for line in lines]
            cleaned_search_list_IEEE=[line.replace("\n","") for line in search_list]
            for line in cleaned_lines:
                   tree.insert("","end",values=(line))
            def searchmore():
                def auto_export_result():
                    root1.destroy()
                    button1.config(state=tk.DISABLED)
                    messagebox.showinfo("保存",f"请选择脚本结束后自动保存的位置\n  文件类型:xlsx")
                    file_path=filedialog.asksaveasfilename(defaultextension=".xlsx",title="选择保存位置")
                    if not os.path.exists(file_path) and os.path.isfile(file_path):
                            messagebox.showinfo("警告","请输入有效的地址！")
                    final_result=search_IEEE(cleaned_search_list_IEEE)
                    for item,value in zip(tree.get_children(),final_result['Total']):
                        tree.set(item,"第二列",value)
                    for item,value in zip(tree.get_children(),final_result['Journals']):
                        tree.set(item,"第三列",value)
                    for item,value in zip(tree.get_children(),final_result['Conferences']):
                        tree.set(item,"第四列",value)
                    if final_result.shape[0]==len(cleaned_search_list_IEEE):
                        final_result.to_excel(file_path,index=True)
                        label=tk.Label(Iroot,text='搜索结果可视化',font=('宋体',20))
                        list1=list(final_result['Total'])
                        list2=list(final_result['Journals'])
                        list3=list(final_result['Conferences'])
                        list1=[int(number.replace(",","")) for number in list1]
                        list2=[int(number.replace(",","")) for number in list2]
                        list3=[int(number.replace(",","")) for number in list3]
                   
                        dict1=dict(zip(final_result.index.tolist(),list1))
                        dict2=dict(zip(final_result.index.tolist(),list2))
                        dict3=dict(zip(final_result.index.tolist(),list3))
                        sorted_items1=sorted(dict1.items(),key=lambda item:item[1],reverse=True)
                        sorted_items2=sorted(dict2.items(),key=lambda item:item[1],reverse=True)
              
                        sorted_items3=sorted(dict3.items(),key=lambda item:item[1],reverse=True)
                        ordereddict1=OrderedDict(sorted_items1)
                        ordereddict2=OrderedDict(sorted_items2)
                        ordereddict3=OrderedDict(sorted_items3)
                        
                     
                        fig,axs=plt.subplots(nrows=1,ncols=3,figsize=(20,10))
                        axs[0].bar(x=list(ordereddict1.keys())[0:5],height=list(ordereddict1.values())[0:5],width=0.6,color='#00C9A7')
                        
                        axs[1].bar(x=list(ordereddict2.keys())[0:5],height=list(ordereddict2.values())[0:5],width=0.6,color='#B39CD0') 
                       
                        axs[0].set_xticklabels(list(ordereddict1.keys())[0:5],rotation=25)
                        axs[0].set_xlabel('关键字')
                        axs[0].set_title('Total')
                        axs[0].set_ylabel('关键字数量',rotation=35)
                       
                        axs[1].set_title('Journals')
                        axs[1].set_xlabel('关键字')
                        axs[1].set_ylabel('关键字数量',rotation=35)
                        axs[1].set_xticklabels(list(ordereddict2.keys())[0:5],rotation=25)
                       
                        axs[2].bar(x=list(ordereddict3.keys())[0:5],height=list(ordereddict3.values())[0:5],width=0.6,color='#B39CD0')
                        
                        axs[2].set_title('Conferences')
                        axs[2].set_xticklabels(list(ordereddict3.keys())[0:5],rotation=25)
                        axs[2].set_xlabel('关键字')
                        axs[2].set_ylabel('关键字数量',rotation=35)
                        
                        fig.subplots_adjust(wspace=0.8)
                        for i,value in enumerate(list(ordereddict1.values())[0:5]):
                            axs[0].text(i,value+5,str(value),ha='center',va='bottom')
                        for i,value in enumerate(list(ordereddict2.values())[0:5]):
                            axs[1].text(i,value+5,str(value),ha='center',va='bottom')
                        for i,value in enumerate(list(ordereddict1.values())[0:5]):
                            axs[2].text(i,value+3,str(value),ha='center',va='bottom')
                     
                        Drawsystem(Iroot,fig)
                        plt.suptitle('关键字搜索结果数量Top5')
                        plt.show()
                        messagebox.showinfo("提示",f"已为您导出结果至{file_path}")
                       
                def manual_export_result():
                    def export_result():
                        file_path=filedialog.asksaveasfilename(defaultextension=".xlsx",title="选择保存位置")
                        if  os.path.exists(file_path) and os.path.isfile(file_path):
                            final_result.to_excel(file_path,index=True)
                            messagebox.showinfo("提示",f"已为您导出结果至{file_path}")
                        else:
                            messagebox.showinfo("警告","请输入有效的地址！")
                    root1.destroy()
                    final_result=search_IEEE(cleaned_search_list_IEEE)
                    for item,value in zip(tree.get_children(),final_result['Total']):
                        tree.set(item,"第二列",value)
                    for item,value in zip(tree.get_children(),final_result['Journals']):
                        tree.set(item,"第三列",value)
                    for item,value in zip(tree.get_children(),final_result['Conferences']):
                        tree.set(item,"第四列",value)
                    if len(final_result)==len(cleaned_search_list_IEEE):
                      button1.config(state=tk.DISABLED)
                      button3=tk.Button(Iroot,text="一键手动导出结果",command=export_result,width=21,height=3,relief="raised",fg="#000000",bg="#0087FF")
                      button3.place(x=1150,y=300)   
                root1=tk.Tk()
                root1.geometry("400x300")
                root1.title('选择保存方式')
                button4=tk.Button(root1,text="脚本结束后自动保存",command=auto_export_result,width=21,height=3,relief="raised",fg="#000000",bg="#0087FF")
                button4.place(x=80,y=40)
                button5=tk.Button(root1,text="脚本结束后手动保存",command=manual_export_result,width=21,height=3,relief="raised",fg="#000000",bg="#0087FF")
                button5.place(x=80,y=200)
                root1.mainloop()

            button1=tk.Button(Iroot,text="运行脚本",command=searchmore,width=21,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button1.place(x=1150,y=150)
    def helper():
        messagebox.showinfo("说明","该GUI只适用于自动化操作IEEEEXPLORE与GOOGLE学术以及中国知网\n的搜索过程并统计搜索结果，不适合其他搜索路径。\n点击运行脚本后请勿操纵鼠标，请耐心等待自动化操作完成。\n注意：导出数据时请提前建好空白xlsx文件然后选中，否则会\n覆盖掉已有txt原先内容，选中时询问替换，点击确定即可。\n                                        感谢您的使用！")
    def abouts():
        messagebox.showinfo("版本信息","开发者：Mr Crab\n版本号：8.0")
    menubar=tk.Menu(Iroot)
    Iroot.config(menu=menubar)
    menubar.add_command(label="导入存放搜索关键字的txt文件",font=("宋体",15),command=file)
    menubar.add_separator()
    menubar.add_command(label="说明",font=("宋体",15),command=helper)
    menubar.add_separator()
    menubar.add_command(label="关于",font=("宋体",15),command=abouts)
    Iroot.mainloop()
def Google():
    Groot=tk.Toplevel()
    Groot.geometry("1100x600")
    Groot.resizable(False,False)
 
    Groot.title("Google学术自动化搜索结果统计")
    Groot.iconphoto(True,Base64_to_Image(Google_icon.img,100,100))
    Groot.configure(bg="#C4FCEF")
    frame=tk.Frame(Groot,bg='#d7e8f0')
    frame.pack(fill="both",expand=True)
    tree=ttk.Treeview(frame,columns=("第一列","第二列"),height=20,show="headings")
    tree.heading("第一列",text="待搜索内容")
    tree.heading("第二列",text="搜索结果")
    tree.column("第一列",width=400)
    tree.column("第二列",width=400)
    tree.grid(row=3,column=3,sticky="nsew")
    scrollbar=ttk.Scrollbar(frame,orient="vertical",command=tree.yview)
    scrollbar.grid(row=3,column=4,sticky="ns")
    tree.configure(yscrollcommand=scrollbar.set)
    def file():
        search_list=[]
        file_path=filedialog.askopenfilename(title='选择存放关键字的txt文件')
        if  os.path.exists(file_path) and os.path.isfile(file_path):
            with open(file_path, 'r', encoding='utf-8') as infile:
                lines=infile.readlines()
                for line in lines:
                    search_list.append(line)
            cleaned_lines=[line.replace(" ","") for line in lines]
            cleaned_search_list=[line.replace("\n","") for line in search_list]
            for line in cleaned_lines:
                   tree.insert("","end",values=(line))
            final_result=[]
            def searchmore(flag):
                def auto_export_result():
                    root1.destroy()
                    button1.config(state="disabled")
                    messagebox.showinfo("保存","请选择脚本结束后自动保存的位置")
                    file_path=filedialog.asksaveasfilename(defaultextension=".xlsx",title=f"选择保存位置\n  文件类型：xlsx")
                    if not os.path.exists(file_path) and os.path.isfile(file_path):
                            messagebox.showinfo("警告","请输入有效的地址！")
                    final_result=search_google(cleaned_search_list,flag)
                    for item,value in zip(tree.get_children(),final_result['搜索结果']):
                        tree.set(item,"第二列",value)
                    if final_result.shape[0]==len(cleaned_search_list):
                                final_result.to_excel(file_path,index=True)
                                messagebox.showinfo("提示",f"已为您导出结果至{file_path}")
                def manual_export_result():
                    def export_result():
                        file_path=filedialog.asksaveasfilename(defaultextension=".xlsx",title="选择保存位置(文件类型：xlsx)")
                        if  os.path.exists(file_path) and os.path.isfile(file_path):
                            final_result.to_excel(file_path,index=True)
                            messagebox.showinfo("提示",f"已为您导出结果至{file_path}")
                        else:
                            messagebox.showinfo("警告","请输入有效的地址！")
                    root1.destroy()
                    final_result=search_google(cleaned_search_list,flag)
                    for item,value in zip(tree.get_children(),final_result['搜索结果']):
                        tree.set(item,"第二列",value)
                    if final_result.shape[0]==len(cleaned_search_list):
                      button1.config(state=tk.DISABLED)
                      button3=tk.Button(Groot,text="一键手动导出结果",command=export_result,width=21,height=3,relief="raised",fg="#000000",bg="#0087FF")
                      button3.place(x=590,y=500)   
                root1=tk.Tk()
                root1.geometry("400x300")
                root1.title('选择保存方式')
                button4=tk.Button(root1,text="脚本结束后自动保存",command=auto_export_result,width=21,height=3,relief="raised",fg="#000000",bg="#0087FF")
                button4.place(x=80,y=40)
                button5=tk.Button(root1,text="脚本结束后手动保存",command=manual_export_result,width=21,height=3,relief="raised",fg="#000000",bg="#0087FF")
                button5.place(x=80,y=200)
                root1.mainloop()
            
            label1=tk.Label(Groot,text='选择浏览器模式',font=('宋体',10),fg='red',bg='#d7e8f0')
            label1.place(x=850,y=20)
            combo1=ttk.Combobox(Groot,validate='none')
            combo1['values']=('打开浏览器','不打开浏览器')
            combo1.current(0)
            combo1.place(x=850,y=40)
            def set_flag(option):
                dic={'打开浏览器':0,'不打开浏览器':1}
                global flag
                flag=dic.get(option)
                return flag
          
            sure_button1=tk.Button(Groot,text="确认",command=lambda:set_flag(combo1.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
            sure_button1.place(x=950-scaling_factor+65 if scaling_factor!=150 else 950,y=100)
            button1=tk.Button(Groot,text="立即运行脚本",command=lambda:searchmore(flag),width=21,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button1.place(x=190,y=500)
         
            
  
    def helper():
        messagebox.showinfo("说明","该GUI只适用于自动化操作IEEEEXPLORE与GOOGLE学术以及中国知网\n的搜索过程并统计搜索结果，不适合其他搜索路径。\n点击运行脚本后请勿操纵鼠标，请耐心等待自动化操作完成。\n注意：导出数据时请提前建好空白xlsx文件然后选中，否则会\n覆盖掉已有txt原先内容，选中时询问替换，点击确定即可。\n                                        感谢您的使用！")
    def abouts():
        messagebox.showinfo("版本信息","开发者：Mr Crab\n版本号：8.0")    
    menubar=tk.Menu(Groot,tearoff=True)
    Groot.config(menu=menubar)
    menubar.add_command(label="导入存放搜索关键字的txt文件",font=("宋体",15),command=file)
    menubar.add_separator()
    menubar.add_command(label="说明",font=("宋体",15),command=helper)
    menubar.add_separator()
    menubar.add_command(label="关于",font=("宋体",15),command=abouts)
    Groot.mainloop()
    
def CNKI():
    Croot=tk.Toplevel()
    Croot.geometry("1400x1100")
    Croot.resizable(True,True)
    Croot.title("中国知网自动化搜索结果统计")

    Croot.iconphoto(True,Base64_to_Image(中国知网_icon.img,200,200))
    Croot.configure(bg="#C4FCEF")
    frame=tk.Frame(Croot,bg='#d7e8f0')
    frame.pack(fill="both",expand=True)
    tree=ttk.Treeview(frame,columns=("第一列","第二列","第三列"),height=20,show="headings")
    tree.heading("第一列",text="待搜索内容")
    tree.heading("第二列",text="学术期刊数量")
    tree.heading("第三列",text="学术论文数量")
    tree.column("第一列",width=600)
    tree.column("第二列",width=200)
    tree.column("第三列",width=200)
    tree.grid(row=3,column=3,sticky="nsew")
    scrollbar=ttk.Scrollbar(frame,orient="vertical",command=tree.yview)
    scrollbar.grid(row=3,column=4,sticky="ns")
    tree.configure(yscrollcommand=scrollbar.set)
    def file():
        search_list=[]
        file_path=filedialog.askopenfilename(title='选择存放关键字的txt文件')
        if  os.path.exists(file_path) and os.path.isfile(file_path):
            with open(file_path, 'r', encoding='utf-8') as infile:
                lines=infile.readlines()
                for line in lines:
                    search_list.append(line)
            for line in search_list:
                   tree.insert("","end",values=(line))
            def searchmore(flag):
                def auto_export_result():
                    root1.destroy()
                    button1.config(state=tk.DISABLED)
                    messagebox.showinfo("保存",f"请选择脚本结束后自动保存的位置\n  (文件类型：xlsx)")
                    file_path=filedialog.asksaveasfilename(defaultextension=".xlsx",title="选择保存位置(文件类型：xlsx)")
                    if not os.path.exists(file_path) and os.path.isfile(file_path):
                            messagebox.showinfo("警告","请输入有效的地址！")
                    final_result=search_cnki(search_list,flag)
                    for item,value in zip(tree.get_children(),final_result['学术期刊数量']):
                        tree.set(item,"第二列",value)
                    for item,value in zip(tree.get_children(),final_result['学位论文数量']):
                        tree.set(item,"第三列",value)
                    if final_result.shape[0]==len(search_list):
                        list1=list(final_result['学术期刊数量'])
                        list2=list(final_result['学位论文数量'])
                        result1=[Convert_digitstring(element) for element in list1]
                        result2=[Convert_digitstring(element) for element in list2]
                        result1=[int(element) for element in result1]
                        result2=[int(element) for element in result2]
                     
                        dict1=dict(zip(final_result.index.tolist(),result1))
                        dict2=dict(zip(final_result.index.tolist(),result2))
                        
                        sorted_items1=sorted(dict1.items(),key=lambda item:item[1],reverse=True)
                        sorted_items2=sorted(dict2.items(),key=lambda item:item[1],reverse=True)
                        ordereddict1=OrderedDict(sorted_items1)
                        ordereddict2=OrderedDict(sorted_items2)
                        final_result.to_excel(file_path,index=True)         
                    
                       
                        
                
                        fig,axs=plt.subplots(nrows=1,ncols=2,figsize=(10,10))
                        axs[0].bar(x=list(ordereddict1.keys())[0:10],height=list(ordereddict1.values())[0:10],width=0.6,color='#00C9A7')
                        axs[1].bar(x=list(ordereddict2.keys())[0:10],height=list(ordereddict2.values())[0:10],width=0.6,color='#B39CD0') 
                        axs[0].set_xticklabels(list(ordereddict1.keys())[0:10],rotation=35)
                        axs[0].set_xlabel('关键字')
                        axs[0].set_ylabel('学术期刊数量',rotation=35)
                        axs[1].set_xticklabels(list(ordereddict2.keys())[0:10],rotation=35)
                        axs[1].set_xlabel('关键字')
                        axs[1].set_ylabel('学位论文数量',rotation=35)
                        for i,value in enumerate(list(ordereddict1.values())[0:10]):
                            axs[0].text(i,value+5,str(value),ha='center',va='bottom')
                        for i,value in enumerate(list(ordereddict2.values())[0:10]):
                            axs[1].text(i,value+5,str(value),ha='center',va='bottom') 
                        Drawsystem(Croot,fig)
                        plt.suptitle('关键字搜索结果数量Top10')
                        plt.show()
                        messagebox.showinfo("提示",f"已为您导出结果至{file_path}")
                def manual_export_result():
                    
                    def export_result():
                        file_path=filedialog.asksaveasfilename(defaultextension=".txt",title="选择保存位置")
                        if  os.path.exists(file_path) and os.path.isfile(file_path):
                            final_result.to_excel(file_path,index=True)
                            messagebox.showinfo("提示",f"已为您导出结果至{file_path}")
                        else:
                            messagebox.showinfo("警告","请输入有效的地址！")
                    root1.destroy()
                    final_result=search_cnki(search_list,flag)
                    for item,value in zip(tree.get_children(),final_result['学术期刊数量']):
                        tree.set(item,"第二列",value)
                    for item,value in zip(tree.get_children(),final_result['学位论文数量']):
                        tree.set(item,"第三列",value)
                    if final_result.shape[0]==len(search_list):
                      button1.config(state=tk.DISABLED)
                      button3=tk.Button(Croot,text="一键手动导出结果",command=export_result,width=21,height=3,relief="raised",fg="#000000",bg="#0087FF")
                      button3.place(x=1100,y=300)   
                root1=tk.Tk()
                root1.geometry("400x300")
                root1.title('选择保存方式')
                button4=tk.Button(root1,text="脚本结束后自动保存",command=auto_export_result,width=21,height=3,relief="raised",fg="#000000",bg="#0087FF")
                button4.place(x=80,y=40)
                button5=tk.Button(root1,text="脚本结束后手动保存",command=manual_export_result,width=21,height=3,relief="raised",fg="#000000",bg="#0087FF")
                button5.place(x=80,y=200)
                root1.mainloop()
            def set_flag(option):
                dic={'打开浏览器':0,'不打开浏览器':1}
                global flag
                flag=dic.get(option)
                return flag
            label1=tk.Label(Croot,text='选择浏览器模式',font=('宋体',10),fg='red',bg='#d7e8f0')
            label1.place(x=1050,y=20)
            combo1=ttk.Combobox(Croot,validate='none')
            combo1['values']=('打开浏览器','不打开浏览器')
            combo1.current(0)
            combo1.place(x=1050,y=40)
            sure_button1=tk.Button(Croot,text="确认",command=lambda:set_flag(combo1.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
            sure_button1.place(x=1300-scaling_factor+65 if scaling_factor!=150 else 1300,y=40)
            button1=tk.Button(Croot,text="立即运行脚本",command=lambda:searchmore(flag),width=21,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button1.place(x=1100,y=150)
    def helper():
        messagebox.showinfo("说明","该GUI只适用于自动化操作IEEEEXPLORE与GOOGLE学术以及中国知网\n的搜索过程并统计搜索结果，不适合其他搜索路径。\n点击运行脚本后请勿操纵鼠标，请耐心等待自动化操作完成。\n注意：导出数据时请提前建好空白xlsx文件然后选中，否则会\n覆盖掉已有txt原先内容，选中时询问替换，点击确定即可。\n                                        感谢您的使用！")
    def abouts():
        messagebox.showinfo("版本信息","开发者：Mr Crab\n版本号：8.0")    
    menubar=tk.Menu(Croot,tearoff=True)
    Croot.config(menu=menubar)
    menubar.add_command(label="导入存放搜索关键字的txt文件",font=("宋体",15),command=file)
    menubar.add_separator()
    menubar.add_command(label="说明",font=("宋体",15),command=helper)
    menubar.add_separator()
    menubar.add_command(label="关于",font=("宋体",15),command=abouts)
    Croot.mainloop()


################################################################################################################################    
'''pdf关键字查询'''
def statistic():
        Sroot=tk.Toplevel()
        Sroot.geometry("1200x900")
        Sroot.resizable(False,False)
        Sroot.title("PDF统计")
       
        Sroot.iconphoto(True,Base64_to_Image(PDF_icon.img,150,150))
        frame=tk.Frame(Sroot,bg='#d7e8f0')
        frame.pack(fill="both",expand=True)
        tree=ttk.Treeview(frame,columns=("第一列","第二列","第三列"),height=30,show="headings")
        tree.heading("第一列",text="待搜索关键词")
        tree.heading("第二列",text="含该关键字pdf数")
        tree.heading("第三列",text="该关键字在所有pdf中出现总次数")
        tree.column("第一列",width=350)
        tree.column("第二列",width=250)
        tree.column("第三列",width=350)
        tree.grid(row=3,column=3,sticky="nsew")
        scrollbar=ttk.Scrollbar(frame,orient="vertical",command=tree.yview)
        scrollbar.grid(row=3,column=4,sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)
        
        frame1=tk.Frame(Sroot)
        frame1.pack(fill="both",expand=True)
        tree1=ttk.Treeview(frame1,columns=("第一列","第二列"),height=10,show="headings")
        tree1.heading("第一列",text="待搜索pdf文件")
        tree1.heading("第二列",text="总页数")
        tree1.column("第一列",width=750)
        tree1.column("第二列",width=250)
        tree1.grid(row=6,column=10,sticky="nsew")
        scrollbar=ttk.Scrollbar(frame1,orient="vertical",command=tree.yview)
        scrollbar.grid(row=6,column=11,sticky="ns")
        tree1.configure(yscrollcommand=scrollbar.set)
        def file():
            global search_list
            search_list=[]
            messagebox.showinfo("提示","请选择存放所有关键字的txt文件")
            file_path=filedialog.askopenfilename(title='选择存放所有关键字的txt文件')
            if  os.path.exists(file_path) and os.path.isfile(file_path):
                with open(file_path, 'r', encoding='utf-8') as infile:
                    lines=infile.readlines()
                    for line in lines:
                        search_list.append(line)
                cleaned_lines=[line.replace(" ","") for line in lines]
                global cleaned_searchword_list
                cleaned_searchword_list=[line.replace("\n","") for line in search_list]
                for line in cleaned_lines:
                       tree.insert("","end",values=(line))
            else:
                messagebox.showinfo('提示','请输入有效的地址')
                cleaned_searchword_list=[]
        def pdffile():
            messagebox.showinfo("提示","请选择存放所有pdf的文件夹")
            folder_path=filedialog.askdirectory(title='选择存放所有pdf的文件夹')
            global pdf_files
            pdf_files = []
            if os.path.exists(folder_path):
                for root, dirs, files in os.walk(folder_path):
                      for file in files:
                            if file.lower().endswith('.pdf'):
                                pdf_file_path = os.path.join(root, file)
                                pdf_files.append(pdf_file_path)
                cleaned_search_list=[os.path.basename(line).replace(" ","") for line in pdf_files]
                for line in cleaned_search_list:
                           tree1.insert("","end",values=(line))
            else:
                 messagebox.showinfo('提示','未导入PDF文件夹！')
            
            def calculate_pages(pdf_files):
                total_pages=[] 
                for pdf in pdf_files:
                      pdf_document=fitz.open(pdf)
                      total_pages.append(pdf_document.page_count)
                return total_pages
            for item,value in zip(tree1.get_children(),calculate_pages(pdf_files)):
                            tree1.set(item,"第二列",value)
        def statistics():
            def read_pdf(file_path):
                if os.path.exists(file_path):
                    content = ""
                    try:
                        with fitz.open(file_path) as doc:
                            for page in doc:
                                content += page.get_text()
                    except Exception as e:
                        return None
                    return content
                else:
                    messagebox.showinfo('提示','未导入PDF文件夹！')
            def count_keywords(content, keywords):
                counts = {keyword: 0 for keyword in keywords}
                for keyword in keywords:
                    counts[keyword] = len(re.findall(r'\b' + re.escape(keyword) + r'\b', content, re.IGNORECASE))
                return counts
            keyword_stats = {keyword: {'doc_count': 0, 'total_count': 0} for keyword in cleaned_searchword_list }
            stats_data = {"关键字":cleaned_searchword_list,'含该关键字pdf数': [], '该关键字在所有pdf中出现总次数': []}   
            for pdf_file in pdf_files:
                    content = read_pdf(pdf_file)
                    if content is not None:  # 只有在成功读取内容时才继续处理
                        counts = count_keywords(content, cleaned_searchword_list)
                        for keyword, count in counts.items():
                            if count > 0:
                                keyword_stats[keyword]['doc_count'] += 1
                                keyword_stats[keyword]['total_count'] += count
            for keyword, stats in keyword_stats.items():
                stats_data['含该关键字pdf数'].append(stats['doc_count'])
                stats_data['该关键字在所有pdf中出现总次数'].append(stats['total_count'])
            stats_data=DataFrame(stats_data)
            def manual_export_result():
                        messagebox.showwarning("注意","请务必确认选中的excel文件是空白文件，且未被打开")
                        file_path=filedialog.asksaveasfilename(defaultextension=".xlsx",title="选择空白xlsx文件")
                        if  os.path.exists(file_path) and os.path.isfile(file_path):
                            stats_data.to_excel(file_path,index=False)
                            messagebox.showinfo("提示",f"已为您导出结果至{file_path}")
                            button2.config(state="disabled")
                        else:
                            messagebox.showinfo("警告","请输入有效的地址！")
            for item,value in zip(tree.get_children(),stats_data['含该关键字pdf数']):
                        tree.set(item,"第二列",value)
            for item,value in zip(tree.get_children(),stats_data['该关键字在所有pdf中出现总次数']):
                        tree.set(item,"第三列",value)
            button2=tk.Button(Sroot,text="导出结果至excel",command=manual_export_result,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button2.place(x=1010,y=500)
            button1.config(state="disabled")
            
        def helper():
            messagebox.showinfo("说明","\n点击运行脚本后请勿操纵鼠标，请耐心等待自动化操作完成。\n注意：导出数据时请提前建好空白xlsx文件然后选中，否则会\n覆盖掉已有xlsx文件原先内容，选中时询问替换，点击确定即可。\n                                        感谢您的使用！")
            
        def abouts():
            messagebox.showinfo("版本信息","开发者：Mr Crab\n版本号：8.0")
        menubar=tk.Menu(Sroot,tearoff=True)
        Sroot.config(menu=menubar)
        menubar.add_command(label="导入存放关键词的文件",font=("宋体",15),command=file)
        menubar.add_command(label="导入存放PDF的文件夹",font=("宋体",15),command=pdffile)
        menubar.add_separator()
        menubar.add_command(label="说明",font=("宋体",15),command=helper)
        menubar.add_separator()
        menubar.add_command(label="关于",font=("宋体",15),command=abouts)
        button1=tk.Button(Sroot,text="开始查询",command=statistics,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
        button1.place(x=1010,y=200)



#################################################################################################################################
'''当月论文自动化下载'''
year=datetime.now().year#获取当前年
month=datetime.now().month#获取当前月
month_dict={1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
def Light_Science_Application_interface():      
        Sroot=tk.Toplevel()
        Sroot.geometry("1200x700")
        Sroot.resizable(False,False)
        Sroot.title("Light: Science & Applications光：科学与应用")
        Sroot.iconphoto(True,Base64_to_Image(LightScience_icon.img,100,100))
      
        frame=tk.Frame(Sroot,bg='#d7e8f0')
        frame.pack(fill="both",expand=True)
        tree=ttk.Treeview(frame,columns=("第一列","第二列"),height=30,show="headings")
        tree.heading("第一列",text="pdf名称")
        tree.heading("第二列",text="下载状态")
        tree.column("第一列",width=400)
        tree.column("第二列",width=350)
        tree.grid(row=3,column=3,sticky="nsew")
        scrollbar=ttk.Scrollbar(frame,orient="vertical",command=tree.yview)
        scrollbar.grid(row=3,column=4,sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)
        class Light_science_Application():
            def __init__(self,path):
                self.urls=[]
                self.pdf_names=[]
                self.path=path
                self.flag=1
                self.browser_flag=0
                self.radio_falg=1
            def empty_folder(self):
                if os.listdir(self.path):
                    for filename in os.listdir(self.path):
                        file_path=os.path.join(self.path,filename)
                        os.remove(file_path)
                    messagebox.showinfo('提示',f'已为您清空{self.path}内所有文件')
                else:
                    messagebox.showinfo('提示','该文件夹内没有任何文件，无需清空！')
            def light_science_applications(self):
                open_Listening_mode()
                options=Options()
                options.add_argument("--disable-blink-features=AutomationControlled")#欺骗网页，隐藏自动化操作的表头
                options.add_argument('--ignore-ssl-errosr')
                options.add_argument('--ignore-certificate-errors')
                options.add_experimental_option('excludeSwitches',['enable-automation'])
                options.page_load_strategy = 'eager'
                prefs = {
	                'download.default_directory': self.path,  # 设置默认下载路径
	                "profile.default_content_setting_values.automatic_downloads": True  # 允许多文件下载
                }
                options.add_experimental_option("prefs", prefs)
                user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 S```afari/537.36'
                options.add_argument(f'user-agent={user_agent}')
                if self.browser_flag:
                        options.add_argument("-headless=new")
                        options.add_argument("--disable-gpu")
                        options.add_argument('--window-position=-2400,-2400')
                browser=webdriver.ChromiumEdge(options)
   
                browser.get("https://www.nature.com/lsa/articles")
                browser.maximize_window()
               
                year_Choose_button=browser.find_element(By.XPATH,'//*[@id="content"]/div[2]/section/div/div[2]/div/button')
                browser.execute_script('arguments[0].click()',year_Choose_button)
                year=browser.find_element(By.XPATH,'//*[@id="Year-target"]/ul/li[2]/a')
                browser.execute_script('arguments[0].click()',year)
                win=browser.window_handles
                browser.switch_to.window(win[-1])
                sleep(2)
                #查找第一页内所有论文信息，每一页的论文都被存放在一个session中
    
                Latest_essaies_section_in_Page1=browser.find_element(By.XPATH,'//*[@id="new-article-list"]/div/ul ')#
                Latest_essaies_in_page1=Latest_essaies_section_in_Page1.find_elements(By.TAG_NAME,"li")#在section容器中遍历查找所有论文
                Record1=[essay.text for essay in Latest_essaies_in_page1 ]
                #查找第二页内所有论文信息，每一页的论文都被存放在一个session中
                page2_button=browser.find_element(By.XPATH,'//*[@id="content"]/div[3]/nav/ul/li[3]/a')
                browser.execute_script('arguments[0].click()',page2_button)
                win=browser.window_handles
                browser.switch_to.window(win[-1])
                sleep(2)
                Latest_essaies_section_in_Page2=browser.find_element(By.XPATH,'//*[@id="new-article-list"]/div/ul')
                Latest_essaies_in_page2=Latest_essaies_section_in_Page2.find_elements(By.TAG_NAME,"li")#在section容器中遍历查找所有论文的位置
                Record2=[essay.text for essay in Latest_essaies_in_page2]
                #查找第三页内所有论文信息，每一页的论文都被存放在一个session中
                page3_button=browser.find_element(By.XPATH,'//*[@id="content"]/div[3]/nav/ul/li[4]/a')
                browser.execute_script('arguments[0].click()',page3_button)
                win=browser.window_handles
                browser.switch_to.window(win[-1])
                sleep(2)
                Latest_essaies_section_in_Page3=browser.find_element(By.XPATH,'//*[@id="new-article-list"]/div/ul')
                Latest_essaies_in_page3=Latest_essaies_section_in_Page3.find_elements(By.TAG_NAME,"li")#在section容器中遍历查找所有论文的位置
                Record3=[essay.text for essay in Latest_essaies_in_page3]
                #使用filter根据筛选第一页内当前月份内的论文信息,并使用text.splitlines()[0]得到第一页内当月论文的题目
                Record1=filter(lambda element:month_dict.get(month) in element,Record1)
                Record1=list(Record1)
                essay_titles_in_page1=[text.splitlines()[0] for text in Record1]#第一页内当月论文的题目
                #使用filter筛选第二页内当前月份内的论文信息，使用text.splitlines()[0]得到第二页内当月论文的题目
                Record2=filter(lambda element:month_dict.get(month) in element,Record2)
                Record2=list(Record2)
                essay_titles_in_page2=[text.splitlines()[0] for text in Record2]#第二页内当月论文的题目
                #使用filter筛选第三页内当前月份内的论文信息，使用text.splitlines()[0]得到第三页内当月论文的题目
                Record3=filter(lambda element:month_dict.get(month) in element,Record3)
                Record3=list(Record3)
                essay_titles_in_page3=[text.splitlines()[0] for text in Record3]#第三页内当月论文的题目
                self.essay_titles=essay_titles_in_page1+essay_titles_in_page2+essay_titles_in_page3
                #根据排列组合的知识，前三页的每一页是否有当月论文，共有2**3 8种可能，其中，若第一页没有，后两页一定没有，以及第一页和第三页不可能同时有，这可以排除掉5种情况。故剩下3种情况为：
               # 第一页有，第二三页都没有
                #第一页有，第二页有，第三页没有
                #第一二三页全都有，这是最极端的情况，这说明该期刊这个月收录了60篇论文，按照往年的历史数据来看这一事件发生概率极低
                if len(essay_titles_in_page1)==0:#第一页内没有当月论文，那么这个期刊这个月一定还未收录任何论文
                    self.flag=0
                    browser.quit()
                elif (len(essay_titles_in_page1)!=0 and len(essay_titles_in_page2)==0 and len(essay_titles_in_page3)==0):#第一页有，第二三页都没有
                        page1_button=browser.find_element(By.XPATH,'//*[@id="content"]/div[3]/nav/ul/li[2]/a')
                        page1_button=browser.execute_script('arguments[0].click()',page1_button)
                        win=browser.window_handles
                        browser.switch_to.window(win[-1])
                        for i in range(len(essay_titles_in_page1)):
                            title=browser.find_element(By.PARTIAL_LINK_TEXT,essay_titles_in_page1[i])
                            browser.execute_script('arguments[0].click()',title)
                            win=browser.window_handles
                            browser.switch_to.window(win[-1])
                            sleep(1)
                            self.urls.append(browser.current_url)
                            browser.back()
                elif(len(essay_titles_in_page1)!=0 and len(essay_titles_in_page2)!=0 and len(essay_titles_in_page3)==0):#第一页有，第二页有，第三页没有
                    page1_button=browser.find_element(By.XPATH,'//*[@id="content"]/div[3]/nav/ul/li[2]/a')
                    page1_button=browser.execute_script('arguments[0].click()',page1_button)
                    win=browser.window_handles
                    browser.switch_to.window(win[-1])
                    for i in range(len(essay_titles_in_page1)):
                        sleep(2)
                        title=browser.find_element(By.PARTIAL_LINK_TEXT,essay_titles_in_page1[i])
                        browser.execute_script('arguments[0].click()',title)
                        win=browser.window_handles
                        browser.switch_to.window(win[-1])
                        sleep(1)
                        self.urls.append(browser.current_url)
                        browser.back() 
                    page2_button=browser.find_element(By.XPATH,'//*[@id="content"]/div[3]/nav/ul/li[3]/a')
                    browser.execute_script('arguments[0].click()',page2_button)
                    win=browser.window_handles
                    browser.switch_to.window(win[-1])
                    for i in range(len(essay_titles_in_page2)):
                        title=browser.find_element(By.PARTIAL_LINK_TEXT,essay_titles_in_page2[i])
                        browser.execute_script('arguments[0].click()',title)
                        win=browser.window_handles
                        browser.switch_to.window(win[-1])
                        sleep(1)
                        self.urls.append(browser.current_url)
                        browser.back()
                else:                                                                                  #第一二三页全都有，这是最极端的情况
                    page1_button=browser.find_element(By.XPATH,'//*[@id="content"]/div[3]/nav/ul/li[2]/a')
                    page1_button=browser.execute_script('arguments[0].click()',page1_button)
                    win=browser.window_handles
                    browser.switch_to.window(win[-1])
                    for i in range(len(essay_titles_in_page1)):
                        title=browser.find_element(By.PARTIAL_LINK_TEXT,essay_titles_in_page1[i])
                        browser.execute_script('arguments[0].click()',title)
                        win=browser.window_handles
                        browser.switch_to.window(win[-1])
                        sleep(1)
                        self.urls.append(browser.current_url)
                        browser.back()
                    page2_button=browser.find_element(By.XPATH,'//*[@id="content"]/div[3]/nav/ul/li[3]/a')
                    browser.execute_script('arguments[0].click()',page2_button)
                    win=browser.window_handles
                    browser.switch_to.window(win[-1])
                    for i in range(len(essay_titles_in_page2)):
                        title=browser.find_element(By.PARTIAL_LINK_TEXT,essay_titles_in_page2[i])
                        browser.execute_script('arguments[0].click()',title)
                        win=browser.window_handles
                        browser.switch_to.window(win[-1])
                        self.urls.append(browser.current_url)
                        browser.back()
                    page3_button=browser.find_element(By.XPATH,'//*[@id="content"]/div[3]/nav/ul/li[4]/a')
                    browser.execute_script('arguments[0].click()',page3_button)
                    win=browser.window_handles
                    browser.switch_to.window(win[-1])
                    for i in range(len(essay_titles_in_page3)):
                        title=browser.find_element(By.PARTIAL_LINK_TEXT,essay_titles_in_page3[i])
                        browser.execute_script('arguments[0].click()',title)
                        win=browser.window_handles
                        browser.switch_to.window(win[-1])
                        self.urls.append(browser.current_url)
                        browser.back()
                browser.quit()
                close_Listening_mode()
            def Check(self,filename):
                for root, dirs, files in os.walk(self.path):
                    if not files:
                        return True 
                    else:
                        for file in files:
                            if filename not in files:
                                return True
                            else:
                                return False 
            def change_browser_flag(self,option):
                dic={'不打开浏览器':1,'打开浏览器':0}
                self.browser_flag=dic.get(option)
                return self.browser_flag
            def change_radio_flag(self,option):
                dic={'语音播报':1,'不语音播报':0}
                self.radio_flag=dic.get(option)
                return self.radio_flag 
            def get_pdf_names(self):
                   #得到每一个论文pdf下载时的名称，为了后序判断是否在目标文件夹中
                for url in self.urls:
                    self.pdf_names.append(url[32:]+'.pdf')
                full_path=os.path.join(self.path,'下载论文列表.txt')
                with open(full_path,'w',encoding='utf-8') as pdf:
                   for ele in self.pdf_names:
                       pdf.write(ele+'\n')
            def download(self):
                open_Listening_mode()
                if self.flag==1:
                    check=dict(zip(self.pdf_names,self.urls))
                    options = Options()
                    options.add_argument("--disable-blink-features=AutomationControlled")#欺骗网页，隐藏自动化操作的表头
                    options.add_argument('--ignore-ssl-errors')#忽略ssl错误，发生在浏览器向数据库发送请求过程中
                    options.add_argument('--ignore-certificate-errors')
                    options.add_experimental_option('excludeSwitches',['enable-automation'])
                    options.page_load_strategy = 'eager'
                    prefs = {
	                    'download.default_directory': self.path,  # 设置默认下载路径
	                    "profile.default_content_setting_values.automatic_downloads": True  # 允许多文件下载
                    }
                    options.add_experimental_option("prefs", prefs)
                    user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 S```afari/537.36'
                    options.add_argument(f'user-agent={user_agent}')
                    if self.browser_flag:
                        options.add_argument("-headless=new")
                        options.add_argument("--disable-gpu")
                        options.add_argument('--window-position=-2400,-2400')
                    browser=webdriver.ChromiumEdge(options)
                    browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
              "source": """
                Object.defineProperty(navigator, 'webdriver', {
                  get: () => undefined
                })
              """})
                    for pdf in self.pdf_names:
                         if self.Check(pdf):
                             browser.get(check.get(pdf))
                             sleep(3)
                             download_button=browser.find_element(By.XPATH,'//*[@id="content"]/aside/div[1]/div/a')
                             browser.execute_script('arguments[0].click()',download_button)
                             for item in tree.get_children():      
                                tree.set(item,"第二列",'已下载')
                             sleep(20)
                         else:
                              for item in tree.get_children():    
                                 tree.set(item,"第二列",'本地文件夹中已存在')
                              pass
                    sleep(60)
                    browser.quit()
                    if self.radio_flag:
                        set_Volume_to_100()
                        Speak(f'Light: Science & Applications光：科学与应用 {year}年{month}月发布的最新论文已经为您下载到{self.path}，请查看')
                    messagebox.showinfo('提示',f"Light: Science & Applications光：科学与应用 {year}年{month}月发布的最新论文已经为您下载到{self.path}，请查看")
                else:
                    if self.radio_flag:
                        set_Volume_to_100()
                        Speak(f"Light: Science & Applications光：科学与应用 {year}年{month}月还未收录任何论文，请耐心等待！")
                    messagebox.showinfo('注意',f"Light: Science & Applications光：科学与应用 {year}年{month}月还未收录任何论文，请耐心等待！")
                self.urls.clear()
                self.pdf_names.clear()
                close_Listening_mode()
                

            def search(self):
                button1.config(state="disabled")
                thread1=threading.Thread(target=self.light_science_applications)
                thread2=threading.Thread(target=self.get_pdf_names)
                thread1.start()
                thread1.join()
                if self.flag!=0:
                    thread2.start()
                    thread2.join()
                    for pdf_name in self.pdf_names:
                        tree.insert("","end",values=(pdf_name))
                    sleep(5)
                    self.download()
                else:
                    if self.radio_flag:
                        set_Volume_to_100()
                        Speak(f"该期刊今年{month}月还未收录任何论文，请耐心等待！")
                    messagebox.showinfo('注意',f"该期刊今年{month}月还未收录任何论文，请耐心等待！") 
                     
        def file():
            messagebox.showinfo("提示","请选择存放所有下载pdf的文件夹")
            folder_path=filedialog.askdirectory(title='选择存放所有下载pdf的文件夹')
            if not os.path.exists(folder_path):
                messagebox.showinfo('注意','请输入有效的地址！')
            folder_path=folder_path.replace('/','\\')
            Light=Light_science_Application(folder_path)
            global button1
            combo1=ttk.Combobox(Sroot,validate='none')
            label1=tk.Label(Sroot,text='选择浏览器模式',font=('宋体',10),fg='red',bg='#d7e8f0')
            label1.place(x=830,y=20)
            combo1['values']=('打开浏览器')
            combo1.current(0)
            combo1.place(x=830,y=40)
            sure_button1=tk.Button(Sroot,text="确认",command=lambda:Light.change_browser_flag(combo1.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
            sure_button1.place(x=1120-scaling_factor+65 if scaling_factor!=150 else 1120,y=40)
            combo2=ttk.Combobox(Sroot,validate='none')
            label2=tk.Label(Sroot,text='选择是否语音播报',font=('宋体',10),fg='red',bg='#d7e8f0')
            label2.place(x=830,y=120)
            combo2['values']=('语音播报','不语音播报')
            combo2.place(x=830,y=140)
            sure_button2=tk.Button(Sroot,text="确认",command=lambda:Light.change_radio_flag(combo2.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
            sure_button2.place(x=1120-scaling_factor+65 if scaling_factor!=150 else 1120,y=140)
            combo2.current(0)
            button1=tk.Button(Sroot,text="运行脚本",command=Light.search,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button1.place(x=880,y=250)
            button2=tk.Button(Sroot,text="清空文件夹内文件",command=Light.empty_folder,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button2.place(x=880,y=400)
        def helper():
                messagebox.showinfo("说明","\n点击运行脚本后请勿操纵鼠标，请耐心等待自动化操作完成。\n                                        感谢您的使用！")
        def abouts():
                messagebox.showinfo("版本信息","开发者：Mr Crab\n版本号：8.0")
       
        menubar=tk.Menu(Sroot,tearoff=True)
        Sroot.config(menu=menubar)
        menubar.add_command(label="选择存放pdf的文件夹",font=("宋体",15),command=file)
        menubar.add_separator()
        menubar.add_command(label="说明",font=("宋体",15),command=helper)
        menubar.add_separator()
        menubar.add_command(label="关于",font=("宋体",15),command=abouts)
         
        
def SCIENCE_CHINA_Information_Sciences_interface():      
        Sroot=tk.Toplevel()
        Sroot.geometry("1200x700")
        Sroot.resizable(False,False)
       
        Sroot.title("SCIENCE CHINA Information Sciences中国科学")
        Sroot.iconphoto(True,Base64_to_Image(Springer_icon.img,200,200))
        frame=tk.Frame(Sroot,bg='#d7e8f0')
        frame.pack(fill="both",expand=True)
        tree=ttk.Treeview(frame,columns=("第一列","第二列"),height=30,show="headings")
        tree.heading("第一列",text="论文pdf名称")
        tree.heading("第二列",text="下载状态")
        tree.column("第一列",width=400)
        tree.column("第二列",width=350)
        tree.grid(row=3,column=3,sticky="nsew")
        scrollbar=ttk.Scrollbar(frame,orient="vertical",command=tree.yview)
        scrollbar.grid(row=3,column=4,sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)
        class SCIENCE_CHINA_Information_Sciences():
            def __init__(self,path):
                 self.path=path
                 self.pdf_names=[]
                 self.urls=[]
                 self.browser_flag=0
                 self.radio_flag=0
                 self.flag=1
            def empty_folder(self):
                if os.listdir(self.path):
                    for filename in os.listdir(self.path):
                        file_path=os.path.join(self.path,filename)
                        os.remove(file_path)
                    messagebox.showinfo('提示',f'已为您清空{self.path}内所有文件')
                else:
                    messagebox.showinfo('提示','该文件夹内没有任何文件，无需清空！')
            def change_browser_flag(self,option):
                dic={'不打开浏览器':1,'打开浏览器':0}
                self.browser_flag=dic.get(option)
                return self.browser_flag
            def change_radio_flag(self,option):
                dic={'语音播报':1,'不语音播报':0}
                self.radio_flag=dic.get(option)
                return self.radio_flag 
            def science_CHINA_Information_Sciences(self):
                        open_Listening_mode()
                        options=Options()
                        options.add_argument("--disable-blink-features=AutomationControlled")#欺骗网页，隐藏自动化操作的表头
                        options.add_argument('--ignore-ssl-errosr')
                        options.add_argument('--ignore-certificate-errors')
                        options.add_experimental_option('excludeSwitches',['enable-automation'])
                        options.page_load_strategy = 'eager'
                        prefs = {
	                            'download.default_directory': self.path,  # 设置默认下载路径
	                            "profile.default_content_setting_values.automatic_downloads": True  # 允许多文件下载
                            }
                        options.add_experimental_option("prefs", prefs)
                        if self.browser_flag:
                            options.add_argument("-headless=new")
                            options.add_argument("--disable-gpu")
                            options.add_argument('--window-position=-2400,-2400')
                        self.browser=webdriver.ChromiumEdge(options)
                        self.browser.get("https://link.springer.com/journal/11432")  
                        self.browser.maximize_window()
                        view_all_articles_button=self.browser.find_element(By.XPATH,'//*[@id="main"]/section[2]/div/div/div[1]/a')
                        self.browser.execute_script('arguments[0].click()',view_all_articles_button)
                        win=self.browser.window_handles
                        self.browser.switch_to.window(win[-1])
                        sleep(1)
                        year_choose=self.browser.find_element(By.XPATH,'//*[@id="filter-by-volume"]/option[2]')#年份永远选择最新时间的，即下拉列表第二个，第一个是所有，第二个是当前年份
                        year_choose.click()
                        win=self.browser.window_handles
                        self.browser.switch_to.window(win[-1])
                        sleep(1)
                        #将二页内100篇论文信息摘录下来到Record中
                        Latest_Journal_Containers_Page1=self.browser.find_element(By.XPATH,'//*[@id="main"]/div/div/div/section')#第一页所有论文存放在一个section容器中
                        Latest_Journals_in_page1=Latest_Journal_Containers_Page1.find_elements(By.TAG_NAME,"li")#在section容器中遍历查找所有论文的位置
                        Record1=[journal.text for journal in Latest_Journals_in_page1]#将所有论文的相关信息存放到列表Record1中
                        next_button=self.browser.find_element(By.XPATH,'//*[@id="main"]/div/div/div/nav/ul/li[5]/a')
                        self.browser.execute_script('arguments[0].click()',next_button)
                        win=self.browser.window_handles
                        self.browser.switch_to.window(win[-1])
                        year_choose=self.browser.find_element(By.XPATH,'//*[@id="filter-by-volume"]/option[2]')#年份永远选择最新时间的，即下拉列表第二个，第一个是所有，第二个是当前年份
                        year_choose.click()
                        win=self.browser.window_handles
                        self.browser.switch_to.window(win[-1])
                        sleep(2)
                        Latest_Journal_Containers_Page2=self.browser.find_element(By.XPATH,'//*[@id="main"]/div/div/div/section')#第二页所有论文存放在一个section容器中
                        Latest_Journals_in_page2=Latest_Journal_Containers_Page2.find_elements(By.TAG_NAME,"li")#在section容器中遍历查找所有论文的位置
                        Record2=[journal.text for journal in Latest_Journals_in_page2]#将所有论文的相关信息存放到列表Record2中
                        Record=Record1+Record2
                        Record=filter(lambda element:month_dict.get(month) in element,Record)#使用filter筛选当前月份内的论文
                        Record=list(Record)
                        search=[text.splitlines()[0] for text in Record if len(text)>20]
                        self.browser.back()
                        if len(Record)==0:#没有当前月份的
                            self.flag=0
                            self.browser.quit()
                            close_Listening_mode()  
                        # 论文列表
                        else:               
                            for i in range(len(search)):
                                try:
                                    essay_button=self.browser.find_element(By.PARTIAL_LINK_TEXT,search[i])
                                    self.browser.execute_script('arguments[0].click()',essay_button)
                                    win=self.browser.window_handles
                                    self.browser.switch_to.window(win[-1])
                                    url=self.browser.current_url
                                    self.urls.append(url)
                                except NoSuchElementException:
                                    page2=self.browser.find_element(By.XPATH,'//*[@id="main"]/div/div/div/nav/ul/li[2]/a')
                                    self.browser.execute_script("arguments[0].click()",page2)
                                    win=self.browser.window_handles
                                    self.browser.switch_to.window(win[-1])
                                    sleep(2)
                                    essay_button=self.browser.find_element(By.PARTIAL_LINK_TEXT,search[i])
                                    self.browser.execute_script('arguments[0].click()',essay_button)
                                    win=self.browser.window_handles
                                    self.browser.switch_to.window(win[-1])
                                    url=self.browser.current_url
                                    self.urls.append(url)
                                finally:
                                    self.browser.back()
                        self.browser.quit()  
                        close_Listening_mode()
            def download(self):
                open_Listening_mode()
                if self.flag!=0:
                    check=dict(zip(self.pdf_names,self.urls))
                    option = Options()
                    option.add_argument("--disable-blink-features=AutomationControlled")#欺骗网页，隐藏自动化操作的表头
                    option.add_argument('--ignore-ssl-errors')#忽略ssl错误，发生在浏览器向数据库发送请求过程中
                    option.add_argument('--ignore-certificate-errors')
                    option.add_experimental_option('excludeSwitches',['enable-automation'])
                    #option.page_load_strategy = 'eager'
                    prefs = {
	                    'download.default_directory': self.path,  # 设置默认下载路径为自定义路径
	                    "profile.default_content_setting_values.automatic_downloads": True  # 允许多文件下载
                    }
                    option.add_experimental_option("prefs", prefs)
                    if self.browser_flag:
                        option.add_argument("-headless=new")
                        option.add_argument("--disable-gpu")
                        option.add_argument('--window-position=-2400,-2400')
                    browser=webdriver.ChromiumEdge(option) 
                    
                    browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
              "source": """
                Object.defineProperty(navigator, 'webdriver', {
                  get: () => undefined
                })
              """})
                    for pdf in self.pdf_names:             
                        if self.Check(pdf):
                                browser.get(check.get(pdf))
                                browser.maximize_window()
                                sleep(2)
                                try:
                                    login_via_an_institution=browser.find_element(By.PARTIAL_LINK_TEXT,'Log in via an institution')
                                    browser.execute_script('arguments[0].click()',login_via_an_institution)
                                    win=browser.window_handles
                                    browser.switch_to.window(win[-1])
                                    input_institution=browser.find_element(By.XPATH,'//*[@id="searchFormTextInput"]')
                                    input_institution.send_keys('Hefei University of Technology')
                                    pyautogui.hotkey('Enter')
                                    sleep(5)
                                    school=browser.find_element(By.XPATH,'//*[@id="autocomplete-results"]/a')
                                    browser.execute_script("arguments[0].click()",school)
                                    win=browser.window_handles
                                    browser.switch_to.window(win[-1])
                                    sleep(10)
                               
                                    accounts=str(base64.b64decode(bytes(account_and_pasword.get('account').encode('utf-8'))).decode('utf-8'))
                                    passwords=str(base64.b64decode(bytes(account_and_pasword.get('password').encode('utf-8'))).decode('utf-8'))
                                    account=browser.find_element(By.XPATH,'//*[@id="username"]')
                                    account.click()
                                    account.send_keys(accounts)
                                    password=browser.find_element(By.XPATH,'//*[@id="pwd"]')
                                    password.click()
                                    password.send_keys(passwords)
                                    login_button=browser.find_element(By.ID,'sb2')
                                    login_button.click()
                                    win=browser.window_handles
                                    browser.switch_to.window(win[-1])
                                    sleep(5)
                                    accept_button=browser.find_element(By.XPATH,'/html/body/form/div/div[2]/p[2]/input[2]')
                                    browser.execute_script('arguments[0].click()',accept_button)
                                    win=browser.window_handles
                                    browser.switch_to.window(win[-1])
                                    sleep(5)
                                    download_button=browser.find_element(By.XPATH,'//*[@id="main"]/section/div/div/div[1]/div/div/div/a')
                                    browser.execute_script('arguments[0].click()',download_button)
                                    sleep(10)
                                    for item in tree.get_children():      
                                        tree.set(item,"第二列",'已下载')
                                    sleep(5)
                                except NoSuchElementException:
                                    sleep(5)
                                    download_button=browser.find_element(By.XPATH,'//*[@id="main"]/section/div/div/div[1]/div/div/div/a')
                                    browser.execute_script('arguments[0].click()',download_button)
                                    for item in tree.get_children():      
                                        tree.set(item,"第二列",'已下载')
                                    sleep(10)
                        else:
                            for item in tree.get_children():    
                                tree.set(item,"第二列",'本地文件夹中已存在')
                            pass
                    sleep(60)
                    browser.quit()
                    close_Listening_mode()
                    if self.radio_flag:
                        set_Volume_to_100()
                        Speak(f"SCIENCE CHINA Information Sciences中国科学：信息科学（英文版）{year}年{month}月发布的最新论文已经为您下载到{self.path}，请查看")
                    messagebox.showinfo('提示',f"SCIENCE CHINA Information Sciences中国科学：信息科学（英文版）{year}年{month}月发布的最新论文已经为您下载到{self.path}，请查看")
                    
                else:      
                    
                    browser.quit()
                    close_Listening_mode()
                    if self.radio_flag:
                        set_Volume_to_100()
                        Speak(f"SCIENCE CHINA Information Sciences中国科学：信息科学（英文版）{year}年{month}月发布的最新论文已经为您下载到{self.path}，请查看")
                    messagebox.showinfo('注意',f"SCIENCE CHINA Information Sciences中国科学：信息科学（英文版） {year}年{month}月还未收录任何论文，请耐心等待！")
                self.urls.clear()
                self.pdf_names.clear()
            def get_pdf_names(self):
                self.pdf_names=[]
                for url in self.urls:
                    self.pdf_names.append(url[42:]+'.pdf')
                full_path=os.path.join(self.path,'下载论文列表.txt')
                with open(full_path,'w',encoding='utf-8') as pdf:
                    for ele in self.pdf_names:
                        pdf.write(ele+'\n')
                #得到每一个论文pdf下载时的名称，为了后序判断是否在目标文件夹中
            def Check(self,filename):
                for root, dirs, files in os.walk(self.path):
                    if len(files)==0:
                        return True 
                    else:
                        for file in files:
                            if filename not in files:
                                return True
                            else:
                                return False          
            
            def search(self):
                button1.config(state="disabled")
                thread1=threading.Thread(target=self.science_CHINA_Information_Sciences)
                thread2=threading.Thread(target=self.get_pdf_names)
                thread1.start()
                thread1.join()
                if self.flag!=0:
                    thread2.start()
                    thread2.join()
             
                    for pdf_name in self.pdf_names:
                        tree.insert("","end",values=(pdf_name))
                    sleep(5)
                    self.download()
                else:
                    if self.radio_flag:
                        set_Volume_to_100()
                        Speak(f"该期刊今年{month}月还未收录任何论文，请耐心等待！")
                    messagebox.showinfo('注意',f"该期刊今年{month}月还未收录任何论文，请耐心等待！") 
                    
               
        def file():
            messagebox.showinfo("提示","请选择存放所有下载pdf的文件夹")
            folder_path=filedialog.askdirectory(title='选择存放所有下载pdf的文件夹')
            if not os.path.exists(folder_path):
                messagebox.showinfo('注意','请输入有效的地址！')
            folder_path=folder_path.replace('/','\\')
            science=SCIENCE_CHINA_Information_Sciences(folder_path)
            global button1
            combo1=ttk.Combobox(Sroot,validate='none')
            label1=tk.Label(Sroot,text='选择浏览器模式',font=('宋体',10),fg='red',bg='#d7e8f0')
            label1.place(x=830,y=20)
            combo1['values']=('打开浏览器')
            combo1.current(0)
            combo1.place(x=830,y=40)
            sure_button1=tk.Button(Sroot,text="确认",command=lambda:science.change_browser_flag(combo1.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
            sure_button1.place(x=1120-scaling_factor+65 if scaling_factor!=150 else 1120,y=40)
            combo2=ttk.Combobox(Sroot,validate='none')
            label2=tk.Label(Sroot,text='选择是否语音播报',font=('宋体',10),fg='red',bg='#d7e8f0')
            label2.place(x=830,y=120)
            combo2['values']=('语音播报','不语音播报')
            combo2.place(x=830,y=140)
            sure_button2=tk.Button(Sroot,text="确认",command=lambda:science.change_radio_flag(combo2.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
            sure_button2.place(x=1120-scaling_factor+65 if scaling_factor!=150 else 1120,y=140)
            combo2.current(0)
            button1=tk.Button(Sroot,text="运行脚本",command=science.search,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button1.place(x=880,y=250)
            button2=tk.Button(Sroot,text="清空文件夹内文件",command=science.empty_folder,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button2.place(x=880,y=400)
        def helper():
            messagebox.showinfo("说明","\n点击运行脚本后请勿操纵鼠标，请耐心等待自动化操作完成。\n                                        感谢您的使用！")
            
        def abouts():
            messagebox.showinfo("版本信息","开发者：Mr Crab\n版本号：8.0")
        menubar=tk.Menu(Sroot,tearoff=True)
        menubar.add_command(label="选择存放pdf的文件夹",font=("宋体",15),command=file)
        menubar.add_separator()
        menubar.add_command(label="说明",font=("宋体",15),command=helper)
        menubar.add_separator()
        menubar.add_command(label="关于",font=("宋体",15),command=abouts)
        Sroot.config(menu=menubar)
       
def journal_of_lightwave_technology_interface():    
        
        jroot=tk.Toplevel()
        jroot.geometry("1250x700")
        jroot.resizable(False,False)
       
        jroot.title("Journal of Lightwave Technology")
        jroot.iconphoto(True,Base64_to_Image(IEEE_icon.img,100,100))
        frame=tk.Frame(jroot,bg='#d7e8f0')
        frame.pack(fill="both",expand=True)
        tree=ttk.Treeview(frame,columns=("第一列","第二列"),height=30,show="headings")
        tree.heading("第一列",text="论文名称")
        tree.heading("第二列",text="下载状态")
        tree.column("第一列",width=650)
        tree.column("第二列",width=150)
        tree.grid(row=3,column=3,sticky="nsew")
        scrollbar=ttk.Scrollbar(frame,orient="vertical",command=tree.yview)
        scrollbar.grid(row=3,column=4,sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)
        class Journal_of_Lightwave_Technology():
            def __init__(self,path):
                 self.path=path
                 self.essay_titles=[]
                 self.browser_flag=0
                 self.radio_flag=1
            def empty_folder(self):
                if os.listdir(self.path):
                    for filename in os.listdir(self.path):
                        file_path=os.path.join(self.path,filename)
                        os.remove(file_path)
                    messagebox.showinfo('提示',f'已为您清空{self.path}内所有文件')
                else:
                    messagebox.showinfo('提示','该文件夹内没有任何文件，无需清空！')
            def extract_zip_file(self):
                        filelist=os.listdir(self.path)
                        for file in filelist:
                            if file.endswith('.zip'):
                                zip_file_path=os.path.join(self.path,file)
                                with zipfile.ZipFile(zip_file_path,'r') as zip:
                                    zip.extractall(self.path)
            def delete_file(self):
                if os.listdir(self.path):
                    for filename in os.listdir(self.path):
                        if filename.endswith('zip'):
                            zip_path=os.path.join(self.path,filename)
                            os.remove(zip_path)
                        else:
                            pass
                        if filename.endswith('.pdf'):
                                pdf_path=os.path.join(self.path,filename)
                                os.remove(pdf_path)
                        else:
                            pass 

                else:
                    pass
            def change_browser_flag(self,option):
                dic={'不打开浏览器':1,'打开浏览器':0}
                self.browser_flag=dic.get(option)
                return self.browser_flag
            def change_radio_flag(self,option):
                dic={'语音播报':1,'不语音播报':0}
                self.radio_flag=dic.get(option)
                return self.radio_flag 
            def journal_of_Lightwave_Technology(self):
                    open_Listening_mode()
                    button1.config(state='disabled')
                    options=Options()
                    options.add_argument("--disable-blink-features=AutomationControlled")#欺骗网页，隐藏自动化操作的表头
                    options.add_argument('--ignore-ssl-errosr')
                    options.add_argument('--ignore-certificate-errors')
                    options.add_experimental_option('excludeSwitches',['enable-automation'])
                    prefs = {
	                    'download.default_directory': self.path,  # 设置默认下载路径
	                    "profile.default_content_setting_values.automatic_downloads": 1  # 允许多文件下载
                    }
                    options.add_experimental_option("prefs", prefs)
                    if self.browser_flag:
                        options.add_argument("-headless=new")
                        options.add_argument("--disable-gpu")
                        options.add_argument('--window-position=-2400,-2400')
                    browser=webdriver.ChromiumEdge(options)
                    browser.get('https://ieeexplore.ieee.org/Xplore/guesthome.jsp')
                    browser.maximize_window()
                    sleep(5)
                    try:
                        sign=browser.find_element(By.XPATH,'//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/xpl-navbar/div/div[1]/div[3]/div/xpl-institution-details/div/div/div')
                    except  NoSuchElementException:
                        WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/xpl-navbar/div/div[1]/div[3]/xpl-login-modal-trigger/a')))
                        institution_log_in_button=browser.find_element(By.XPATH,'//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/xpl-navbar/div/div[1]/div[3]/xpl-login-modal-trigger/a')
                        institution_log_in_button.click()
                        sleep(6)
                        access_through_institution_button=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/xpl-login-modal/div[1]/div[2]/div/div/xpl-login/div/section/div/div/xpl-seamless-access/div/div[1]/button/div/div[2]/div')
                        access_through_institution_button.click()
                        sleep(1)
                        input=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/xpl-login-modal/div[1]/div[2]/div/div/xpl-login/div/section/div/div/div[2]/div[2]/xpl-inst-typeahead/div/div/input')
                        input.send_keys('Hefei University of Technology')
                        sleep(5)
                        pyautogui.hotkey('enter')
                        school=browser.find_element(By.XPATH,'//*[@id="Hefei University of Technology"]')
                        browser.execute_script('arguments[0].click()',school)
                        win=browser.window_handles
                        browser.switch_to.window(win[-1])
                        sleep(5)
                        accounts=str(base64.b64decode(bytes(account_and_pasword.get('account').encode('utf-8'))).decode('utf-8'))
                        passwords=str(base64.b64decode(bytes(account_and_pasword.get('password').encode('utf-8'))).decode('utf-8'))
                        account=browser.find_element(By.XPATH,'//*[@id="username"]')
                        account.click()
                        account.send_keys(accounts)
                        password=browser.find_element(By.XPATH,'//*[@id="pwd"]')
                        password.click()
                        password.send_keys(passwords)
                        login_button=browser.find_element(By.ID,'sb2')
                        login_button.click()
                        win=browser.window_handles
                        browser.switch_to.window(win[-1])
                        sleep(5)
                        accept_button=browser.find_element(By.XPATH,'/html/body/form/div/div[2]/p[2]/input[2]')
                        browser.execute_script('arguments[0].click()',accept_button)
                        win=browser.window_handles
                        browser.switch_to.window(win[-1])
                    sleep(10) 
                    input_element=browser.find_element(By.XPATH, '//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/div[2]/div[2]/xpl-search-bar-migr/div/form/div[2]/div/div[1]/xpl-typeahead-migr/div/input')
                    input_element.send_keys('Journal of Lightwave Technology')
                    input_element.send_keys(Keys.ENTER)
                    win=browser.window_handles
                    browser.switch_to.window(win[-1])
                    sleep(10) 
                    journal_button=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[2]/xpl-suggested-publication-list/div/div[2]/div[1]/a/div')
                    browser.execute_script('arguments[0].click()',journal_button)
                    win=browser.window_handles
                    browser.switch_to.window(win[-1])
                    sleep(10) 
                    latest_journal_button=browser.find_element(By.XPATH,'//*[@id="toggle-button-2"]')
                   # latest_journal_button=browser.find_element(By.XPATH,'//*[@id="xplMainContentLandmark"]/div/xpl-xpl-delegate/xpl-journals/div/div[1]/div[3]/section/xpl-journals-home/section/section/section[2]/div[2]/div[1]/div[2]')
                    browser.execute_script('arguments[0].click()',latest_journal_button)
                    sleep(10)
                    show_all=browser.find_element(By.XPATH,'//*[@id="xplMainContentLandmark"]/div/xpl-xpl-delegate/xpl-journals/div/div[1]/div[3]/section/xpl-journals-home/section/section/section[2]/div[2]/xpl-journal-recent-article/div[2]/a')
                    browser.execute_script("arguments[0].click()",show_all)
                    win=browser.window_handles
                    browser.switch_to.window(win[-1])
                    sleep(15)
                    total_journals1=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[2]/div[1]/div/div/span[1]/span[2]')
                    total_journals=int(total_journals1.text)
                    items_per_page=browser.find_element(By.XPATH,'//*[@id="dropdownPerPageLabel"]')
                    browser.execute_script('arguments[0].click()',items_per_page)
                    sleep(2)
                    fifity_per_page=browser.find_element(By.CSS_SELECTOR,r'#publicationIssueMainContent\ global-margin-px > div.ng-Dashboard > xpl-issue-search-dashboard > div > div.col-12.action-bar.Dashboard-section.hide-mobile > ul > div > div:nth-child(2) > xpl-rows-per-page-drop-down > div > div > button:nth-child(3)')
                    browser.execute_script('arguments[0].click()',fifity_per_page)
                    sleep(10)
                    self.delete_file()
                    if total_journals==50:
                                    for k in range(3,53,10):
                                        sleep(3)
                                        for i in range(k,k+10):    
                                            try:
                                                
                                                select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                browser.execute_script("arguments[0].click()",select)
                                                self.essay_titles.append(essay_title.text)
                                            except NoSuchElementException:
                                                pass
                                        download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                                        browser.execute_script("arguments[0].click()",download_pdf)
                                        download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                                        browser.execute_script("arguments[0].click()",download)
                                        WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                                        sleep(2)
                                        close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                                        browser.execute_script("arguments[0].click()",close)
                                        for i in range(k,k+10):  
                                            try:
                                                select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                browser.execute_script("arguments[0].click()",select)
                                                sleep(0.5)
                                            except NoSuchElementException:
                                                pass
                                        sleep(10)
                                    pass
                    elif total_journals<50:
        
                        res=total_journals%10
                        int_part=total_journals//10
                        for k in range(2,int_part*10+2,10):
                                
                                        sleep(3)
                                        for i in range(k,k+10):    
                                            try:
                                                select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                browser.execute_script("arguments[0].click()",select)
                                                self.essay_titles.append(essay_title.text)
                                                sleep(0.5)
                                            except NoSuchElementException:
                                                pass
                        
                                        download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                                        browser.execute_script("arguments[0].click()",download_pdf)
                                        download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                                        browser.execute_script("arguments[0].click()",download)
                                        WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                                        sleep(2)
                                        close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                                        browser.execute_script("arguments[0].click()",close)
                                        for i in range(k,k+10):  
                                            try:
                                                        select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                        browser.execute_script("arguments[0].click()",select)
                                                        sleep(0.5)
                                            except NoSuchElementException:
                                                pass
                                        sleep(10)
                        if res!=0:
                            for i in range(int_part*10+2,int_part*10+2+res):    
                                                try:
                                 
                                                   select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                   essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                   browser.execute_script("arguments[0].click()",select)
                                                   self.essay_titles.append(essay_title.text)
                                                   sleep(0.5)
                                                except NoSuchElementException:
                                                    pass                               
                            download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                            browser.execute_script("arguments[0].click()",download_pdf)
                            download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                            browser.execute_script("arguments[0].click()",download)
                            WebDriverWait(browser, 60).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                            close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                            browser.execute_script("arguments[0].click()",close)
                            for i in range(int_part*10+2,int_part*10+2+res):   
                                                try:
                                                        select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                        browser.execute_script("arguments[0].click()",select)
                                                        sleep(0.5)
                                                except NoSuchElementException:
                                                    pass
                        else:
                            pass  
                        pass
                    else:
                        
                        for k in range(3,53,10):
                                        sleep(3)
                                        for i in range(k,k+10):    
                                            try:
                                                select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                browser.execute_script("arguments[0].click()",select)
                                                self.essay_titles.append(essay_title.text)
                                                sleep(0.5)
                                            except NoSuchElementException:
                                                pass
                                        download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                                        browser.execute_script("arguments[0].click()",download_pdf)
                                        download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                                        browser.execute_script("arguments[0].click()",download)
                                        WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                                        sleep(2)
                                        close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                                        browser.execute_script("arguments[0].click()",close)
                                        for i in range(k,k+10):  
                                            try:
                                                        select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                        browser.execute_script("arguments[0].click()",select)
                                                        sleep(0.5)
                                            except NoSuchElementException:
                                                pass 
                                        sleep(3)
                        journal_nums_in_page2=total_journals-50 
                        if journal_nums_in_page2==50:
                                    page_button=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-paginator/div[2]/ul/li[{2}]/button')
                                    browser.execute_script("arguments[0].click()",page_button)
                                    win=browser.window_handles
                                    browser.switch_to.window(win[-1])
                                    sleep(10)
                                    for k in range(2,52,10):
                                        sleep(3)
                                        for i in range(k,k+10):    
                                            try:
                                                select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                browser.execute_script("arguments[0].click()",select)
                                                self.essay_titles.append(essay_title.text)
                                                sleep(0.5)
                                            except NoSuchElementException:
                                                pass
                                        download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                                        browser.execute_script("arguments[0].click()",download_pdf)
                                        download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                                        browser.execute_script("arguments[0].click()",download)
                                        WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                                        sleep(2)
                                        close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                                        browser.execute_script("arguments[0].click()",close)
                                        for i in range(k,k+10):  
                                            try:
                                                        select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                        browser.execute_script("arguments[0].click()",select)
                                                        sleep(0.5)
                                            except NoSuchElementException:
                                                pass
                                        sleep(10)
                        elif 10<journal_nums_in_page2<50:
                            page_button=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-paginator/div[2]/ul/li[{2}]/button')
                            browser.execute_script("arguments[0].click()",page_button)
                            win=browser.window_handles
                            browser.switch_to.window(win[-1])
                            sleep(10)
                            res=total_journals%10
                            int_part=total_journals//10
                            for k in range(2,int_part*10+2,10):
                                            sleep(3)
                                            for i in range(k,k+10):    
                                                try:
                                                    select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                    essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                    browser.execute_script("arguments[0].click()",select)
                                                    self.essay_titles.append(essay_title.text)
                                                    sleep(0.5)
                                                except NoSuchElementException:
                                                    pass
                                            download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                                            browser.execute_script("arguments[0].click()",download_pdf)
                                            download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                                            browser.execute_script("arguments[0].click()",download)
                                            WebDriverWait(browser, 60).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                                            sleep(2)
                                            close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                                            browser.execute_script("arguments[0].click()",close)
                                            for i in range(k,k+10):  
                                                try:

                                                        select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                        browser.execute_script("arguments[0].click()",select)
                                                        sleep(0.5)
                                                except NoSuchElementException:
                                                    pass
                                            sleep(10)

                            if res!=0:
                                for i in range(int_part*10+2,int_part+2*10+res):    
                                                    try:
                                                        select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                        essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                        browser.execute_script("arguments[0].click()",select)
                                                        self.essay_titles.append(essay_title.text)
                                                        sleep(0.5)
                                                    except NoSuchElementException:
                                                        pass
                              
                                download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                                browser.execute_script("arguments[0].click()",download_pdf)
                                download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                                browser.execute_script("arguments[0].click()",download)
                                WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                                close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                                browser.execute_script("arguments[0].click()",close)
                                for i in range(int_part*10+2,int_part+2*10+res):   
                                               try:
                                                       
                                                        select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                        browser.execute_script("arguments[0].click()",select)
                                                        sleep(0.5)
                                               except NoSuchElementException:
                                                        pass
                            else:
                                pass
                                        
                        elif journal_nums_in_page2<=10: 
                            page_button=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-paginator/div[2]/ul/li[{2}]/button')
                            browser.execute_script("ar guments[0].click()",page_button)
                            sleep(10)
                            for i in range(2,journal_nums_in_page2+2):    
                                                    try:
                                                        select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                        essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                        browser.execute_script('arguments[0].click()',select)
                                                        self.essay_titles.append(essay_title.text)
                                                        sleep(0.5)
                                                    except NoSuchElementException:
                                                        pass
                            sleep(5)
                            download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                            browser.execute_script("arguments[0].click()",download_pdf)
                            download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                            browser.execute_script("arguments[0].click()",download)
                            WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                            close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                            browser.execute_script("arguments[0].click()",close)
                     
                        else:
                             pass    
                    full_path=os.path.join(self.path,'下载论文列表.txt')
                    with open(full_path,'w',encoding='utf-8') as pdf:
                        for ele in self.essay_titles:
                            pdf.write(ele+'\n')
                    cleaned_lines=[essay.replace(" ","") for essay in self.essay_titles]
                    cleaned_lines=[essay.replace("\n","") for essay in cleaned_lines]
                    for essay_title in cleaned_lines:
                        tree.insert("","end",values=(essay_title))
                    for item in tree.get_children():      
                                tree.set(item,"第二列",'已下载')
                    self.essay_titles.clear()
                    cleaned_lines.clear()
                    browser.quit()
                    close_Listening_mode()
                    self.extract_zip_file()
                    if self.radio_flag:
                        set_Volume_to_100()
                        Speak(f'Journal of Lightwave Technology 当月最新论文已为您下载到{self.path}请查看')
                    messagebox.showinfo('提示',f'Journal of Lightwave Technology当月最新论文已为您下载到{self.path}请查看')
        def file():
            messagebox.showinfo("提示","请选择存放所有下载pdf的文件夹")
            folder_path=filedialog.askdirectory(title='选择存放所有下载pdf的文件夹')
            if not os.path.exists(folder_path):
                messagebox.showinfo('注意','请输入有效的地址！')
            folder_path=folder_path.replace('/','\\')
            journal=Journal_of_Lightwave_Technology(folder_path)
            global button1
            combo1=ttk.Combobox(jroot,validate='none')
            label1=tk.Label(jroot,text='选择浏览器模式',font=('宋体',10),fg='red',bg='#d7e8f0')
            label1.place(x=880,y=20)
            combo1['values']=('打开浏览器')
            combo1.current(0)
            combo1.place(x=880,y=40)
            sure_button1=tk.Button(jroot,text="确认",command=lambda:journal.change_browser_flag(combo1.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
            sure_button1.place(x=1150-scaling_factor+65 if scaling_factor!=150 else 1150,y=40)
            combo2=ttk.Combobox(jroot,validate='none')
            label2=tk.Label(jroot,text='选择是否语音播报',font=('宋体',10),fg='red',bg='#d7e8f0')
            label2.place(x=880,y=120)
            combo2['values']=('语音播报','不语音播报')
            combo2.place(x=880,y=140)
            sure_button2=tk.Button(jroot,text="确认",command=lambda:journal.change_radio_flag(combo2.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
            sure_button2.place(x=1150-scaling_factor+65 if scaling_factor!=150 else 1150,y=140)
            combo2.current(0)
            button1=tk.Button(jroot,text="运行脚本",command=journal.journal_of_Lightwave_Technology,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button1.place(x=930,y=250)
            button2=tk.Button(jroot,text="清空文件夹内文件",command=journal.empty_folder,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button2.place(x=930,y=400)
        def helper():
            messagebox.showinfo("说明","\n点击运行脚本后请勿操纵鼠标，请耐心等待自动化操作完成。\n                                        感谢您的使用！") 
        def abouts():
            messagebox.showinfo("版本信息","开发者：Mr Crab\n版本号：8.0")
        menubar=tk.Menu(jroot,tearoff=True)
        jroot.config(menu=menubar)
        menubar.add_command(label="选择存放pdf的文件夹",font=("宋体",15),command=file)
        menubar.add_separator()
        menubar.add_command(label="说明",font=("宋体",15),command=helper)
        menubar.add_separator()
        menubar.add_command(label="关于",font=("宋体",15),command=abouts) 
        



def journal_of_Optical_Communications_and_Networking_interface():      
        jroot=tk.Toplevel()
        jroot.geometry("1250x700")
        jroot.resizable(False,False)
       
        jroot.title("Journal of Optical Communications and Networking")
        jroot.iconphoto(True,Base64_to_Image(IEEE_icon.img,100,100))
        frame=tk.Frame(jroot,bg='#d7e8f0')
        frame.pack(fill="both",expand=True)
        tree=ttk.Treeview(frame,columns=("第一列","第二列"),height=30,show="headings")
        tree.heading("第一列",text="论文名称")
        tree.heading("第二列",text="下载状态")
        tree.column("第一列",width=650)
        tree.column("第二列",width=150)
        tree.grid(row=3,column=3,sticky="nsew")
        scrollbar=ttk.Scrollbar(frame,orient="vertical",command=tree.yview)
        scrollbar.grid(row=3,column=4,sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)
        class Journal_of_Optical_Communications_and_Networking():
            def __init__(self,path):
                 self.path=path
                 self.essay_titles=[]
                 self.browser_flag=0
                 self.radio_flag=1
            def extract_zip_file(self):
                        filelist=os.listdir(self.path)
                        for file in filelist:
                            if file.endswith('.zip'):
                                zip_file_path=os.path.join(self.path,file)
                                with zipfile.ZipFile(zip_file_path,'r') as zip:
                                    zip.extractall(self.path)
            def empty_folder(self):
                if os.listdir(self.path):
                    for filename in os.listdir(self.path):
                        file_path=os.path.join(self.path,filename)
                        os.remove(file_path)
                    messagebox.showinfo('提示',f'已为您清空{self.path}内所有文件')
                else:
                    messagebox.showinfo('提示','该文件夹内没有任何文件，无需清空！')
            def delete_file(self):
                if os.listdir(self.path):
                    for filename in os.listdir(self.path):
                        if filename.endswith('zip'):
                            zip_path=os.path.join(self.path,filename)
                            os.remove(zip_path)
                        else:
                            pass
                        if filename.endswith('.pdf'):
                                pdf_path=os.path.join(self.path,filename)
                                os.remove(pdf_path)
                        else:
                            pass 

                else:
                    pass
            def change_browser_flag(self,option):
                dic={'不打开浏览器':1,'打开浏览器':0}
                self.browser_flag=dic.get(option)
                return self.browser_flag
            def change_radio_flag(self,option):
                dic={'语音播报':1,'不语音播报':0}
                self.radio_flag=dic.get(option)
                return self.radio_flag 
            def journal_of_Optical_Communications_and_Networking(self):
                    open_Listening_mode()
                    button1.config(state='disabled')
                    options=Options()
                    options.add_argument("--disable-blink-features=AutomationControlled")#欺骗网页，隐藏自动化操作的表头
                    options.add_argument('--ignore-ssl-errosr')
                    options.add_argument('--ignore-certificate-errors')
                    options.add_experimental_option('excludeSwitches',['enable-automation'])
                    prefs = {
	                    'download.default_directory': self.path,  # 设置默认下载路径
	                    "profile.default_content_setting_values.automatic_downloads": 1  # 允许多文件下载
                    }
                    options.add_experimental_option("prefs", prefs)
                    if self.browser_flag:
                        options.add_argument('--headless=new')
                        options.add_argument('--disable-gpu')
                        options.add_argument('-window-position=-2400,-2400')
                    browser=webdriver.ChromiumEdge(options)
                    browser.get('https://ieeexplore.ieee.org/Xplore/guesthome.jsp')
                    browser.maximize_window()
                    sleep(5)
                    try:
                        sign=browser.find_element(By.XPATH,'//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/xpl-navbar/div/div[1]/div[3]/div/xpl-institution-details/div/div/div')
                    except NoSuchElementException:
                        WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/xpl-navbar/div/div[1]/div[3]/xpl-login-modal-trigger/a')))
                        institution_log_in_button=browser.find_element(By.XPATH,'//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/xpl-navbar/div/div[1]/div[3]/xpl-login-modal-trigger/a')
                        institution_log_in_button.click()
                        sleep(6)
                        access_through_institution_button=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/xpl-login-modal/div[1]/div[2]/div/div/xpl-login/div/section/div/div/xpl-seamless-access/div/div[1]/button/div/div[2]/div')
                        access_through_institution_button.click()
                        sleep(1)
                        input=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/xpl-login-modal/div[1]/div[2]/div/div/xpl-login/div/section/div/div/div[2]/div[2]/xpl-inst-typeahead/div/div/input')
                        input.send_keys('Hefei University of Technology')
                        sleep(5)
                        pyautogui.hotkey('enter')
                        school=browser.find_element(By.XPATH,'//*[@id="Hefei University of Technology"]')
                        browser.execute_script('arguments[0].click()',school)
                        win=browser.window_handles
                        browser.switch_to.window(win[-1])
                        sleep(5)
                        accounts=str(base64.b64decode(bytes(account_and_pasword.get('account').encode('utf-8'))).decode('utf-8'))
                        passwords=str(base64.b64decode(bytes(account_and_pasword.get('password').encode('utf-8'))).decode('utf-8'))
                        account=browser.find_element(By.XPATH,'//*[@id="username"]')
                        account.click()
                        account.send_keys(accounts)
                        password=browser.find_element(By.XPATH,'//*[@id="pwd"]')
                        password.click()
                        password.send_keys(passwords)
                        login_button=browser.find_element(By.ID,'sb2')
                        login_button.click()
                        win=browser.window_handles
                        browser.switch_to.window(win[-1])
                        sleep(5)
                        accept_button=browser.find_element(By.XPATH,'/html/body/form/div/div[2]/p[2]/input[2]')
                        browser.execute_script('arguments[0].click()',accept_button)
                        win=browser.window_handles
                        browser.switch_to.window(win[-1])
                    sleep(10)
                    input_element=browser.find_element(By.XPATH, '//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/div[2]/div[2]/xpl-search-bar-migr/div/form/div[2]/div/div[1]/xpl-typeahead-migr/div/input')
                    input_element.send_keys('Journal of Optical Communications and Networking')
                    input_element.send_keys(Keys.ENTER)
                    win=browser.window_handles
                    browser.switch_to.window(win[-1])
                    sleep(10)
                    journal_button=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[2]/xpl-suggested-publication-list/div/div[2]/div[1]/a/div')
                    browser.execute_script('arguments[0].click()',journal_button)
                    win=browser.window_handles
                    browser.switch_to.window(win[-1])
                    sleep(10)
                    latest_journal_button=browser.find_element(By.XPATH,'//*[@id="toggle-button-2"]')
                    #latest_journal_button=browser.find_element(By.XPATH,'//*[@id="xplMainContentLandmark"]/div/xpl-xpl-delegate/xpl-journals/div/div[1]/div[3]/section/xpl-journals-home/section/section/section[2]/div[2]/div[1]/div[2]')
                    browser.execute_script('arguments[0].click()',latest_journal_button)
                    sleep(10)  
                    show_all=browser.find_element(By.XPATH,'//*[@id="xplMainContentLandmark"]/div/xpl-xpl-delegate/xpl-journals/div/div[1]/div[3]/section/xpl-journals-home/section/section/section[2]/div[2]/xpl-journal-recent-article/div[2]/a')
                    browser.execute_script("arguments[0].click()",show_all)
                    win=browser.window_handles
                    browser.switch_to.window(win[-1])
                    sleep(15)
                    total_journals1=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[2]/div[1]/div/div/span[1]/span[2]')
                    total_journals=int(total_journals1.text)
                    sleep(3)
                    self.delete_file()
                    if total_journals==25:
                                    for k in range(2,27,10):
                                        sleep(3)
                                        for i in range(k,k+10):    
                                            try:
                                                essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                browser.execute_script("arguments[0].click()",select)
                                                self.essay_titles.append(essay_title.text)
                                                sleep(0.5)
                                            except NoSuchElementException:
                                                pass
                                        download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                                        browser.execute_script("arguments[0].click()",download_pdf)
                                        download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                                        browser.execute_script("arguments[0].click()",download)
                                        WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                                        sleep(2)
                                        close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                                        browser.execute_script("arguments[0].click()",close)
                                        for i in range(k,k+10):  
                                          try:
                                        
                                                select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                browser.execute_script("arguments[0].click()",select)
                                      
                                                sleep(0.5)
                                          except NoSuchElementException:
                                                pass
                                        sleep(10)
                    elif 10<total_journals<25:
        
                        res=total_journals%10
                        int_part=total_journals//10
                        for k in range(2,int_part*10+2,10):
                                
                                        sleep(3)
                                        for i in range(k,k+10):    
                                          try:
                                                essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                browser.execute_script("arguments[0].click()",select)
                                                self.essay_titles.append(essay_title.text)
                                                sleep(0.5)
                                          except NoSuchElementException:
                                                pass
                        
                                        download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                                        browser.execute_script("arguments[0].click()",download_pdf)
                                        download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                                        browser.execute_script("arguments[0].click()",download)
                                        WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                                        sleep(2)
                                        close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                                        browser.execute_script("arguments[0].click()",close)
                                        for i in range(k,k+10):  
                                            try:
                                        
                                                select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                browser.execute_script("arguments[0].click()",select)
                                  
                                                sleep(0.5)
                                            except NoSuchElementException:
                                                pass
                                        sleep(10)
                        if res!=0:
                            for i in range(int_part*10+2,int_part*10+2+res):    
                                          try:
                                                essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                browser.execute_script("arguments[0].click()",select)
                                                self.essay_titles.append(essay_title.text)
                                                sleep(0.5)
                                          except NoSuchElementException:
                                                    pass                               
                            download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                            browser.execute_script("arguments[0].click()",download_pdf)
                            download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                            browser.execute_script("arguments[0].click()",download)
                            WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                            sleep(2)
                            close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                            browser.execute_script("arguments[0].click()",close)
                            for i in range(int_part*10+2,int_part*10+2+res):   
                                              try:
                                        
                                                select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                browser.execute_script("arguments[0].click()",select)
                                      
                                                sleep(0.5)
                                              except NoSuchElementException:
                                                    pass
                        else:
                            pass  
                    elif total_journals>25:
                        for k in range(2,27,10):
                                        sleep(3)
                                        for i in range(k,k+10):    
                                              try:
                                                    essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                    select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                    browser.execute_script("arguments[0].click()",select)
                                                    self.essay_titles.append(essay_title.text)
                                                    sleep(0.5)
                                              except NoSuchElementException:
                                                    pass
                                        download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                                        browser.execute_script("arguments[0].click()",download_pdf)
                                        download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                                        browser.execute_script("arguments[0].click()",download)
                                        WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                                        sleep(2)
                                        close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                                        browser.execute_script("arguments[0].click()",close)
                                        for i in range(k,k+10):  
                                            try:
                                                select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                browser.execute_script("arguments[0].click()",select)
                                                sleep(0.5)
                                            except NoSuchElementException:
                                                pass
                                        sleep(10)
                        journal_nums_in_page2=total_journals-25    
                        page_button=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-paginator/div[2]/ul/li[{2}]/button')
                        browser.execute_script("arguments[0].click()",page_button)
                        win=browser.window_handles
                        browser.switch_to.window(win[-1])
                        sleep(4)
                        if journal_nums_in_page2==25:
                                    for k in range(2,27,10):
                                        sleep(3)
                                        for i in range(k,k+10):    
                                            try:
                                                essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                browser.execute_script("arguments[0].click()",select)
                                                self.essay_titles.append(essay_title.text)
                                                sleep(0.5)
                                            except NoSuchElementException:
                                                pass
                                        download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                                        browser.execute_script("arguments[0].click()",download_pdf)
                                        download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                                        browser.execute_script("arguments[0].click()",download)
                                        WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                                        sleep(2)
                                        close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                                        browser.execute_script("arguments[0].click()",close)
                                        for i in range(k,k+10):  
                                            try:
                                                select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                browser.execute_script("arguments[0].click()",select)
                                                sleep(0.5)
                                            except NoSuchElementException:
                                                pass
                                        sleep(10)
                        elif journal_nums_in_page2<25:
                            res=total_journals%10
                            int_part=total_journals//10
                            for k in range(2,int_part*10+2,10):
                                            sleep(3)
                                            for i in range(k,k+10):    
                                                try:
                                                    essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                    select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                    browser.execute_script("arguments[0].click()",select)
                                                    self.essay_titles.append(essay_title.text)
                                                    sleep(0.5)
                                                except NoSuchElementException:
                                                    pass
                                            download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                                            browser.execute_script("arguments[0].click()",download_pdf)
                                            download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                                            browser.execute_script("arguments[0].click()",download)
                                            WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                                            sleep(2)
                                            close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                                            browser.execute_script("arguments[0].click()",close)
                                            for i in range(k,k+10):  
                                                try:
                                                    select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                    browser.execute_script("arguments[0].click()",select)
                                                    sleep(0.5)
                                                except NoSuchElementException:
                                                    pass
                                            sleep(10)
                            if res!=0:
                                for i in range(int_part*10+2,int_part+2*10+res):    
                                                    try:
                                                        essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                        select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                        browser.execute_script("arguments[0].click()",select)
                                                        self.essay_titles.append(essay_title.text)
                                                        sleep(0.5)
                                                    except NoSuchElementException:
                                                        pass
                                download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                                browser.execute_script("arguments[0].click()",download_pdf)
                                download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                                browser.execute_script("arguments[0].click()",download)
                                WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                                sleep(2)
                                close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                                browser.execute_script("arguments[0].click()",close)                     
                                for i in range(int_part*10+2,int_part+2*10+res):   
                                                    try:
                                                        select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                        browser.execute_script("arguments[0].click()",select)
                                                        sleep(0.5)
                                                    except NoSuchElementException:
                                                        pass

                            else:
                                pass          
                        else:
                            pass
                    else:
                        for i in range(2,total_journals+2):    
                            essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                        
                            self.essay_titles.append(essay_title.text)
                            sleep(0.5)
                        select_all_on_page=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[1]/label/input')
                        browser.execute_script("arguments[0].click()",select_all_on_page)
                        download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                        browser.execute_script("arguments[0].click()",download_pdf)
                        download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                        browser.execute_script("arguments[0].click()",download)
                        WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                        sleep(2)
                        close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                        browser.execute_script("arguments[0].click()",close) 
                        
            
               
                    full_path=os.path.join(self.path,'下载论文列表.txt')
                    with open(full_path,'w',encoding='utf-8') as pdf:
                        for ele in self.essay_titles:
                            pdf.write(ele+'\n')
                    cleaned_lines1=[essay.replace(" ","") for essay in self.essay_titles]
                    cleaned_lines2=[essay.replace("\n","") for essay in cleaned_lines1]
                    for essay_title in cleaned_lines2:
                        tree.insert("","end",values=(essay_title))
                    for item in tree.get_children():      
                                tree.set(item,"第二列",'已下载')
                    self.essay_titles.clear()
                    cleaned_lines1.clear()
                    cleaned_lines2.clear()
                    browser.quit()  
                    close_Listening_mode()
                    self.extract_zip_file()
                    if self.radio_flag:
                        set_Volume_to_100()
                        Speak(f'Journal of Optical Communications and Networking当月最新论文已为您下载到{self.path}请查看')
                    messagebox.showinfo('提示',f'Journal of Optical Communications and Networking当月最新论文已为您下载到{self.path}请查看')
        def file():
            messagebox.showinfo("提示","请选择存放所有下载pdf的文件夹")
            folder_path=filedialog.askdirectory(title='选择存放所有下载pdf的文件夹')
            if not os.path.exists(folder_path):
                messagebox.showinfo('注意','请输入有效的地址！')
            folder_path=folder_path.replace('/','\\')
            journal=Journal_of_Optical_Communications_and_Networking(folder_path)
            global button1
            combo1=ttk.Combobox(jroot,validate='none')
            label1=tk.Label(jroot,text='选择浏览器模式',font=('宋体',10),fg='red',bg='#d7e8f0')
            label1.place(x=880,y=20)
            combo1['values']=('打开浏览器')
            combo1.current(0)
            combo1.place(x=880,y=40)
            sure_button1=tk.Button(jroot,text="确认",command=lambda:journal.change_browser_flag(combo1.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
            sure_button1.place(x=1150-scaling_factor+65 if scaling_factor!=150 else 1150,y=40)
            combo2=ttk.Combobox(jroot,validate='none')
            label2=tk.Label(jroot,text='选择是否语音播报',font=('宋体',10),fg='red',bg='#d7e8f0')
            label2.place(x=880,y=120)
            combo2['values']=('语音播报','不语音播报')
            combo2.place(x=880,y=140)
            sure_button2=tk.Button(jroot,text="确认",command=lambda:journal.change_radio_flag(combo2.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
            sure_button2.place(x=1150-scaling_factor+65 if scaling_factor!=150 else 1150,y=140)
            combo2.current(0)
            button1=tk.Button(jroot,text="运行脚本",command=journal.journal_of_Optical_Communications_and_Networking,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button1.place(x=930,y=250)
            button2=tk.Button(jroot,text="清空文件夹内文件",command=journal.empty_folder,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button2.place(x=930,y=400)
        def helper():
            messagebox.showinfo("说明","\n点击运行脚本后请勿操纵鼠标，请耐心等待自动化操作完成。                                      \n 感谢您的使用！")
            
        def abouts():
            messagebox.showinfo("版本信息","开发者：Mr Crab\n版本号：8.0")
        menubar=tk.Menu(jroot,tearoff=True)
        menubar.add_command(label="选择存放pdf的文件夹",font=("宋体",15),command=file)
        menubar.add_separator()
        menubar.add_command(label="说明",font=("宋体",15),command=helper)
        menubar.add_separator()
        menubar.add_command(label="关于",font=("宋体",15),command=abouts)
        jroot.config(menu=menubar)
        
def IEEE_Photonics_Technology_Letters_interface():      
        jroot=tk.Toplevel()
        jroot.geometry("1250x700")
        jroot.resizable(False,False)
        jroot.title("IEEE Photonics Technology Letters")
        jroot.iconphoto(True,Base64_to_Image(IEEE_icon.img,100,100))
        frame=tk.Frame(jroot,bg='#d7e8f0')
        frame.pack(fill="both",expand=True)
        tree=ttk.Treeview(frame,columns=("第一列","第二列"),height=30,show="headings")
        tree.heading("第一列",text="论文名称")
        tree.heading("第二列",text="下载状态")
        tree.column("第一列",width=650)
        tree.column("第二列",width=150)
        tree.grid(row=3,column=3,sticky="nsew")
        scrollbar=ttk.Scrollbar(frame,orient="vertical",command=tree.yview)
        scrollbar.grid(row=3,column=4,sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)
        class IEEE_Photonics_Technology_Letters():
            def __init__(self,path):
                  self.path=path
                  self.essay_titles=[]
                  self.browser_flag=0
                  self.radio_flag=1
            def extract_zip_file(self):
                        filelist=os.listdir(self.path)
                        for file in filelist:
                            if file.endswith('.zip'):
                                zip_file_path=os.path.join(self.path,file)
                                with zipfile.ZipFile(zip_file_path,'r') as zip:
                                    zip.extractall(self.path)
            def empty_folder(self):
                if os.listdir(self.path):
                    for filename in os.listdir(self.path):
                        file_path=os.path.join(self.path,filename)
                        os.remove(file_path)
                    messagebox.showinfo('提示',f'已为您清空{self.path}内所有文件')
                else:
                    messagebox.showinfo('提示','该文件夹内没有任何文件，无需清空！')
            def delete_file(self):
                if os.listdir(self.path):
                    for filename in os.listdir(self.path):
                        if filename.endswith('zip'):
                            zip_path=os.path.join(self.path,filename)
                            os.remove(zip_path)
                        else:
                            pass
                        if filename.endswith('.pdf'):
                                pdf_path=os.path.join(self.path,filename)
                                os.remove(pdf_path)
                        else:
                            pass 

                else:
                    pass
            def change_browser_flag(self,option):
                dic={'不打开浏览器':1,'打开浏览器':0}
                self.browser_flag=dic.get(option)
                return self.browser_flag
            def change_radio_flag(self,option):
                dic={'语音播报':1,'不语音播报':0}
                self.radio_flag=dic.get(option)
                return self.radio_flag 
            def ieee_Photonics_Technology_Letters(self):
                    open_Listening_mode()
                    button1.config(state='disabled')
                    options=Options()
                    options.add_argument("--disable-blink-features=AutomationControlled")#欺骗网页，隐藏自动化操作的表头
                    options.add_argument('--ignore-ssl-errosr')
                    options.add_argument('--ignore-certificate-errors')
                    options.add_experimental_option('excludeSwitches',['enable-automation'])
                    prefs = {
	                    'download.default_directory': self.path,  # 设置默认下载路径
	                    "profile.default_content_setting_values.automatic_downloads": 1  # 允许多文件下载
                    }
                    options.add_experimental_option("prefs", prefs)
                    if self.browser_flag:
                        options.add_argument("-headless=new")
                        options.add_argument("--disable-gpu")
                        options.add_argument('--window-position=-2400,-2400')
                    browser=webdriver.ChromiumEdge(options)
                    browser.get('https://ieeexplore.ieee.org/Xplore/guesthome.jsp')
                    browser.maximize_window()
                    sleep(5)
                    try:
                        sign=browser.find_element(By.XPATH,'//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/xpl-navbar/div/div[1]/div[3]/div/xpl-institution-details/div/div/div')
                    except NoSuchElementException:
                        WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/xpl-navbar/div/div[1]/div[3]/xpl-login-modal-trigger/a')))
                        institution_log_in_button=browser.find_element(By.XPATH,'//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/xpl-navbar/div/div[1]/div[3]/xpl-login-modal-trigger/a')
                        institution_log_in_button.click()
                        sleep(6)
                        access_through_institution_button=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/xpl-login-modal/div[1]/div[2]/div/div/xpl-login/div/section/div/div/xpl-seamless-access/div/div[1]/button/div/div[2]/div')
                        access_through_institution_button.click()
                        sleep(1)
                        input=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/xpl-login-modal/div[1]/div[2]/div/div/xpl-login/div/section/div/div/div[2]/div[2]/xpl-inst-typeahead/div/div/input')
                        input.send_keys('Hefei University of Technology')
                        pyautogui.hotkey('enter')
                        sleep(5)
                        school=browser.find_element(By.XPATH,'//*[@id="Hefei University of Technology"]')
                        browser.execute_script('arguments[0].click()',school)
                        win=browser.window_handles
                        browser.switch_to.window(win[-1])
                        sleep(5)
                        accounts=str(base64.b64decode(bytes(account_and_pasword.get('account').encode('utf-8'))).decode('utf-8'))
                        passwords=str(base64.b64decode(bytes(account_and_pasword.get('password').encode('utf-8'))).decode('utf-8'))
                        account=browser.find_element(By.XPATH,'//*[@id="username"]')
                        account.click()
                        account.send_keys(accounts)
                        password=browser.find_element(By.XPATH,'//*[@id="pwd"]')
                        password.click()
                        password.send_keys(passwords)
                        login_button=browser.find_element(By.ID,'sb2')
                        login_button.click()
                        win=browser.window_handles
                        browser.switch_to.window(win[-1])
                        sleep(5)
                        accept_button=browser.find_element(By.XPATH,'/html/body/form/div/div[2]/p[2]/input[2]')
                        browser.execute_script('arguments[0].click()',accept_button)
                        win=browser.window_handles
                        browser.switch_to.window(win[-1])
                    sleep(10)
                    input_element=browser.find_element(By.XPATH, '//*[@id="LayoutWrapper"]/div/div/div[3]/div/xpl-root/header/xpl-header/div/div[2]/div[2]/xpl-search-bar-migr/div/form/div[2]/div/div[1]/xpl-typeahead-migr/div/input')
                    sleep(4)
                    input_element.send_keys('IEEE Photonics Technology Letters')
                    input_element.send_keys(Keys.ENTER)
                    sleep(10)
                    journal_button=browser.find_element(By.XPATH,'//*[@id="xplMainContent"]/div[2]/xpl-suggested-publication-list/div/div[2]/div[1]/a/div')
                    browser.execute_script("arguments[0].click()",journal_button)
                    win=browser.window_handles
                    browser.switch_to.window(win[-1])
                    sleep(10)
                    latest_journal_button=browser.find_element(By.XPATH,'//*[@id="toggle-button-2"]')
                    #latest_journal_button=browser.find_element(By.XPATH,'//*[@id="xplMainContentLandmark"]/div/xpl-xpl-delegate/xpl-journals/div/div[1]/div[3]/section/xpl-journals-home/section/section/section[2]/div[2]/div[1]/div[2]')
                    browser.execute_script("arguments[0].click()",latest_journal_button)
                    sleep(10)
                    show_all=browser.find_element(By.XPATH,'//*[@id="xplMainContentLandmark"]/div/xpl-xpl-delegate/xpl-journals/div/div[1]/div[3]/section/xpl-journals-home/section/section/section[2]/div[2]/xpl-journal-recent-article/div[2]/a')
                    browser.execute_script("arguments[0].click()",show_all)
                    win=browser.window_handles
                    browser.switch_to.window(win[-1])
                    sleep(15)
                    total_journals1=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[2]/div[1]/div/div/span[1]/span[2]')
                    total_journals=int(total_journals1.text)
                    sleep(3)
                    self.delete_file()
                    if total_journals==25:
                                    for k in range(2,27,10):
                                        sleep(3)
                                        for i in range(k,k+10):    
                                            try:  
                                                        select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                        essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                        browser.execute_script("arguments[0].click()",select)
                                                        self.essay_titles.append(essay_title.text)
                                            except NoSuchElementException:
                                                pass
                                        download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                                        browser.execute_script("arguments[0].click()",download_pdf)
                                        download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                                        browser.execute_script("arguments[0].click()",download)
                                        WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                                        sleep(2)
                                        close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                                        browser.execute_script("arguments[0].click()",close)
                                        for i in range(k,k+10):  
                                            try:
                                                select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                browser.execute_script("arguments[0].click()",select)
                                                sleep(0.5)
                                            except NoSuchElementException:
                                                pass
                    elif 10<total_journals<25:
        
                        res=total_journals%10
                        int_part=total_journals//10
                        for k in range(2,int_part*10+2,10):
                                        sleep(3)
                                        for i in range(k,k+10):    
                                            try:  
                                                        select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                        essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                        browser.execute_script("arguments[0].click()",select)
                                                        self.essay_titles.append(essay_title.text)
                                            except NoSuchElementException:
                                                pass
                        
                                        download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                                        browser.execute_script("arguments[0].click()",download_pdf)
                                        download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                                        browser.execute_script("arguments[0].click()",download)
                                        WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                                        sleep(2)
                                        close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                                        browser.execute_script("arguments[0].click()",close)
                                        for i in range(k,k+10):  
                                            try:
                                                select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                browser.execute_script("arguments[0].click()",select)
                                                sleep(0.5)
                                            except NoSuchElementException:
                                                pass
                        if res!=0:
                            for i in range(int_part*10+2,int_part*10+2+res):    
                                                try:  
                                                        select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                        essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                        browser.execute_script("arguments[0].click()",select)
                                                        self.essay_titles.append(essay_title.text)
                                                except NoSuchElementException:
                                                    pass                               
                            download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                            browser.execute_script("arguments[0].click()",download_pdf)
                            download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                            browser.execute_script("arguments[0].click()",download)
                            WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                            sleep(2)
                            close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                            browser.execute_script("arguments[0].click()",close)
                            for i in range(int_part*10+2,int_part*10+2+res):   
                                                try:
                                                    select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                    browser.execute_script("arguments[0].click()",select)
                                                    sleep(0.5)
                                                except NoSuchElementException:
                                                    pass
                        else:
                            pass  
                    elif total_journals>25:
                        for k in range(2,27,10):
                                        sleep(3)
                                        for i in range(k,k+10):    
                                            try:  
                                                        select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                        essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                        browser.execute_script("arguments[0].click()",select)
                                                        self.essay_titles.append(essay_title.text)
                                            except NoSuchElementException:
                                                pass
                                        download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                                        browser.execute_script("arguments[0].click()",download_pdf)
                                        download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                                        browser.execute_script("arguments[0].click()",download)
                                        WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                                        sleep(2)
                                        close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                                        browser.execute_script("arguments[0].click()",close)
                                        for i in range(k,k+10):  
                                            try:
                                                select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                browser.execute_script("arguments[0].click()",select)
                                                sleep(0.5)
                                            except NoSuchElementException:
                                                pass    
                        journal_nums_in_page2=total_journals-25    
                        page_button=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-paginator/div[2]/ul/li[{2}]/button')
                        browser.execute_script("arguments[0].click()",page_button)
                        win=browser.window_handles
                        browser.switch_to.window(win[-1])
                        sleep(4)
                        if journal_nums_in_page2==25:
                                    for k in range(2,27,10):
                                        sleep(3)
                                        for i in range(k,k+10):     
                                            try:  
                                                        select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                        essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                        browser.execute_script("arguments[0].click()",select)
                                                        self.essay_titles.append(essay_title.text)
                                            except NoSuchElementException:
                                                pass
                                        download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                                        browser.execute_script("arguments[0].click()",download_pdf)
                                        download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                                        browser.execute_script("arguments[0].click()",download)
                                        WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                                        sleep(2)
                                        close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                                        browser.execute_script("arguments[0].click()",close)
                                        for i in range(k,k+10):  
                                            try:
                                                select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                browser.execute_script("arguments[0].click()",select)
                                                sleep(0.5)
                                            except NoSuchElementException:
                                                pass
                        elif journal_nums_in_page2<25:
                            res=total_journals%10
                            int_part=total_journals//10
                            for k in range(2,int_part*10+2,10):
                                            sleep(3)
                                            for i in range(k,k+10):    
                                                try:  
                                                        select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                        essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                        browser.execute_script("arguments[0].click()",select)
                                                        self.essay_titles.append(essay_title.text)
                                                except NoSuchElementException:
                                                    pass
                                            download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                                            browser.execute_script("arguments[0].click()",download_pdf)
                                            download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                                            browser.execute_script("arguments[0].click()",download)
                                            WebDriverWait(browser, 60).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                                            sleep(2)
                                            close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                                            browser.execute_script("arguments[0].click()",close)
                                            for i in range(k,k+10):  
                                                try:
                                                    select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                    browser.execute_script("arguments[0].click()",select)
                                                    sleep(0.5)
                                                except NoSuchElementException:
                                                    pass
                            if res!=0:
                                for i in range(int_part*10+2,int_part+2*10+res):    
                                                    try:  
                                                        select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                        essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                                                        browser.execute_script("arguments[0].click()",select)
                                                        self.essay_titles.append(essay_title.text)
                                                    except NoSuchElementException:
                                                        pass
                                download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                                browser.execute_script("arguments[0].click()",download_pdf)
                                download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                                browser.execute_script("arguments[0].click()",download)
                                WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                                sleep(2)
                                close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                                browser.execute_script("arguments[0].click()",close)
                                for i in range(int_part*10+2,int_part+2*10+res):   
                                                    try:
                                                        select=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[1]/input')
                                                        browser.execute_script("arguments[0].click()",select)
                                                        sleep(0.5)
                                                    except NoSuchElementException:
                                                        pass
                            else:
                                pass          
                        else:
                             pass
                    else:
                        for i in range(2,total_journals+2):
                            essay_title=browser.find_element(By.XPATH,f'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[{i}]/div/xpl-issue-results-items/div[1]/div[1]/div[2]/h2/a')
                            self.essay_titles.append(essay_title.text)
                            sleep(0.5)
                        select_all_on_page=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[2]/div/div[2]/div/xpl-issue-results-list/div[2]/div[1]/label/input')
                        browser.execute_script("arguments[0].click()",select_all_on_page)
                        download_pdf=browser.find_element(By.XPATH,'//*[@id="publicationIssueMainContent global-margin-px"]/div[1]/xpl-issue-search-dashboard/div/div[1]/ul/div/div[1]/xpl-download-pdf/button')
                        browser.execute_script("arguments[0].click()",download_pdf)
                        download=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/section[2]/div/button/span')
                        browser.execute_script("arguments[0].click()",download)
                        WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/ngb-modal-window/div/div/div/div/div[1]')))
                        sleep(2)
                        close=browser.find_element(By.XPATH,'/html/body/ngb-modal-window/div/div/div/div/div[1]/i')
                        browser.execute_script("arguments[0].click()",close) 
                    full_path=os.path.join(self.path,'下载论文列表.txt')
                    with open(full_path,'w',encoding='utf-8') as pdf:
                        for ele in self.essay_titles:
                            pdf.write(ele+'\n')
                    cleaned_lines1=[essay.replace(" ","") for essay in self.essay_titles]
                    cleaned_lines2=[essay.replace("\n","") for essay in cleaned_lines1]
                    for essay_title in cleaned_lines2:
                        tree.insert("","end",values=(essay_title))
                    for item in tree.get_children():      
                                tree.set(item,"第二列",'已下载')
                    self.essay_titles.clear()
                    cleaned_lines1.clear()
                    cleaned_lines2.clear()
                    browser.quit()
                    close_Listening_mode()
                    self.extract_zip_file()
                    if self.radio_flag:
                        set_Volume_to_100()
                        Speak(f'IEEE Photonics Technology Letters当月最新论文已为您下载到{self.path}请查看')
                    messagebox.showinfo('提示',f'IEEE Photonics Technology Letters当月最新论文已为您下载到{self.path}请查看')
        def file():
            messagebox.showinfo("提示","请选择存放所有下载pdf的文件夹")
            folder_path=filedialog.askdirectory(title='选择存放所有下载pdf的文件夹')
            if not os.path.exists(folder_path):
                messagebox.showinfo('注意','请输入有效的地址')
            folder_path=folder_path.replace('/','\\')
            journal=IEEE_Photonics_Technology_Letters(folder_path)
            global button1
            combo1=ttk.Combobox(jroot,validate='none')
            label1=tk.Label(jroot,text='选择浏览器模式',font=('宋体',10),fg='red',bg='#d7e8f0')
            label1.place(x=880,y=20)
            combo1['values']=('打开浏览器')
            combo1.current(0)
            combo1.place(x=880,y=40)
            sure_button1=tk.Button(jroot,text="确认",command=lambda:journal.change_browser_flag(combo1.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
            sure_button1.place(x=1150-scaling_factor+65 if scaling_factor!=150 else 1150,y=40)
            combo2=ttk.Combobox(jroot,validate='none')
            label2=tk.Label(jroot,text='选择是否语音播报',font=('宋体',10),fg='red',bg='#d7e8f0')
            label2.place(x=880,y=120)
            combo2['values']=('语音播报','不语音播报')
            combo2.place(x=880,y=140)
            sure_button2=tk.Button(jroot,text="确认",command=lambda:journal.change_radio_flag(combo2.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
            sure_button2.place(x=1150-scaling_factor+65 if scaling_factor!=150 else 1150,y=140)
            combo2.current(0)
            button1=tk.Button(jroot,text="运行脚本",command=journal.ieee_Photonics_Technology_Letters,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button1.place(x=930,y=250)
            button2=tk.Button(jroot,text="清空文件夹内文件",command=journal.empty_folder,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button2.place(x=930,y=400)
        def helper():
            messagebox.showinfo("说明","\n点击运行脚本后请勿操纵鼠标，请耐心等待自动化操作完成。\n                                        感谢您的使用！")
            
        def abouts():
            messagebox.showinfo("版本信息","开发者：Mr Crab\n版本号：8.0")
        menubar=tk.Menu(jroot,tearoff=True)
        jroot.config(menu=menubar)
        menubar.add_command(label="选择存放pdf的文件夹",font=("宋体",15),command=file)
        menubar.add_separator()
        menubar.add_command(label="说明",font=("宋体",15),command=helper)
        menubar.add_separator()
        menubar.add_command(label="关于",font=("宋体",15),command=abouts) 
def under_developing():
     Speak('该项功能正在开发中') 
     messagebox.showinfo('提示','该项功能正在开发中')
     
def optic_express_interface():
        messagebox.showinfo('注意','请提前备份本地下载位置中\n除optic_express外所有其他pdf至其他位置')    
        oroot=tk.Toplevel()
        oroot.geometry("1050x700")
        oroot.resizable(False,False)
        oroot.title("Optics Express")
        oroot.iconphoto(True,Base64_to_Image(LightScience_icon.img,100,100))
        frame=tk.Frame(oroot,bg='#d7e8f0')
        frame.pack(fill="both",expand=True)
        tree=ttk.Treeview(frame,columns=("第一列","第二列"),height=30,show="headings")
        tree.heading("第一列",text="pdf名称")
        tree.heading("第二列",text="下载状态")
        tree.column("第一列",width=400)
        tree.column("第二列",width=350)
        tree.grid(row=3,column=3,sticky="nsew")
        scrollbar=ttk.Scrollbar(frame,orient="vertical",command=tree.yview)
        scrollbar.grid(row=3,column=4,sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)
        class optic_express():
            def __init__(self,downloads_path,target_path):
                 self.downloads_path=downloads_path
                 self.target_path=target_path
                 self.pdf_names=[]
            def empty_folder(self):
                if os.listdir(self.target_path):
                    for filename in os.listdir(self.target_path):
                        file_path=os.path.join(self.target_path,filename)
                        os.remove(file_path)
                    messagebox.showinfo('提示',f'已为您清空{self.target_path}内所有文件')
                else:
                    messagebox.showinfo('提示','该文件夹内没有任何文件，无需清空！')
            def identify_captcha(self,file):
                ocr=ddddocr.DdddOcr(show_ad=False,beta=True)
                answer=ocr.classification(file)
                return answer   
            def random_time(self):
                return random.randint(120,130) 
            def move(self):
                for root,dirs,files in os.walk(self.downloads_path):
                    for file in files:  
                      if file.endswith('pdf'):
                           source_file_path=os.path.join(root,file)
                           target_file_path=os.path.join(self.target_path,file)
                           shutil.move(source_file_path,target_file_path)

            def scraping_pdf_names(self,url):
                cookies = {
                    'cfid': 'b1f4894a-6180-43e4-ac30-6496a3fe7b83',
                    'cftoken': '0',
                    'COOKIECONSENT': 'yes',
                    '_gid': 'GA1.2.2008189368.1721529395',
                    '_ga_YB0MS6VTVF': 'GS1.1.1721632504.1.1.1721632806.0.0.0',
                    '_ga_NSD664KQBT': 'GS1.1.1721745948.1.0.1721745955.53.0.0',
                    '_ga_EGFN5DM8HB': 'GS1.1.1721823983.36.1.1721826740.58.0.0',
                    '_ga': 'GA1.1.161727879.1720530808',
                }

                headers = {
                    'Accept': '*/*',
                    'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
                    'Cache-Control': 'no-cache',
                    'Connection': 'keep-alive',
                    'Cookie': 'cfid=b1f4894a-6180-43e4-ac30-6496a3fe7b83; cftoken=0; COOKIECONSENT=yes; _gid=GA1.2.2008189368.1721529395; _ga_YB0MS6VTVF=GS1.1.1721632504.1.1.1721632806.0.0.0; _ga_NSD664KQBT=GS1.1.1721745948.1.0.1721745955.53.0.0; _ga_EGFN5DM8HB=GS1.1.1721823983.36.1.1721826740.58.0.0; _ga=GA1.1.161727879.1720530808',
                    'Pragma': 'no-cache',
                    'Referer': 'https://opg.optica.org/oe/issue.cfm?volume=32&issue=15&tocid=1153182',
                    'Sec-Fetch-Dest': 'empty',
                    'Sec-Fetch-Mode': 'cors',
                    'Sec-Fetch-Site': 'same-origin',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0',
                    'X-Requested-With': 'XMLHttpRequest',
                    'sec-ch-ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Microsoft Edge";v="126"',
                    'sec-ch-ua-mobile': '?0',
                    'sec-ch-ua-platform': '"Windows"',
                }

                params = {
                    'jrnid': '4',
                    'volid': '32',
                    'issid': '15',
                    'tocid': '1153182',
                }
                response = requests.get(url, params=params, cookies=cookies, headers=headers)  
                soup = BeautifulSoup(response.text, 'html.parser')
                divs = soup.find_all('div')
                for div in divs:
                    links = div.find_all('a')
                    for link in links:
                        href = link.get('href')
                        if "viewmedia" in str(href) and 'seq=0' in str(href) and 'oe' in str(href) and 'html' not in str(href) :
                            pdf_name=href[-20:-6]
                            if "=" not in pdf_name:
                                self.pdf_names.append(pdf_name)
                self.pdf_names=list(set(self.pdf_names))
               
                folder_path=p=os.path.join(self.target_path,'pdf_names.txt')
                with open(folder_path,'w',encoding='utf-8') as f:
                      for pdf_name in self.pdf_names:
                          f.write(pdf_name+'\n')
                for pdf_name in self.pdf_names:
                        tree.insert("","end",values=(pdf_name))
            
            def download_Optical_Communications_and_Interconnects(self):
                button1.config(state='disabled')
                open_Listening_mode()
                edgeOption=Options()
                # edgeOption.add_argument(f'user-agent={self.User()}')
                #禁止当前网页内一切js脚本的运行 
                prefs = {'profile.default_content_setting_values.javascript': 2} 
                edgeOption.add_argument("--disable-blink-features=AutomationControlled")#伪装网页，隐藏自动化操作的标头
                edgeOption.add_experimental_option('excludeSwitches',['enable-automation'])#伪装网页，隐藏自动化操作的标头
                edgeOption.add_experimental_option('prefs', prefs)
                browser=webdriver.Edge(edgeOption)
                browser.get('https://opg.optica.org/oe/issue.cfm?volume=32&issue=20&tocid=1166387')
                browser.maximize_window()
                browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
              "source": """
                Object.defineProperty(navigator, 'webdriver', {
                  get: () => undefined
                })
              """})
                

                main_window=browser.current_window_handle
                self.scraping_pdf_names(browser.current_url)
                browser.implicitly_wait(10)

                pdf=browser.find_element(By.XPATH,f'//*[@id="accordion-current"]/div[1]/div/div[2]/div/div/div[2]/p[1]/a')
                browser.execute_script('arguments[0].click()',pdf)
    
                win=browser.window_handles
                browser.switch_to.window(win[-1])
                get_pdf_button=browser.find_element(By.XPATH,'//*[@id="articleContainer"]/div[1]/div[3]/div/div[2]/div/div/ul/li[1]/a')  
                browser.execute_script('arguments[0].click()',get_pdf_button)
                win=browser.window_handles
                browser.switch_to.window(win[-1])
                
                sleep(8)
                captcha=browser.find_element(By.XPATH,'/html/body/div/div/form/span[1]')
                img=captcha.screenshot_as_png
                sleep(3)
                answer=self.identify_captcha(img)
                input_element=browser.find_element(By.ID,'Answer')
                input_element.clear()
                input_element.send_keys(answer)
                input_element.send_keys(Keys.ENTER)
               
                ################################################################

                #如果第一次验证码失败，第二次再验证,验证失败的主要原因是速度太快
                try:
                    WebDriverWait(browser,8).until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div/form/span[4]')))#查询是否有输入错误的提示符
                    #有输入错误的提示符直接删除掉原来的验证码，再识别一次
        
                    sleep(8)
                    captcha=browser.find_element(By.XPATH,'/html/body/div/div/form/span[1]').screenshot_as_png
                    answer=self.identify_captcha(captcha)
                    sleep(4)
                    input_element=browser.find_element(By.XPATH,'//*[@id="Answer"]')
                    input_element.clear()
                    input_element.send_keys(answer)
                    input_element.send_keys(Keys.ENTER)
                    #第二次识别不可能出错,除非速度太快
                    sleep(7)
                    pyautogui.hotkey('ctrl','s')
                    sleep(2)
                    pyautogui.hotkey('alt','s')
                    sleep(20)
                    self.move()
                    browser.switch_to.window(main_window)
                    browser.back()
                except TimeoutException:
                    #没有错误的提示符，说明正确页面会跳转到阅览论文的页面，直接ctrl+s
                    sleep(7)
                    pyautogui.hotkey('ctrl','s')
                    sleep(2)
                    pyautogui.hotkey('alt','s')
                    sleep(20)
                    self.move()
                    browser.switch_to.window(main_window)
                    browser.back()
              
                ##################################################################
                for i in range(3,len(self.pdf_names)+2): 
                                
                                browser.implicitly_wait(10)
                                essay_title=f'//*[@id="accordion-current"]/div[1]/div/div[{i}]/div/div/div[2]/p[1]/a'
                                pdf=browser.find_element(By.XPATH,essay_title)
                                browser.execute_script('arguments[0].click()',pdf)
                        
                                #处理封ip
                                try:
                                    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/div/div/p')))#监测是否有出现封ip的请款
                                    sleep(self.random_time())#有封ip的情况随机等待2-3分种再刷新页面就可以了
                                    browser.refresh()
                                    sleep(10)
                                    get_pdf_button=browser.find_element(By.XPATH,'//*[@id="articleContainer"]/div[1]/div[3]/div/div[2]/div/div/ul/li[1]/a')  
                                    browser.execute_script('arguments[0].click()',get_pdf_button)
                                    sleep(5)
                                    win=browser.window_handles
                                    browser.switch_to.window(win[-1])
                                    sleep(5)
                                    captcha=browser.find_element(By.XPATH,'/html/body/div/div/form/span[1]').screenshot_as_png
                                    answer=self.identify_captcha(captcha)
                                    sleep(4)
                                    input_element=browser.find_element(By.XPATH,'//*[@id="Answer"]')
                                    input_element.clear()
                                    input_element.send_keys(answer)
                                    input_element.send_keys(Keys.ENTER)
                                    #监测是否有验证码输入错误的情况
                                    try:
                                        WebDriverWait(browser,8).until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div/form/span[4]')))
                                        sleep(8)
                                        captcha=browser.find_element(By.XPATH,'/html/body/div/div/form/span[1]').screenshot_as_png
                                        answer=self.identify_captcha(captcha)
                                        sleep(4)
                                        input_element=browser.find_element(By.XPATH,'//*[@id="Answer"]')
                                        input_element.clear()
                                        input_element.send_keys(answer)
                                        input_element.send_keys(Keys.ENTER)
                                        sleep(10)
                                        pyautogui.hotkey('ctrl','s')
                                        sleep(2)
                                        pyautogui.hotkey('alt','s')
                                        sleep(20)
                                        self.move()
                                        browser.switch_to.window(main_window)
                                        browser.back()
                                    except TimeoutException:
                                        sleep(7)
                                        pyautogui.hotkey('ctrl','s')
                                        sleep(2)
                                        pyautogui.hotkey('alt','s')
                                        sleep(20)
                                        self.move()
                                        browser.switch_to.window(main_window)
                                        browser.back()
                                except TimeoutException:
                                    get_pdf_button=browser.find_element(By.XPATH,'//*[@id="articleContainer"]/div[1]/div[3]/div/div[2]/div/div/ul/li[1]/a')  
                                    browser.execute_script('arguments[0].click()',get_pdf_button)
                                    win=browser.window_handles
                                    browser.switch_to.window(win[-1])
                                   
                                    sleep(8)
                                    captcha=browser.find_element(By.XPATH,'/html/body/div/div/form/span[1]').screenshot_as_png
                                    answer=self.identify_captcha(captcha)
                                    sleep(4)
                                    input_element=browser.find_element(By.XPATH,'//*[@id="Answer"]')
                                    input_element.clear()
                                    input_element.send_keys(answer)
                                    input_element.send_keys(Keys.ENTER)
                                    try:
                                        WebDriverWait(browser,8).until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div/form/span[4]')))
                                        sleep(8)
                                        captcha=browser.find_element(By.XPATH,'/html/body/div/div/form/span[1]').screenshot_as_png
                                        answer=self.identify_captcha(captcha)
                                        sleep(4)
                                        input_element=browser.find_element(By.XPATH,'//*[@id="Answer"]')
                                        input_element.clear()
                                        input_element.send_keys(answer)
                                        
                                        input_element.send_keys(Keys.ENTER)
                                        sleep(7)
                                        pyautogui.hotkey('ctrl','s')
                                        sleep(2)
                                        pyautogui.hotkey('alt','s')
                                        sleep(20)
                                        self.move()
                                        browser.switch_to.window(main_window)
                                        browser.back()
                                    except TimeoutException:
                                        sleep(7)
                                        pyautogui.hotkey('ctrl','s')
                                        sleep(2)
                                        pyautogui.hotkey('alt','s')
                                        sleep(20)
                                        self.move()
                                        browser.switch_to.window(main_window)
                                        browser.back()
                sleep(35)
                self.move()
                browser.quit()
                close_Listening_mode()
                for item in tree.get_children():
                    tree.set(item,'第二列','已下载')  
                self.pdf_names.clear()
                Speak(f"Optical Communications and Interconnects最新论文已经为您下载到{self.target_path}，请查看")
                messagebox.showinfo('提示',f'Optical Communications and Interconnects最新论文已经为您下载到{self.target_path}，请查看')
        def file():
            messagebox.showinfo("提示","请选择存放所有下载pdf的文件夹")
            folder_path=filedialog.askdirectory(title='选择存放所有下载pdf的文件夹')
            messagebox.showinfo("提示","请点击选中电脑本地下载downloads")
            downloads_path=filedialog.askdirectory(title='点击选中电脑本地下载downloads')
            if not os.path.exists(folder_path) or not os.path.exists(downloads_path):
                messagebox.showinfo('注意','请输入有效的地址！')
            downloads_path=downloads_path.replace('/','\\')
            folder_path=folder_path.replace('/','\\')
            optic=optic_express(downloads_path=downloads_path,target_path=folder_path)
            global button1
            button1=tk.Button(oroot,text="开始下载",command=optic.download_Optical_Communications_and_Interconnects,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button1.place(x=830,y=100)
            button2=tk.Button(oroot,text="清空文件夹内文件",command=optic.empty_folder,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button2.place(x=830,y=350)
        def helper():
                messagebox.showinfo("说明","\n点击运行脚本后请勿操纵鼠标，请耐心等待自动化操作完成。\n                                        感谢您的使用！")
        def abouts():
                messagebox.showinfo("版本信息","开发者：Mr Crab\n版本号：8.0")
       
        menubar=tk.Menu(oroot,tearoff=True)
        oroot.config(menu=menubar)
        menubar.add_command(label="选择存放pdf的文件夹\n及本地下载的位置",font=("宋体",15),command=file)
        menubar.add_separator()
        menubar.add_command(label="说明",font=("宋体",15),command=helper)
        menubar.add_separator()
        menubar.add_command(label="关于",font=("宋体",15),command=abouts)
def 中国科学_interface():
        zroot=tk.Toplevel()
        zroot.geometry("1200x700")
        zroot.resizable(False,False)
        zroot.title("中国科学：")
      
        zroot.iconphoto(True,Base64_to_Image(中国科学_icon.img,100,100))
        frame=tk.Frame(zroot,bg='#d7e8f0')
        frame.pack(fill="both",expand=True)
        tree=ttk.Treeview(frame,columns=("第一列","第二列"),height=30,show="headings")
        tree.heading("第一列",text="论文名称")
        tree.heading("第二列",text="下载状态")
        tree.column("第一列",width=400)
        tree.column("第二列",width=400)
        tree.grid(row=3,column=3,sticky="nsew")
        scrollbar=ttk.Scrollbar(frame,orient="vertical",command=tree.yview)
        scrollbar.grid(row=3,column=4,sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)
        class 中国科学():
            def __init__(self,download_path):
                self.download_path=download_path
                self.essay_titles=[]
                self.browser_flag=0
                self.radio_flag=1
            def empty_folder(self):
                if os.listdir(self.download_path):
                    for filename in os.listdir(self.download_path):
                        file_path=os.path.join(self.download_path,filename)
                        os.remove(file_path)
                    messagebox.showinfo('提示',f'已为您清空{self.download_path}内所有文件')
                else:
                    messagebox.showinfo('提示','该文件夹内没有任何文件，无需清空！')
            def delete_file(self):
                if os.listdir(self.download_path):
                    for filename in os.listdir(self.download_path):
                        if filename.endswith('zip'):
                            zip_path=os.path.join(self.download_path,filename)
                            os.remove(zip_path)
                        else:
                            pass
                        if filename.endswith('.pdf'):
                                pdf_path=os.path.join(self.download_path,filename)
                                os.remove(pdf_path)
                        else:
                            pass 
                else:
                    pass
            def extract_zip_file(self):
                filelist=os.listdir(self.download_path)
                for file in filelist:
                    if file.endswith('.zip'):
                        zip_file_path=os.path.join(self.download_path,file)
                        with zipfile.ZipFile(zip_file_path,'r') as zip:
                            zip.extractall(self.download_path)
            def change_browser_flag(self,option):
                dic={'不打开浏览器':1,'打开浏览器':0}
                self.browser_flag=dic.get(option)
                return self.browser_flag
            def change_radio_flag(self,option):
                dic={'语音播报':1,'不语音播报':0}
                self.radio_flag=dic.get(option)
                return self.radio_flag 
            def download(self):
                open_Listening_mode()
                button1.config(state='disabled')
                prefs = {
                'download.default_directory': self.download_path,  # 设置默认下载路径
                "profile.default_content_setting_values.automatic_downloads": 1  # 允许多文件下载
                } 
                self.Options=Options()
                self.Options.add_argument('--disable-blink-features=AutomationControlled')#隐藏自动化控制
                self.Options.add_argument('--ignore-ssl-errosr')#忽略ssl错误
                self.Options.add_argument('--ignore-certificate-errors')
                self.Options.add_experimental_option("prefs", prefs)
                self.Options.add_experimental_option('excludeSwitches', ['enable-logging'])
                self.Options.add_experimental_option('excludeSwitches',['enable-automation'])#隐藏自动化控制
                if self.browser_flag==1:
                    self.Options.add_argument('headless')
                    self.Options.add_argument('--disable-gpu')
                    self.Options.add_argument('--window-position=-2400,-2400')
                    
                else:
                    pass
        
                self.browser=webdriver.ChromiumEdge(self.Options)
                self.browser.maximize_window()
                self.browser.get('https://www.sciengine.com/SSI/home')
                self.browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", { #执行一段js代码，隐藏自动化控制
                "source": """
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined
                })
                """})
                current_issue=self.browser.find_element(By.XPATH,'//*[@id="currentIssueTab"]')
                win=self.browser.window_handles
                self.browser.switch_to.window(win[-1])
      
                self.browser.execute_script('arguments[0].click()',current_issue)
                self.browser.implicitly_wait(10)
                more=self.browser.find_element(By.XPATH,'//*[@id="journal-content"]/div[1]/div[1]/div[3]/div[5]/div')
                self.browser.execute_script('arguments[0].click()',more)
                self.browser.implicitly_wait(10)
                win=self.browser.window_handles
                self.browser.switch_to.window(win[-1])
                select_all=self.browser.find_element(By.XPATH,'//*[@id="journal-list"]/div[1]/div[1]/div[3]/div[1]')
                self.browser.execute_script('arguments[0].click()',select_all)
                total_number=self.browser.find_element(By.XPATH,'//*[@id="selectedArticleNum"]/strong')
                total_number=int(total_number.text)
                for i in range(1,total_number+1):
                    essay_title=self.browser.find_element(By.XPATH,f'//*[@id="journal-list"]/div[1]/div[1]/div[4]/div/div[{i}]/div/div[2]/div[2]/a/span')
                    self.essay_titles.append(essay_title.text)
                full_path=os.path.join(self.download_path,'下载论文列表')
                with open(full_path,'w',encoding='utf-8') as file:
                    for essay_title in self.essay_titles: 
                        file.write(essay_title+'\n')
                self.delete_file()
                download_pdf=self.browser.find_element(By.XPATH,'//*[@id="journal-list"]/div[1]/div[1]/div[3]/div[4]')
                self.browser.execute_script('arguments[0].click()',download_pdf)
                sleep(30)
                self.browser.quit()
                cleaned_lines=[essay.replace(" ","") for essay in self.essay_titles]
                cleaned_lines=[essay.replace("\n","") for essay in cleaned_lines]
                for essay_title in cleaned_lines:
                        tree.insert("","end",values=(essay_title))
                for item in tree.get_children():      
                                tree.set(item,"第二列",'已下载')
                self.essay_titles.clear()
                self.extract_zip_file()
                close_Listening_mode()
                if self.radio_flag==1:
                    set_Volume_to_100()
                    Speak(f'中国科学:信息科学当月最新论文已为您下载到{self.download_path}请查看')
                messagebox.showinfo('提示',f'中国科学:信息科学当月最新论文当月最新论文已为您下载到{self.download_path}请查看')
        def file():
            messagebox.showinfo("提示","请选择存放所有下载pdf的文件夹")
            folder_path=filedialog.askdirectory(title='选择存放所有下载pdf的文件夹')
            if not os.path.exists(folder_path):
                            messagebox.showinfo("警告","请输入有效的地址！")
            folder_path=folder_path.replace('/','\\')
            science=中国科学(download_path=folder_path)
            global button1
            combo1=ttk.Combobox(zroot,validate='none')
            label1=tk.Label(zroot,text='选择浏览器模式',font=('宋体',10),fg='red',bg='#d7e8f0')
            label1.place(x=850,y=20)
            combo1['values']=('不打开浏览器','打开浏览器')
            combo1.current(1)
            combo1.place(x=850,y=40)
            sure_button1=tk.Button(zroot,text="确认",command=lambda:science.change_browser_flag(combo1.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
            sure_button1.place(x=1120-scaling_factor+65 if scaling_factor!=150 else 1120,y=40)
            combo2=ttk.Combobox(zroot,validate='none')
            label2=tk.Label(zroot,text='选择是否语音播报',font=('宋体',10),fg='red',bg='#d7e8f0')
            label2.place(x=850,y=120)
            combo2['values']=('语音播报','不语音播报')
            combo2.place(x=850,y=140)
            sure_button2=tk.Button(zroot,text="确认",command=lambda:science.change_radio_flag(combo2.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
            sure_button2.place(x=1120-scaling_factor+65 if scaling_factor!=150 else 1120,y=140)
            combo2.current(0)
            button1=tk.Button(zroot,text="开始下载",command=science.download,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button1.place(x=900,y=250)
            button2=tk.Button(zroot,text="清空文件夹内文件",command=science.empty_folder,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button2.place(x=900,y=400)
          
        def helper():
                messagebox.showinfo("说明","\n点击运行脚本后请勿操纵鼠标，请耐心等待自动化操作完成。\n                                        感谢您的使用！")
        def abouts():
                messagebox.showinfo("版本信息","开发者：Mr Crab\n版本号：8.0")
        menubar=tk.Menu(zroot,tearoff=True)
        zroot.config(menu=menubar)
        menubar.add_command(label="选择存放pdf的文件夹\n及本地下载的位置",font=("宋体",15),command=file)
        menubar.add_separator()
        menubar.add_command(label="说明",font=("宋体",15),command=helper)
        menubar.add_separator()
        menubar.add_command(label="关于",font=("宋体",15),command=abouts)
def 光学学报_interface():
        groot=tk.Toplevel()
        groot.geometry("1200x700")
        groot.resizable(False,False)
        groot.title("光学学报")
   
        groot.iconphoto(True,Base64_to_Image(LightScience_icon.img,100,100))
        frame=tk.Frame(groot,bg='#d7e8f0')
        frame.pack(fill="both",expand=True)
        tree=ttk.Treeview(frame,columns=("第一列","第二列"),height=30,show="headings")
        tree.heading("第一列",text="论文名称")
        tree.heading("第二列",text="下载状态")
        tree.column("第一列",width=400)
        tree.column("第二列",width=400)
        tree.grid(row=3,column=3,sticky="nsew")
        scrollbar=ttk.Scrollbar(frame,orient="vertical",command=tree.yview)
        scrollbar.grid(row=3,column=4,sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)
        class 光学学报():
            def __init__(self,download_path):
                self.download_path=download_path
                self.download_urls=[]
                self.essay_titles=[]
                self.browser_flag=0
                self.radio_flag=1
            def empty_folder(self):
                if os.listdir(self.download_path):
                    for filename in os.listdir(self.download_path):
                        file_path=os.path.join(self.download_path,filename)
                        os.remove(file_path)
                    messagebox.showinfo('提示',f'已为您清空{self.download_path}内所有文件')
                else:
                    messagebox.showinfo('提示','该文件夹内没有任何文件，无需清空！')
            def check(self,filename):
                filename=filename.replace(' ','+')
                filename=filename.replace('【增强内容出版】','')+'_光学学报.pdf'
                if '封面文章特邀综述' in filename:
                    filename=filename.replace('封面文章特邀综述','')
                if '/' in filename:
                    filename=filename.replace('/','_')
                if '亮点文章特邀综述' in filename:
                    filename=filename.replace('亮点文章特邀综述','')
                if not os.listdir(self.download_path):
                    return True
                elif filename not in os.listdir(self.download_path):
                    return True
                else:
                    return False
            def change_browser_flag(self,option):
                dic={'不打开浏览器':1,'打开浏览器':0}
                self.browser_flag=dic.get(option)
                return self.browser_flag
            def change_radio_flag(self,option):
                dic={'语音播报':1,'不语音播报':0}
                self.radio_flag=dic.get(option)
                return self.radio_flag 
            def downloads(self):
                open_Listening_mode()
                prefs = {
                'download.default_directory': self.download_path,  # 设置默认下载路径
                "profile.default_content_setting_values.automatic_downloads": 1  # 允许多文件下载
                } 
                self.Options=Options()
                self.Options.add_argument('--disable-blink-features=AutomationControlled')#隐藏自动化控制
                self.Options.add_argument('--ignore-ssl-errosr')#忽略ssl错误
                self.Options.add_argument('--ignore-certificate-errors')
                self.Options.add_experimental_option("prefs", prefs)
                self.Options.add_experimental_option('excludeSwitches',['enable-automation'])#隐藏自动化控制
                if self.browser_flag==1:
                    self.Options.add_argument('-headless=new')
                    self.Options.add_argument('--disable-gpu')
                    self.Options.add_argument('--window-position=-2400,-2400')
                else:
                    pass
                self.browser=webdriver.ChromiumEdge(self.Options)
                self.browser.maximize_window()
                self.browser.get('https://www.opticsjournal.net/Journals/gxxb.cshtml')
                self.browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", { #执行一段js代码，隐藏自动化控制
                "source": """
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined
                })
                """})
                self.browser.implicitly_wait(10)
                latest_journals_list=self.browser.find_element(By.CLASS_NAME,'latest-issue')
                main_window=self.browser.current_window_handle
                正在出版=latest_journals_list.find_element(By.PARTIAL_LINK_TEXT,'正在出版')
                已出版=latest_journals_list.find_elements(By.PARTIAL_LINK_TEXT,'已出版') 
                num=int(正在出版.text[-2])
                self.browser.execute_script('arguments[0].click()',正在出版)
                win=self.browser.window_handles
                self.browser.switch_to.window(win[-1])
                sleep(5)
                journals=self.browser.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div[2]/a')
                self.browser.execute_script('arguments[0].click()',journals)
                win=self.browser.window_handles
                self.browser.switch_to.window(win[-1])
                sleep(5)
                for i in range(2,(num+1)*2,2):
                    title=self.browser.find_element(By.XPATH,f'/html/body/div[1]/div/div/div[2]/div[3]/div[{i}]/div[1]/a')
                    if self.check(title.text):
                        self.essay_titles.append(title.text)
                        pdf=self.browser.find_element(By.XPATH,f'/html/body/div[1]/div/div/div[2]/div[3]/div[{i}]/div[7]/div[1]/a[1]')
                        self.browser.execute_script("arguments[0].click()",pdf)
                        sleep(10)
                    else:
                        pass
                # self.browser.switch_to.window(main_window)
                # self.browser.execute_script('arguments[0].click()',已出版[0])
                # win=self.browser.window_handles
                # self.browser.switch_to.window(win[-1])
                # sleep(5)
                # for i in range(2,(int(正在出版.text[-2])+1)*2,2):
                #     title=self.browser.find_element(By.XPATH,f'/html/body/div[1]/div/div/div[2]/div[3]/div[{i}]/div[1]/a')
                #     if self.check(title.text+'.pdf'):
                #         self.essay_titles.append(title.text)
                #         pdf=self.browser.find_element(By.XPATH,f'/html/body/div[1]/div/div/div[2]/div[3]/div[{i}]/div[1]/a[1]')
                #         self.browser.execute_script("arguments[0].click()",pdf)
                #         sleep(10)
                #     else:
                #         pass
                sleep(10)
                self.browser.quit()
                close_Listening_mode()
                if self.essay_titles:
                    full_path=os.path.join(self.download_path,'下载论文列表.txt')
                    with open(full_path,'a',encoding='utf-8') as titles:
                             for title in self.essay_titles:
                                 titles.write(title)
                    cleaned_lines=[essay.replace(" ","") for essay in self.essay_titles]
                    cleaned_lines=[essay.replace("\n","") for essay in cleaned_lines]
                    for essay_title in cleaned_lines:
                                tree.insert("","end",values=(essay_title))
                    for item in tree.get_children():      
                                        tree.set(item,"第二列",'已下载')
                    self.essay_titles.clear()
                    if self.radio_flag:
                        set_Volume_to_100()
                        Speak(f'光学学报当月最新论文已为您下载到{self.download_path}请查看')
                    messagebox.showinfo('提示',f'光学学报当月最新论文当月最新论文已为您下载到{self.download_path}请查看')   
                else:
                    if self.radio_flag:
                        set_Volume_to_100()
                        Speak(f'本月论文未更新\n本地文件夹{self.download_path}内即为光学学报当月全部论文,无需下载。')
                    messagebox.showinfo('提示',f'本月论文未更新\n本地文件夹内即为光学学报当月全部论文,无需下载。')
 
        def file():
            messagebox.showinfo("提示","请选择存放所有下载pdf的文件夹")
            folder_path=filedialog.askdirectory()
            if not os.path.exists(folder_path):
                            messagebox.showinfo("警告","请输入有效的地址！")
            folder_path=folder_path.replace('/','\\')
            global light
            light=光学学报(folder_path)
            global button1
            combo1=ttk.Combobox(groot,validate='none')
            label1=tk.Label(groot,text='选择浏览器模式',font=('宋体',10),fg='red',bg='#d7e8f0')
            label1.place(x=850,y=20)
            combo1['values']=('不打开浏览器','打开浏览器')
            combo1.current(1)
            combo1.place(x=850,y=40)
            sure_button1=tk.Button(groot,text="确认",command=lambda:light.change_browser_flag(combo1.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
            sure_button1.place(x=1120-scaling_factor+65 if scaling_factor!=150 else 1120,y=40)
            combo2=ttk.Combobox(groot,validate='none')
            label2=tk.Label(groot,text='选择是否语音播报',font=('宋体',10),fg='red',bg='#d7e8f0')
            label2.place(x=850,y=120)
            combo2['values']=('语音播报','不语音播报')
            combo2.place(x=850,y=140)
            sure_button2=tk.Button(groot,text="确认",command=lambda:light.change_radio_flag(combo2.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
            sure_button2.place(x=1120-scaling_factor+65 if scaling_factor!=150 else 1120,y=140)
            combo2.current(0)
            button1=tk.Button(groot,text="运行脚本",command=light.downloads,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button1.place(x=900,y=250)
            button2=tk.Button(groot,text="清空文件夹内文件",command=light.empty_folder,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
            button2.place(x=900,y=400)
        def helper():
                messagebox.showinfo("说明","\n点击运行脚本后请勿操纵鼠标，请耐心等待自动化操作完成。\n                                        感谢您的使用！")
        def abouts():
                messagebox.showinfo("版本信息","开发者：Mr Crab\n版本号：8.0")
        menubar=tk.Menu(groot,tearoff=True)
        groot.config(menu=menubar)
        menubar.add_command(label="选择存放pdf的文件夹\n及本地下载的位置",font=("宋体",15),command=file)
        menubar.add_separator()
        menubar.add_command(label="说明",font=("宋体",15),command=helper)
        menubar.add_separator()
        menubar.add_command(label="关于",font=("宋体",15),command=abouts)
def 电子学报_interface():
        droot=tk.Toplevel()
        droot.geometry("1200x700")
        droot.resizable(False,False)
        droot.title("电子学报")
        
        droot.iconphoto(True,Base64_to_Image(电子学报_icon.img,100,100))
        frame=tk.Frame(droot,bg='#d7e8f0')
        frame.pack(fill="both",expand=True)
        tree=ttk.Treeview(frame,columns=("第一列","第二列"),height=30,show="headings")
        tree.heading("第一列",text="论文名称")
        tree.heading("第二列",text="下载状态")
        tree.column("第一列",width=400)
        tree.column("第二列",width=400)
        tree.grid(row=3,column=3,sticky="nsew")
        scrollbar=ttk.Scrollbar(frame,orient="vertical",command=tree.yview)
        scrollbar.grid(row=3,column=4,sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)
        class 电子学报():
            def __init__(self,download_path):
                self.download_path=download_path
                self.download_urls=[]
                self.essay_titles=[]
                self.browser_flag=0
                self.radio_flag=1
            def empty_folder(self):
                if os.listdir(self.download_path):
                    for filename in os.listdir(self.download_path):
                        file_path=os.path.join(self.download_path,filename)
                        os.remove(file_path)
                    messagebox.showinfo('提示',f'已为您清空{self.download_path}内所有文件')
                else:
                    messagebox.showinfo('提示','该文件夹内没有任何文件，无需清空！')
            def scraping_download_urls(self,url):
    

                headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
            'Cache-Control': 'no-cache',
            'Connection': 'keep-alive',
            'Cookie': 'Secure; Secure; JSESSIONID=C192998007FAFC094F5F55D713364AC9; wkxt3_csrf_token=f0bcf357-a519-40a6-bf9e-488c0f03f411; Secure',
            'Pragma': 'no-cache',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'none',
            'Sec-Fetch-User': '?1',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36 Edg/127.0.0.0',
            'sec-ch-ua': '"Not)A;Brand";v="99", "Microsoft Edge";v="127", "Chromium";v="127"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
        }
                response = requests.get(url,headers=headers)
                soup=BeautifulSoup(response.text,'html.parser')
                divs = soup.find_all('div')
                for div in divs:
                    links = div.find_all('a')
                    for link in links:
                        href = link.get('href')
                        if 'https://www.ejournal.org.cn/CN' in str(href) and 'DZXB' in str(href):
                             self.download_urls.append(href)
                self.download_urls=list(set(self.download_urls))
            def check(self,filename):
                filename=filename.replace(':','_')
                filename=filename.replace('/','_')
                if not os.listdir(self.download_path):
                    return True
                elif filename not in os.listdir(self.download_path):
                    return True
                else:
                    return False 
            def change_browser_flag(self,option):
                dic={'不打开浏览器':1,'打开浏览器':0}
                self.browser_flag=dic.get(option)
                return self.browser_flag
            def change_radio_flag(self,option):
                dic={'语音播报':1,'不语音播报':0}
                self.radio_flag=dic.get(option)
                return self.radio_flag 
            def rectify_download_file(self):
                files=os.listdir(self.download_path)
                filename=list(filter(lambda x:'(' in x,files))
                if filename:
                    for file in filename:
                        filepath=os.path.join(self.download_path,file)
                        os.remove(filepath)
            def get_download_urls(self):
                prefs = {
                'download.default_directory': self.download_path,  # 设置默认下载路径
                "profile.default_content_setting_values.automatic_downloads": 1  # 允许多文件下载
                } 
                self.Options=Options()
                self.Options.add_argument('--disable-blink-features=AutomationControlled')#隐藏自动化控制
                self.Options.add_argument('--ignore-ssl-errosr')#忽略ssl错误
                self.Options.add_argument('--ignore-certificate-errors')
                self.Options.add_experimental_option("prefs", prefs)
                self.Options.add_experimental_option('excludeSwitches',['enable-automation'])#隐藏自动化控制
                if self.browser_flag:
                    self.Options.add_argument('-headless=new')
                    self.Options.add_argument('-disable-gpu')
                    self.Options.add_argument('--window-position=-2400,-2400')
                else:
                    pass
                self.browser=webdriver.ChromiumEdge(self.Options)
                self.browser.maximize_window()
                self.browser.get('https://www.ejournal.org.cn/CN/home')
                self.browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", { #执行一段js代码，隐藏自动化控制
                "source": """
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined
                })
                """})
                self.browser.implicitly_wait(10)
                current_issue=self.browser.find_element(By.XPATH,'//*[@id="myTab"]/li[2]/a')
                self.browser.execute_script('arguments[0].click()',current_issue)
                self.scraping_download_urls(self.browser.current_url)
                self.browser.implicitly_wait(10)
                win=self.browser.window_handles
                self.browser.switch_to.window(win[-1])
                sleep(5)  
                self.browser.quit()
            def download(self):
                open_Listening_mode()
                prefs = {
                'download.default_directory': self.download_path,  # 设置默认下载路径
                "profile.default_content_setting_values.automatic_downloads": 1  # 允许多文件下载
                } 
                self.Options=Options()
                self.Options.add_argument('--disable-blink-features=AutomationControlled')#隐藏自动化控制
                self.Options.add_argument('--ignore-ssl-errosr')#忽略ssl错误
                self.Options.add_argument('--ignore-certificate-errors')
                self.Options.add_experimental_option("prefs", prefs)
                self.Options.add_experimental_option('excludeSwitches',['enable-automation'])#隐藏自动化控制
                if self.browser_flag:
                    self.Options.add_argument('-headless=new')
                    self.Options.add_argument('-disable-gpu')
                    self.Options.add_argument('--window-position=-2400,-2400')
                else:
                    pass
                self.browser=webdriver.ChromiumEdge(self.Options)
                self.browser.maximize_window()
                self.browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", { #执行一段js代码，隐藏自动化控制
                "source": """
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined
                })
                """})
          
                for url in self.download_urls:
                    
                    self.browser.get(url)
                    self.browser.implicitly_wait(10)
                    sleep(5)

                    title=self.browser.find_element(By.XPATH,'//*[@id="metaVue"]/div[11]/div[1]/p')
                    if self.check(title.text+'.pdf'):
                        self.essay_titles.append(title.text)
                    
                        pdf=self.browser.find_element(By.XPATH,'//*[@id="metaVue"]/div[11]/div[2]/div[2]/div[4]/i')
                        self.browser.execute_script('arguments[0].click()',pdf)
                        sleep(10)
                    else:
                        pass 
                sleep(10)
                self.browser.get('https://www.ejournal.org.cn/CN/home')
                sleep(5)
                online_first=self.browser.find_element(By.XPATH,'//*[@id="mytab2"]/div/ul')
                titles=online_first.find_elements(By.CLASS_NAME,'j-title-1')
                download_pdf_buttons=online_first.find_elements(By.CLASS_NAME,'j-pdf')
                titles=[title.text for title in titles]
                self.essay_titles.extend(titles)
                downloads=dict(zip(download_pdf_buttons,titles))
                for button in download_pdf_buttons:
                    if self.check(downloads.get(button)+'.pdf'):
                        self.browser.execute_script('arguments[0].click()',button)
                        sleep(10)
                    else:
                        pass
                sleep(60)
                full_path=os.path.join(self.download_path,'下载论文列表.txt')
                if self.essay_titles:
                    with open(full_path,'a',encoding='utf-8') as file:
                        for essay in self.essay_titles:
                            file.write(essay+'\n')
                self.browser.quit()
                close_Listening_mode()
            def search(self):
                    button1.config(state='disabled')
                    thread1=threading.Thread(target=self.get_download_urls)
                    thread2=threading.Thread(target=self.download)
                    thread1.start()
                    thread1.join()
                    thread2.start()
                    thread2.join()
                    cleaned_lines=[essay.replace(" ","") for essay in self.essay_titles]
                    cleaned_lines=[essay.replace("\n","") for essay in cleaned_lines]
                    self.rectify_download_file()
                    if cleaned_lines:
                        for essay_title in cleaned_lines:
                                tree.insert("","end",values=(essay_title))
                        for item in tree.get_children():      
                                        tree.set(item,"第二列",'已下载')
                        self.essay_titles.clear()
                        self.download_urls.clear()
                        if self.radio_flag:
                            set_Volume_to_100()
                            Speak(f'电子学报当月最新论文已为您下载到{self.download_path}请查看')
                        messagebox.showinfo('提示',f'电子学报当月最新论文当月最新论文已为您下载到{self.download_path}请查看')
                    else:
                        if self.radio_flag:
                            set_Volume_to_100()
                            Speak(f'本月论文未更新\n本地文件夹{self.download_path}即为电子学报当月全部论文,无需下载。')
                        messagebox.showinfo('提示',f'本月论文未更新\n本地文件夹内即为电子学报当月全部论文,无需下载。')
        def file():
                messagebox.showinfo("提示","请选择存放所有下载pdf的文件夹")
                folder_path=filedialog.askdirectory()
                if not os.path.exists(folder_path):
                            messagebox.showinfo("警告","请输入有效的地址！")
                folder_path=folder_path.replace('/','\\')
                electric=电子学报(folder_path)
                global button1
                combo1=ttk.Combobox(droot,validate='none')
                label1=tk.Label(droot,text='选择浏览器模式',font=('宋体',10),fg='red',bg='#d7e8f0')
                label1.place(x=850,y=20)
                combo1['values']=('不打开浏览器','打开浏览器')
                combo1.current(1)
                combo1.place(x=850,y=40)
                sure_button1=tk.Button(droot,text="确认",command=lambda:electric.change_browser_flag(combo1.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
                sure_button1.place(x=1120-scaling_factor+65 if scaling_factor!=150 else 1120,y=40)
                combo2=ttk.Combobox(droot,validate='none')
                label2=tk.Label(droot,text='选择是否语音播报',font=('宋体',10),fg='red',bg='#d7e8f0')
                label2.place(x=850,y=120)
                combo2['values']=('语音播报','不语音播报')
                combo2.place(x=850,y=140)
                sure_button2=tk.Button(droot,text="确认",command=lambda:electric.change_radio_flag(combo2.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
                sure_button2.place(x=1120-scaling_factor+65 if scaling_factor!=150 else 1120,y=140)
                combo2.current(0)
                button1=tk.Button(droot,text="运行脚本",command=electric.search,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
                button1.place(x=900,y=250)
                button2=tk.Button(droot,text="清空文件夹内文件",command=electric.empty_folder,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
                button2.place(x=900,y=400)
        def helper():
                        messagebox.showinfo("说明","\n点击运行脚本后请勿操纵鼠标，请耐心等待自动化操作完成。\n                                        感谢您的使用！")
        def abouts():
                        messagebox.showinfo("版本信息","开发者：Mr Crab\n版本号：8.0")
       
        menubar=tk.Menu(droot,tearoff=True)
        droot.config(menu=menubar)
        menubar.add_command(label="选择存放pdf的文件夹\n及本地下载的位置",font=("宋体",15),command=file)
        menubar.add_separator()
        menubar.add_command(label="说明",font=("宋体",15),command=helper)
        menubar.add_separator()
        menubar.add_command(label="关于",font=("宋体",15),command=abouts)
def 电子测量与仪器_interface():   
        droot=tk.Toplevel()
        droot.geometry("1200x700")
        droot.resizable(False,False)
        droot.title("电子测量与仪器学报")
      
        droot.iconphoto(True,Base64_to_Image(电子测量与仪器学报_icon.img,100,100))
        frame=tk.Frame(droot,bg='#d7e8f0')
        frame.pack(fill="both",expand=True)
        tree=ttk.Treeview(frame,columns=("第一列","第二列"),height=30,show="headings")
        tree.heading("第一列",text="论文名称")
        tree.heading("第二列",text="下载状态")
        tree.column("第一列",width=400)
        tree.column("第二列",width=400)
        tree.grid(row=3,column=3,sticky="nsew")
        scrollbar=ttk.Scrollbar(frame,orient="vertical",command=tree.yview)
        scrollbar.grid(row=3,column=4,sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)
        class 电子测量与仪器():
            def __init__(self,download_path):
          
                self.download_path=download_path
                self.essay_titles=[]
                self.browser_flag=0
                self.radio=1
            def empty_folder(self):
                if os.listdir(self.download_path):
                    for filename in os.listdir(self.download_path):
                        file_path=os.path.join(self.download_path,filename)
                        os.remove(file_path)
                    messagebox.showinfo('提示',f'已为您清空{self.download_path}内所有文件')
                else:
                    messagebox.showinfo('提示','该文件夹内没有任何文件，无需清空！')
            def check(self,filename):
                if not os.listdir(self.download_path):
                    return True
                elif filename not in os.listdir(self.download_path):
                    return True
                else:
                    return False
            def change_browser_flag(self,option):
                dic={'不打开浏览器':1,'打开浏览器':0}
                self.browser_flag=dic.get(option)
                return self.browser_flag
            def change_radio_flag(self,option):
                dic={'语音播报':1,'不语音播报':0}
                self.radio_flag=dic.get(option)
                return self.radio_flag 
            def move(self):
                  for root,dirs,files in os.walk(self.downloads_path):
                      for file in files:  
                        if file.endswith('pdf'):
                             source_file_path=os.path.join(root,file)
                             target_file_path=os.path.join(self.download_path,file)
                             shutil.move(source_file_path,target_file_path)
            def download(self):
                open_Listening_mode()
                button1.config(state='disabled')
                prefs = {
                'download.default_directory': self.download_path,  # 设置默认下载路径
                "profile.default_content_setting_values.automatic_downloads": 1  # 允许多文件下载
                } 
                self.Options=Options()
                self.Options.add_argument('--disable-blink-features=AutomationControlled')#隐藏自动化控制
                self.Options.add_argument('--ignore-ssl-errosr')#忽略ssl错误
                self.Options.add_argument('--ignore-certificate-errors')
                self.Options.add_experimental_option("prefs", prefs)
                self.Options.add_experimental_option('excludeSwitches',['enable-automation'])#隐藏自动化控制
                if self.browser_flag:
                    self.Options.add_argument('-headless=new')
                    self.Options.add_argument('-disable-gpu')
                    self.Options.add_argument('--window-position=-2400,-2400')
                else:
                    pass
                self.browser=webdriver.ChromiumEdge(self.Options)
                self.browser.get('http://jemi.etmchina.com/jemi/home')
                self.browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", { #执行一段js代码，隐藏自动化控制
                "source": """
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined
                })
                """})
                self.browser.maximize_window()
                article_list=self.browser.find_element(By.XPATH,'/html/body/div[2]/div/div[1]/div/div[2]/div/div[2]/ul[1]/div[2]/ul')
                articles=article_list.find_elements(By.TAG_NAME,'li')
                sleep(5)
                for i in range(1,len(articles)+1):
                    title=self.browser.find_element(By.XPATH,f'/html/body/div[2]/div/div[1]/div/div[2]/div/div[2]/ul[1]/div[2]/ul/li[{i}]/div/div[1]/a')
                    pdf=self.browser.find_element(By.XPATH,f'/html/body/div[2]/div/div[1]/div/div[2]/div/div[2]/ul[1]/div[2]/ul/li[{i}]/div/div[2]/a[3]')
                    sleep(3)
                    if self.check(title.text+'.pdf')==True:
                        self.essay_titles.append(title.text)
                        self.browser.execute_script('arguments[0].click()',pdf)
                        sleep(10)
                    else:
                         pass
                sleep(90)
                self.browser.quit()
                close_Listening_mode()
                full_path=os.path.join(self.download_path,'下载论文列表.txt')
                if self.essay_titles:
                    with open(full_path,'a',encoding='utf-8') as file:
                        file.write('电子测量与仪器当月最新论文'+'\n')
                        for essay in self.essay_titles:
                            file.write(essay+'\n')
                    cleaned_lines=[essay.replace(" ","") for essay in self.essay_titles]
                    cleaned_lines=[essay.replace("\n","") for essay in cleaned_lines]
                    for essay_title in cleaned_lines:
                            tree.insert("","end",values=(essay_title))
                    for item in tree.get_children():      
                                  tree.set(item,"第二列",'已下载')
                    self.essay_titles.clear()
                    if self.radio_flag:
                        set_Volume_to_100()
                        Speak(f'电子测量与仪器当月最新论文已为您下载到{self.download_path}请查看')
                    messagebox.showinfo('提示',f'电子测量与仪器当月最新论文当月最新论文已为您下载到{self.download_path}请查看')
             
                else:
                    self.essay_titles.clear()
                    if self.radio_flag:
                        set_Volume_to_100()
                        Speak(f'本月论文未更新\n本地文件夹{self.download_path}即为当月全部论文,无需下载。')
                    messagebox.showinfo('提示',f'本月论文未更新\n本地文件夹内即为电子测量与仪器当月全部论文,无需下载。')
                    
        def file():
                messagebox.showinfo("提示","请选择存放存放下载论文pdf的文件夹")
                folder_path=filedialog.askdirectory()
                if not os.path.exists(folder_path):
                            messagebox.showinfo("警告","请输入有效的地址！")
                folder_path=folder_path.replace('/','\\')
                electric=电子测量与仪器(folder_path)
                global button1
                combo1=ttk.Combobox(droot,validate='none')
                label1=tk.Label(droot,text='选择浏览器模式',font=('宋体',10),fg='red',bg='#d7e8f0')
                label1.place(x=850,y=20)
                combo1['values']=('不打开浏览器','打开浏览器')
                combo1.current(1)
                combo1.place(x=850,y=40)
                sure_button1=tk.Button(droot,text="确认",command=lambda:electric.change_browser_flag(combo1.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
                sure_button1.place(x=1120-scaling_factor+65 if scaling_factor!=150 else 1120,y=40)
                combo2=ttk.Combobox(droot,validate='none')
                label2=tk.Label(droot,text='选择是否语音播报',font=('宋体',10),fg='red',bg='#d7e8f0')
                label2.place(x=850,y=120)
                combo2['values']=('语音播报','不语音播报')
                combo2.place(x=850,y=140)
                sure_button2=tk.Button(droot,text="确认",command=lambda:electric.change_radio_flag(combo2.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
                sure_button2.place(x=1120-scaling_factor+65 if scaling_factor!=150 else 1120,y=140)
                combo2.current(0)
                button1=tk.Button(droot,text="运行脚本",command=electric.download,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
                button1.place(x=900,y=250)
                button2=tk.Button(droot,text="清空文件夹内文件",command=electric.empty_folder,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
                button2.place(x=900,y=400)
        def helper():
                        messagebox.showinfo("说明","\n点击运行脚本后请勿操纵鼠标，请耐心等待自动化操作完成。\n                                        感谢您的使用！")
        def abouts():
                        messagebox.showinfo("版本信息","开发者：Mr Crab\n版本号：8.0")
       
        menubar=tk.Menu(droot,tearoff=True)
        droot.config(menu=menubar)
        menubar.add_command(label="选择存放pdf的文件夹\n及本地下载的位置",font=("宋体",15),command=file)
        menubar.add_separator()
        menubar.add_command(label="说明",font=("宋体",15),command=helper)
        menubar.add_separator()
        menubar.add_command(label="关于",font=("宋体",15),command=abouts)
def 通信学报_interface():   
        troot=tk.Toplevel()
        troot.geometry("1200x700")
        troot.resizable(False,False)
        troot.title("通信学报")
        
        troot.iconphoto(True,Base64_to_Image(通信学报_icon.img,100,100))
        frame=tk.Frame(troot,bg='#d7e8f0')
        frame.pack(fill="both",expand=True)
        tree=ttk.Treeview(frame,columns=("第一列","第二列"),height=30,show="headings")
        tree.heading("第一列",text="论文名称")
        tree.heading("第二列",text="下载状态")
        tree.column("第一列",width=400)
        tree.column("第二列",width=400)
        tree.grid(row=3,column=3,sticky="nsew")
        scrollbar=ttk.Scrollbar(frame,orient="vertical",command=tree.yview)
        scrollbar.grid(row=3,column=4,sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)
        class 通信学报():
            def __init__(self,download_path):
               self.download_path=download_path
               self.essay_titles=[] 
               self.browser_flag=0
               self.radio_flag=1
            def empty_folder(self):
                if os.listdir(self.download_path):
                    for filename in os.listdir(self.download_path):
                        file_path=os.path.join(self.download_path,filename)
                        os.remove(file_path)
                    messagebox.showinfo('提示',f'已为您清空{self.download_path}内所有文件')
                else:
                    messagebox.showinfo('提示','该文件夹内没有任何文件，无需清空！')
            def check(self,filename):
                if not os.listdir(self.download_path):
                    return True
                elif filename not in os.listdir(self.download_path):
                    return True
                else:
                    return False
            def change_browser_flag(self,option):
                dic={'不打开浏览器':1,'打开浏览器':0}
                self.browser_flag=dic.get(option)
                return self.browser_flag
            def change_radio_flag(self,option):
                dic={'语音播报':1,'不语音播报':0}
                self.radio_flag=dic.get(option)
                return self.radio_flag 
            def download(self):
                open_Listening_mode()
                button1.config(state='disabled')
                prefs = {
                'download.default_directory': self.download_path,  # 设置默认下载路径
                "profile.default_content_setting_values.automatic_downloads": 1  # 允许多文件下载
                } 
                self.Options=Options()
                self.Options.add_argument('--disable-blink-features=AutomationControlled')#隐藏自动化控制
                self.Options.add_argument('--ignore-ssl-errosr')#忽略ssl错误
                self.Options.add_argument('--ignore-certificate-errors')
                self.Options.add_experimental_option("prefs", prefs)
                self.Options.add_experimental_option('excludeSwitches',['enable-automation'])#隐藏自动化控制
                if self.browser_flag:
                    self.Options.add_argument('-headless=new')
                    self.Options.add_argument('-disable-gpu')
                    self.Options.add_argument('--window-position=-2400,-2400')
                self.browser=webdriver.ChromiumEdge(self.Options)
                self.browser.get('http://www.joconline.com.cn/')
                self.browser.maximize_window() 
                sleep(5)
                更多=self.browser.find_element(By.XPATH,'//*[@id="container-32623"]/div/div[1]/span')
                self.browser.execute_script('arguments[0].click()',更多)
                win=self.browser.window_handles
                self.browser.switch_to.window(win[-1])
                sleep(5)
                labels=self.browser.find_elements(By.CLASS_NAME,'label-content')
                downloads=[label.find_element(By.CLASS_NAME,'download-span') for label in labels]
                downloadpdf_button=[download.find_elements(By.TAG_NAME,'em')[0] for download in downloads]
                self.essay_titles=self.browser.find_elements(By.CLASS_NAME,'resName')
                self.essay_titles=[title.text for title in self.essay_titles]
                dic=dict(zip(downloadpdf_button,self.essay_titles))
                for pdf_button in dic:
                    if self.check(dic.get(pdf_button)+'_NormalPdf.pdf'):
                        self.browser.execute_script('arguments[0].click()',pdf_button)
                        sleep(10)
                    else:
                        self.essay_titles.pop()
                        pass
                sleep(30)
                self.browser.quit()
                close_Listening_mode()
                full_path=os.path.join(self.download_path,'下载列表.txt')
                if self.essay_titles:
                    with open(full_path,'a',encoding='utf-8') as pdf:
                        pdf.write('\n')
                        for title in self.essay_titles:
                            pdf.write(title+'\n')
                    for essay_title in self.essay_titles:
                                tree.insert("","end",values=(essay_title))
                    for item in tree.get_children():      
                                      tree.set(item,"第二列",'已下载')
                    self.essay_titles.clear()
                    if self.radio_flag: 
                        set_Volume_to_100()
                        Speak(f'通信学报当月最新论文已为您下载到{self.download_path}请查看')
                    messagebox.showinfo('提示',f'通信学报当月最新论文当月最新论文已为您下载到{self.download_path}请查看')
                else:
                    if self.radio_flag:
                        set_Volume_to_100()
                        Speak(f'本月论文未更新\n本地文件夹{self.download_path}即为当月全部论文,无需下载。')
                    messagebox.showinfo('提示',f'本月论文未更新\n本地文件夹内即为通信学报当月全部论文,无需下载。')
        def file():
                messagebox.showinfo("提示","请选择存放所有下载pdf的文件夹")
                folder_path=filedialog.askdirectory()
                if not os.path.exists(folder_path):
                            messagebox.showinfo("警告","请输入有效的地址！")
                folder_path=folder_path.replace('/','\\')
                electric=通信学报(folder_path)
                global button1
                combo1=ttk.Combobox(troot,validate='none')
                label1=tk.Label(troot,text='选择浏览器模式',font=('宋体',10),fg='red',bg='#d7e8f0')
                label1.place(x=850,y=20)
                combo1['values']=('不打开浏览器','打开浏览器')
                combo1.current(1)
                combo1.place(x=850,y=40)
                sure_button1=tk.Button(troot,text="确认",command=lambda:electric.change_browser_flag(combo1.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
                sure_button1.place(x=1120-scaling_factor+65 if scaling_factor!=150 else 1120,y=40)
                combo2=ttk.Combobox(troot,validate='none')
                label2=tk.Label(troot,text='选择是否语音播报',font=('宋体',10),fg='red',bg='#d7e8f0')
                label2.place(x=850,y=120)
                combo2['values']=('语音播报','不语音播报')
                combo2.place(x=850,y=140)
                sure_button2=tk.Button(troot,text="确认",command=lambda:electric.change_radio_flag(combo2.get()),width=6,height=1,relief="raised",fg="#000000",bg="#0087FF")
                sure_button2.place(x=1120-scaling_factor+65 if scaling_factor!=150 else 1120,y=140)
                combo2.current(0)
                button1=tk.Button(troot,text="运行脚本",command=electric.download,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
                button1.place(x=900,y=250)
                button2=tk.Button(troot,text="清空文件夹内文件",command=electric.empty_folder,width=15,height=3,relief="raised",fg="#000000",bg="#0087FF")
                button2.place(x=900,y=400)
        def helper():
                        messagebox.showinfo("说明","\n点击运行脚本后请勿操纵鼠标，请耐心等待自动化操作完成。\n                                        感谢您的使用！")
        def abouts():
                        messagebox.showinfo("版本信息","开发者：Mr Crab\n版本号：8.0")
       
        menubar=tk.Menu(troot,tearoff=True)
        troot.config(menu=menubar)
        menubar.add_command(label="选择存放pdf的文件夹\n及本地下载的位置",font=("宋体",15),command=file)
        menubar.add_separator()
        menubar.add_command(label="说明",font=("宋体",15),command=helper)
        menubar.add_separator()
        menubar.add_command(label="关于",font=("宋体",15),command=abouts)
def 项目原码_interface():
        root=tk.Toplevel()
        root.geometry("700x900")
        root.resizable(False,False)
        root.title("项目原码")
        root.iconphoto(True,Base64_to_Image(项目原码_icon.img,100,100))
        main_canvas=tk.Canvas(root,width=700,height=800)
        im_root=Base64_to_Image(base64_string=微信名片.img,width=700,height=800)
        main_canvas.create_image(350,400,anchor="center",image=im_root)
        main_canvas.pack()
        label=tk.Label(root,text=f'扫描上方二维码添加Mr.Crab微信\n获取项目源码',font=("宋体",15),fg='#ff6f91')
        label.pack()
        root.mainloop()
      
###########################################################################
'''主界面'''
def main_interface():
    global root
    root=tk.Tk()
    root.title("论文自动化助手8.0")
    root.geometry(f"{sizes['main_size'][0]}x{sizes['main_size'][1]}")
    root.iconphoto(False,Base64_to_Image(主界面_icon.img,100,100))
    root.resizable(False,False)
    global main_canvas
    main_canvas=tk.Canvas(root,width=sizes['main_size'][0],height=sizes["main_size"][1])
    im_root=Base64_to_Image(主界面背景.img,sizes['main_size'][0],sizes["main_size"][1])
    main_canvas.create_image(int(sizes['main_size'][0]/2),int(sizes["main_size"][1])/2,anchor='center',image=im_root)
    main_canvas.pack(fill='both',expand=True)
    Welcomelabel=tk.Label(main_canvas,text="欢迎使用",font=("宋体",20),fg="#845ec2",bg='#00c9a7')
    Code_src=tk.Button(main_canvas,text="项目原码",command=项目原码_interface,width=18,height=3,relief="raised",fg="#FFFFFF",bg="#FF8066",activebackground="#ff8066")
    Pdf_statistic=tk.Button(main_canvas,text="pdf关键字统计",command=statistic,width=18,height=3,relief="raised",fg="#FFFFFF",bg="#FF8066",activebackground="#ff8066")
    menubutton1=tk.Menubutton(main_canvas,text="当月论文自动化下载",font=("宋体",13),width=18,height=3,anchor="center",relief="raised",bg="#FF8066",fg="#FFFFFF",activebackground="#ff8066")
    menu1=tk.Menu(menubutton1,tearoff=True)
    
    menu1.add_command(label="中国科学: 信息科学论文下载",activebackground="green",command=中国科学_interface)
    menu1.add_command(label="光学学报论文下载",activebackground="green",command=光学学报_interface)
    menu1.add_command(label="电子学报论文下载",activebackground="green",command=电子学报_interface)
    menu1.add_command(label="电子测量与仪器学报论文下载",activebackground="green",command=电子测量与仪器_interface)
    menu1.add_command(label="通信学报论文下载",activebackground="green",command=通信学报_interface)
    menu1.add_command(label="Light: Science & Application论文下载",activebackground="green",command=Light_Science_Application_interface)
    menu1.add_command(label="SCIENCE CHINA Information Sciences论文下载 ",activebackground="green",command=SCIENCE_CHINA_Information_Sciences_interface)
    menu1.add_command(label="Journal of Lightwave Technology论文下载",activebackground="green",command=journal_of_lightwave_technology_interface)
    menu1.add_command(label="Journal of Optical Communications and Networking论文下载",activebackground="green",command=journal_of_Optical_Communications_and_Networking_interface)
    menu1.add_command(label="IEEE Photonics Technology Letters论文下载",activebackground="green",command=IEEE_Photonics_Technology_Letters_interface)
    menu1.add_command(label="Optics Express论文下载",activebackground="green",command=optic_express_interface)
    menu1.add_command(label="多项论文下载",activebackground="green",command=under_developing)
    menubutton1.config(menu=menu1)

    menubutton2=tk.Menubutton(main_canvas,text="自动化搜索结果统计",font=("宋体",13),width=18,height=3,anchor="center",relief="raised",bg="#FF8066",fg="#FFFFFF",activebackground="#ff8066")
    menu2=tk.Menu(menubutton2,tearoff=True)
    menu2.add_command(label="IEEE自动化搜索结果统计",activebackground="green",command=IEEE)
    menu2.add_command(label="谷歌学术自动化搜索结果统计",activebackground="green",command=Google)
    menu2.add_command(label="中国知网自动化搜索结果统计",activebackground="green",command=CNKI)
    menubutton2.config(menu=menu2)
    # Welcomelabel.place(x=300,y=100)
    # menubutton1.place(x=10,y=240)
    # menubutton2.place(x=530,y=240)
    # Code_src.place(x=10,y=680)
    # Pdf_statistic.place(x=590,y=680)
    Welcomelabel.grid(row=0,column=1,sticky='n')
    menubutton1.grid(row=2,column=0,sticky='w',pady=150)
    menubutton2.grid(row=2,column=2,pady=150,sticky='e')
    Code_src.grid(row=3,column=2,sticky='se',pady=300)
    Pdf_statistic.grid(row=3,column=0,sticky='sw',pady=300)

    if main_canvas.winfo_exists():
        get_time()
    root.mainloop()
 

      
     
      

############################################################################
'''警告和免责声明界面'''           
def run():
    start.destroy()
    main_interface()
scaling_factor=get_windows_scaling_factor()
sizes={"start_size":(int(scaling_factor*3.2+520),int(scaling_factor*1.6+80)),"main_size":(int(scaling_factor*3.2+320),int(scaling_factor*1.6+660))}
start=tk.Tk()
start.title("警告和免责声明")
start.geometry(f"{sizes['start_size'][0]}x{sizes["start_size"][1]}")
start.resizable(False,False)
start.iconphoto(False,Base64_to_Image(注意_icon.img,100,100))
scrolled_text=tk.Label(start,text=f'''   警告和免责声明：

    本GUI由Mr.Crab开发,用于自动化论文下载,搜索结果统计等操作。
使用本GUI须自行承担风险,Mr.Crab不对任何因使用本GUI而产生的直接
间接损失负责。请遵守学术规范和论文网站的使用条款和法律法规,
本GUI不用于任何非法或违反道德的行为。未经Mr.Crab明确授权
禁止对本GUI进行复制、分发或商业使用。感谢您的使用!''',font=("宋体",15),fg="#c34a36")
scrolled_text.pack()
surebutton=tk.Button(start,text="确定(开始使用)",command=run,font=("宋体",12),width=20,bg="#0081cf",activebackground="red",fg="black",height=1)
    # surebutton.place(x=50,y=350)
surebutton.pack(side='left',padx=30*scaling_factor/100)
cancelbutton=tk.Button(start,text="取消(退出使用)",command=start.destroy,font=("宋体",12),bg="#00c9a7",activebackground="red",fg="black",width=20,height=1)
    # cancelbutton.place(x=550,y=350)
cancelbutton.pack(side='right',padx=30)
if start.winfo_exists():
    start.after(ms=6000,func=run)
start.mainloop()  



  