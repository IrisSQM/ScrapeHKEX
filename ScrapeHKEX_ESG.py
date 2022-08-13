#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Scrape HKEX ESG

Created on Sat Aug  6 10:37:53 2022

@author: shiqimeng
"""
from joblib import Parallel, delayed
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
import requests
import pandas as pd
import time
from random import uniform
    
# ------------ 获取所有上市公司代码（主板+GEM） -------------
def getCodes(save_path, driver_path):
    '''
    获取上市公司代码

    Parameters
    ----------
    save_path : TYPE str
        文件保存路径
    driver_path : TYPE str
        webdriver 路径

    Returns
    -------
    ls_main_cd : TYPE list of str
        主板公司代码列表
    ls_gem_cd : TYPE list of str
        中小板公司代码列表

    '''

    # Set webdriver
    browser = webdriver.Chrome(service = driver_path)
    
    # First go to listing info page
    listing_path = 'https://www1.hkexnews.hk/app/appindex.html?lang=zh'
    browser.get(listing_path)
    time.sleep(uniform(1,1.3))
    
    # click "已上市"
    browser.find_element(By.XPATH,'//*[@id="tab-panel-main-board"]/ul/li[3]/a').click()
    time.sleep(uniform(1,1.3))
    
    # 获取主板代码表格
    datalist = []
    header = [i.text for i in browser.find_elements(By.CSS_SELECTOR,'thead>tr>th') if i.text]
    rows = [i for i in browser.find_elements(By.CSS_SELECTOR,'tbody>tr') if i.text]
    
    for row in rows:
        datalist.append([i.text for i in row.find_elements(By.CSS_SELECTOR,'td') if i.text])
    
    df = pd.DataFrame(datalist, columns = header)
    ## 最后四行是注释
    
    # 保存csv
    df.to_csv(save_path  + os.sep + '结果' + os.sep + '主板上市公司及代码.csv')
    
    # 保存主板公司代码，后续爬取ESG报告用
    ls_main_cd = df.iloc[:,0][:-4].to_list()
    
    # click "GEM"-"已上市"
    browser.find_element(By.XPATH,'/html/body/printfriendly/div/main/div[1]/div[1]/ul/div[2]/li[2]/a').click()
    browser.find_element(By.XPATH,'//*[@id="tab-panel-main-board"]/ul/li[3]/a').click()
    time.sleep(uniform(3,4))
    
    # 获取GEM板代码表格
    datalist = []
    header = [i.text for i in browser.find_elements(By.CSS_SELECTOR,'thead>tr>th') if i.text]
    rows = [i for i in browser.find_elements(By.CSS_SELECTOR,'tbody>tr') if i.text]
    
    for row in rows:
        datalist.append([i.text for i in row.find_elements(By.CSS_SELECTOR,'td') if i.text])
    
    df_2 = pd.DataFrame(datalist, columns = header)
    
    # 保存csv
    df_2.to_csv(save_path  + os.sep + '结果' + os.sep + 'GEM上市公司及代码.csv')
    
    # 保存GEM公司代码，后续爬取ESG报告用
    ls_gem_cd = df_2.iloc[:,0].to_list()
    
    browser.close()
    
    return ls_main_cd, ls_gem_cd

# ------------ 下载公司所有的 ESG 报告（主板+GEM） -------------
def getComp(comp, esg_path, driver_path):
    '''

    Parameters
    ----------
    comp : TYPE str
        公司代码
    esg_path : TYPE str
        esg报告文件夹路径
    driver_path : TYPE str
        webdriver 路径
    Returns
    -------
    None.
    文件自动保存到对应文件夹

    '''
    
    # 打开查询页
    browser = webdriver.Chrome(service = driver_path)
    search_path = 'https://www1.hkexnews.hk/search/titlesearch.xhtml?lang=zh'
    browser.get(search_path)
    time.sleep(uniform(1,1.3))
    
    # 输入并选择公司
    browser.find_element(By.XPATH,'//*[@id="searchStockCode"]').send_keys(comp)
    time.sleep(uniform(2,2.3))
    
    try:
        browser.find_element(By.XPATH,
                             '//*[@id="autocomplete-list-0"]/div[1]/div[1]/table/tbody/tr[1]/td[1]').click()
    except:
        browser.find_element(By.XPATH,'//*[@id="autocomplete-list-0"]/div[1]/div[1]/table/tbody/tr[1]/td[2]').click()
    # 选择Headline筛选方式
    browser.find_element(By.XPATH,'//*[@id="tier1-select"]/div/div/a').click()
    browser.find_element(By.XPATH,'//*[@id="hkex_news_header_section"]/section/div[1]/div/div[2]/ul/li[4]/div/div[2]/div[1]/div[2]/div/div/div/div[1]/div[2]/a').click()
    browser.find_element(By.XPATH,'//*[@id="rbAfter2006"]/div[1]/div/div/a').click()
    browser.find_element(By.XPATH,'//*[@id="rbAfter2006"]/div[2]/div/div/div/ul/li[5]').click()
    time.sleep(uniform(0.8,1.3))
    
    # 移动滚动条
    actions = ActionChains(browser)
    bar = browser.find_element(By.XPATH,'//*[@id="rbAfter2006"]/div[2]/div/div/div/div[1]')
    actions.drag_and_drop_by_offset(bar,0,30)
    actions.perform()
    
    browser.find_element(By.XPATH,'//*[@id="rbAfter2006"]/div[2]/div/div/div/ul/li[5]/div/ul/li[3]/a').click()
    
    ele = browser.find_element(By.XPATH,'//*[@id="rbAfter2006"]/div[1]/div/div/a')
    actions.move_to_element(ele).perform()
    actions.move_by_offset(400,60).click().perform() # 模拟移动鼠标点击
    
    time.sleep(uniform(2,3))
    
    # 爬取结果
    # 建公司文件夹
    os.mkdir(esg_path + os.sep + comp)
    doc_path = esg_path + os.sep + comp
    
    # 总共的行数
    row_no = len(browser.find_elements(By.CLASS_NAME,'doc-link'))
        
    # 记录当前页
    main_win = browser.current_window_handle
    
    for r in range(row_no):
        
        doc = browser.find_element(By.XPATH,
                '//*[@id="titleSearchResultPanel"]/div/div[1]/div[3]/div[2]/table/tbody/tr[{}]/td[4]/div[2]/a'.format(str(r+1)))
        name = doc.text + '.pdf'
        
        doc.click()  #点击跳转页面
        
        browser.switch_to.window(browser.window_handles[1])
        
        r = requests.get(browser.current_url)
        
        # 保存文件
        with open(doc_path + os.sep + name, 'wb+') as f:
            f.write(r.content)
            
        # close pop-up window
        browser.close()
        # switch back
        browser.switch_to.window(main_win)
    
    browser.close()

# ------------ 主程序 -------------        
if __name__ == '__main__':
    
    # ！！修改路径！！
    save_path = '<你的文件路径>'
    driver_path = Service(save_path + os.sep + 'chromedriver') # chromedriver路径

    # 创建代码结果文件夹
    os.mkdir(save_path + os.sep + '结果2')
    
    # 先获取公司代码
    ls_main_cd, ls_gem_cd = getCodes(save_path, driver_path)
    
    # 建esg文件夹
    esg_path = save_path + os.sep + '结果2' + os.sep + '年报ESG报告'
    os.mkdir(esg_path)
    
    esg_path_main = esg_path + os.sep + '主板'
    esg_path_gem = esg_path + os.sep + 'GEM'
    os.mkdir(esg_path_main)
    os.mkdir(esg_path_gem)
    
    # 分别爬取主板、GEM板公司 
    Parallel(n_jobs = 4, backend='threading')(delayed(getComp)(
        comp, esg_path_main, driver_path) for comp in ls_main_cd[:4]) # 前4家主板
    
    time.sleep(uniform(2,3))
    
    Parallel(n_jobs = 4, backend='threading')(delayed(getComp)(
        comp, esg_path_gem, driver_path) for comp in ls_gem_cd[:4]) # 前4家中小板
    