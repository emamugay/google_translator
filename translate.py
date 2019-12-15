#!/usr/bin/python
# -*- coding: utf-8 -*-
import pandas as pd
import os, glob, sys, socket, urllib, math, time
import requests, json, re, subprocess
import pycurl
from bs4 import BeautifulSoup
import dateutil.parser
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, NoSuchWindowException, TimeoutException, WebDriverException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import selenium.webdriver.support.ui as ui
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.firefox.options import Options
#import xlwings as xw
import csv
path = os.path.dirname(os.path.realpath(__file__))

from openpyxl import Workbook, load_workbook


platform = sys.platform
arg = sys.argv[1:]

def printer(text):
    print(text)
    return

def run_cmd(cmd):
    p = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
    output = p.communicate()[0].split()
    for n in output:
        if 'addr:' in str(n):
            output = n.replace('addr:','').strip()
            break

    return output

def start_new_session():
    WINDOW_SIZE = "1920,1080"

    #chrome_options = webdriver.ChromeOptions()
    #chrome_options.add_argument("--ignore-ssl-errors=true")
    #chrome_options.add_argument("--ssl-protocol=TLSv1")
    #chrome_options.add_argument("--hide-scrollbars")
    #chrome_options.add_argument("--disable-gpu")
    #chrome_options.add_argument("--log-level=3")
    #chrome_options.add_argument("--window-size=%s" % WINDOW_SIZE)
    #chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.181 Safari/537.36")

    ##chrome_options.add_argument("--headless")
    #chrome_options.add_argument("--headed")
    #path = "/opt/"
    #CHROMEDRIVER_PATH = "/mnt/Devs/Python/scrapy/mmi/driver/geckodriver"
    ##CHROMEDRIVER_PATH = "/mnt/Devs/Python/scrapy/mmi/driver/chromedriver"
    ##chrome_options.binary_location = "/usr/bin/google-chrome"
    #chrome_options.binary_location = "/usr/bin/firefox"

    #chrome_options.add_experimental_option("prefs", {
    #"download.default_directory": path + "data",
    #"download.prompt_for_download": False,
    #"download.directory_upgrade": True,
    #"safebrowsing.enabled": True
    #})

    #os.environ["webdriver.chrome.driver"] = CHROMEDRIVER_PATH
    ##driver = webdriver.Chrome(executable_path = CHROMEDRIVER_PATH, options = chrome_options)
    #driver = webdriver.Firefox(executable_path = CHROMEDRIVER_PATH, options = chrome_options)
    #driver.command_executor._commands["send_command"] = ("POST", '/session/' + driver.session_id + '/chromium/send_command')

    #params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': path + "data"}}
    #command_result = driver.execute("send_command", params)

    options = Options()
    options.headless = False

    DRIVER_PATH = "/mnt/Devs/Python/scrapy/mmi/driver/geckodriver"
    options.binary_location = "/usr/bin/firefox"

    driver = webdriver.Firefox(options=options, executable_path=DRIVER_PATH)
    driver.set_window_size(1920,1080)

    return driver

def end_current_session(driver):
    driver.quit()

def process_file(file_src = ""):
    if file_src != "":
        files = file_src
    else:
        files = '/home/edgar/Downloads/rica-docs/APAC_Chinese Posts.xlsx'
        
    wb = load_workbook(files)
    sheet = wb.active
    new_content = "Translated Contents"
    new_title = "Translated Title"

    sheet.insert_cols(5,1)
    sheet['E1'] = new_content

    sheet.insert_cols(20,1)
    sheet['T1'] = new_title

    #columns = sheet.max_column # Number of fields
    cells = sheet.max_row # Number of records

    driver = start_new_session()
    driver.get("https://www.google.com/search?q=google+translate")
    ss = driver.find_element_by_css_selector('#tw-source-text-ta')
    ss.clear()	

    words = []
    for i in range(2, cells):
        translated_content = ''
        words.clear()
        # Contents
        try:
            word_len = SplitNumber(sheet['D'+str(i)].value)
            if word_len > 300:
                words = SplitWords(sheet['D'+str(i)].value)
                for word in words:
                    ss = driver.find_element_by_css_selector('#tw-source-text-ta').send_keys(word)
                    time.sleep(2)
                    translated_content += ' ' + str(driver.find_element_by_css_selector('#tw-target-text').text)
                    driver.find_element_by_css_selector('#tw-source-text-ta').clear()	

                sheet['E'+str(i)] = translated_content
                ss = driver.find_element_by_css_selector('#tw-source-text-ta').send_keys(sheet['S'+str(i)].value)
                time.sleep(2)
                translated_content = str(driver.find_element_by_css_selector('#tw-target-text').text)
                sheet['T'+str(i)] = translated_content
                driver.find_element_by_css_selector('#tw-source-text-ta').clear()	
                    
            else:                
                ss = driver.find_element_by_css_selector('#tw-source-text-ta').send_keys(sheet['D'+str(i)].value)
                time.sleep(2)
                translated_content = str(driver.find_element_by_css_selector('#tw-target-text').text)
                sheet['E'+str(i)] = translated_content
                driver.find_element_by_css_selector('#tw-source-text-ta').clear()	

                ss = driver.find_element_by_css_selector('#tw-source-text-ta').send_keys(sheet['S'+str(i)].value)
                time.sleep(2)
                translated_content = str(driver.find_element_by_css_selector('#tw-target-text').text)
                sheet['T'+str(i)] = translated_content
                driver.find_element_by_css_selector('#tw-source-text-ta').clear()	

            print(f'Row : {i}')
            
            #if i == 200:
                #wb.save('/home/edgar/Downloads/rica-docs/APAC_Chinese Posts_Translated.xlsx')
                #break    

        except BaseException as e:
            print("ERROR: "+str(e))
            #end_current_session(driver)
            #wb.save('/home/edgar/Downloads/rica-docs/APAC_Chinese Posts_Translated.xlsx')
            #process_file(s)
            continue

    file_name = os.path.basename(os.path.splitext(file_src)[0])
    file_ext = os.path.basename(os.path.splitext(file_src)[1])
    wb.save(file_name + '_Translated_File' + file_ext)

def SplitNumber(src):
    try:
        
        txt_len = src.split(' ')
        wlen = 0
        for word in txt_len:
            wlen += len(word)

        return wlen
    
    except BaseException as e:
        print("ERROR: "+str(e))
        

def SplitWords(src):
    try:
        
        txt_len = src.split(' ')
        wlen = 0
        words = ''
        wordsArray = []
        for word in txt_len:
            wlen += len(word)
            words += ''.join(word)
            if wlen > 300:
                wordsArray.append(words)
                words = ""
                wlen = 0
        return wordsArray
    
    except BaseException as e:
        print("ERROR: "+str(e))
    

    ## Row = D (content) R (title)
    #driver = start_new_session()
    #driver.get("https://www.google.com/search?q=google+translate")
    #ss = driver.find_element_by_css_selector('#tw-source-text-ta')
    #try:
        ##df = pd.read_csv(files,
                        ##header=None,
                        ##names=('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE'),
                        ##compression='infer',
                        ##error_bad_lines=False, warn_bad_lines=True)

        #df = pd.read_csv(files)

        #row = 1
        #ws = wb.Book(files)
        #while row <= len(df['B']):
            #print(row)
            #if str(df['E'][row]) == 'nan':
                    #ss.clear()
                    #ss = driver.find_element_by_css_selector('#tw-source-text-ta').send_keys(s)
                    ## ss.send_keys(df['D'][row])
                    #os.system("echo %s| clip" % str(df['D'][row]))
                    #ss.send_keys(Keys.CONTROL, 'v')
                    #time.sleep(2)
                    #ss.click()
                    #trsltd_c = str(driver.find_element_by_css_selector('#tw-target-text').text)
                    ## print(trsltd)
                    #ss.clear()
                    #ss = driver.find_element_by_css_selector('#tw-source-text-ta')
                    ## ss.send_keys(df['D'][row])
                    #os.system("echo %s| clip" % str(df['S'][row]))
                    #ss.send_keys(Keys.CONTROL, 'v')
                    #time.sleep(2)
                    #ss.click()
                    #trsltd_t = str(driver.find_element_by_css_selector('#tw-target-text').text)
                    #if trsltd_c == '':
                            #trsltd_c = '.'
                    #if trsltd_t == '':
                            #trsltd_t = '.'
                    #row = row + 1
                    #wb.sheets['APAC_Chinese_Posts'+s].range((row,5)).value = trsltd_c
                    #wb.sheets['APAC_Chinese_Posts'+s].range((row,20)).value = trsltd_t
                    #wb.save()
            #else:
                    #row = row + 1
    #except BaseException as e:
        #print("ERROR: "+str(e))
        #end_current_session(driver)
        ##process_file(s)
        #pass


if __name__ == "__main__":   
    # with arguments
    if len(sys.argv) == 2:
        process_file(sys.argv[1])
        exit(1)
    
    if len(sys.argv) >=3:
        for file in sys.argv[1:]:
            process_file(file)
        
        exit(1)
    # without argument    
    process_file()
    
