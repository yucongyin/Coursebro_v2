from selenium import webdriver
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

import win32com.client as wincl
from threading import Timer
import pandas as pd
import numpy as np
import multiprocessing
from multiprocessing import SimpleQueue
import functools
import pickle

echo = wincl.Dispatch("SAPI.SpVoice")
echo.rate = 4 #-1

chrome_option = Options()
chrome_option.add_argument('--headless')
chrome_option.add_argument('--disable-gpu')
chrome_option.add_argument('log-level = 3')
chrome_option.add_argument('window-size=1920x1080')
#chrome_option.add_argument("user-data-dir=selenium")
chrome_option.add_argument("user-data-dir=C:\\Users\\yin00036\\AppData\\Local\\Google\\Chrome\\User Data\\Default") #using stored google profile as chrome cookies


def capture(df):
    url = df.url
    num_course = df.num_course
    driver = webdriver.Chrome(executable_path=r"E:\Coursebro\chromedriver", options = chrome_option)
    driver.get(url)
    driver.maximize_window()
    current_url_str = 'https://schedulebuilder.yorku.ca'
    current_url = driver.current_url 
    
    #load pickle cookies object to the browser drive, if it's the first time and has no cookies, skip and wait for cookies dumping
    try:
        #rb auth read and open file access
        cookies = pickle.load(open("cookies.pkl", "rb"))
        for cookie in cookies:
            driver.add_cookie(cookie)
    except:
        pass



    #func body


    #how to pass security authentications 
    driver.execute_script("var a = prompt('Enter Security Code', 'Pass');document.body.setAttribute('data-id', a)")
    time.sleep(60)  #the prompt will last for 60 seconds for the user to input pass code
    passcode = driver.find_element_by_tag_name('body').get_attribute('data-id')  # get the input

    #under if statement, grab the passcode xpath, use variable passcode and submit form to pass authentication





    #func body ends here

    #dump cookies for next process to use
    pickle.dump( driver.get_cookies() , open("cookies.pkl","wb"))






    
    return 0

def runScrap():
    return 0









if (__name__ == "__main__"):
    main()