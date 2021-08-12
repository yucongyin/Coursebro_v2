from selenium import webdriver
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
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




def capture():
    #Options
    chrome_option = Options()
    #chrome_option.add_argument('--headless')
    #chrome_option.add_argument('--disable-gpu')
    #chrome_option.add_argument('log-level = 3')
    chrome_option.add_argument('window-size=1920x1080')
    #chrome_option.add_experimental_option('excludeSwitches', ['enable-logging'])
    #chrome_option.add_argument("user-data-dir=selenium")
    chrome_option.add_argument("user-data-dir=C:\\Users\\OWNER\\AppData\\Local\\Google\\Chrome\\User Data\\Default") #using stored google profile as chrome cookies


    #url = df.url
    #num_course = df.num_course
    url = 'https://schedulebuilder.yorku.ca'
    driver = webdriver.Chrome(executable_path="E:\\Coursebro\\chromedriver.exe", options = chrome_option)
    driver.get(url)
    driver.maximize_window()
    current_url_str = 'https://schedulebuilder.yorku.ca'
    current_url = driver.current_url 
    
    #load pickle cookies object to the browser drive, if it's the first time and has no cookies, skip and wait for cookies dumping
    # try:
    #     #rb auth read and open file access
    #     cookies = pickle.load(open("cookies.pkl", "rb"))
    #     for cookie in cookies:
    #         driver.add_cookie(cookie)
    # except:
    #     pass
    PassportUsername = "zhx1997"
    PassportPassword = "Niyoyo1201!!"
    wait = WebDriverWait(driver,30)
    catnumber = ""
    accountNameId    = "mli"
    passFieldId      = "password"
    loginButtonXpath = "/html/body/div[3]/div[2]/div[1]/form/div[2]/div[2]/p[2]/input"
    #/html/body/div[3]/div[2]/div[1]/form/div[2]/div[2]/p[2]/input

    AccountNameIdElement = wait.until(lambda driver: driver.find_element_by_id(accountNameId))
    passFieldIdElement   = wait.until(lambda driver: driver.find_element_by_id(passFieldId))

    LoginButtonElement   = wait.until(lambda driver: driver.find_element_by_xpath(loginButtonXpath))
    #AccountNameIdElement.clear()
    AccountNameIdElement.send_keys(PassportUsername)

    #passFieldIdElement.clear()
    passFieldIdElement.send_keys(PassportPassword)
    time.sleep(2)
    LoginButtonElement.click()
    time.sleep(30)


    #func body


    #how to pass security authentications 
    #driver.execute_script("var a = prompt('Enter Security Code', 'Pass');document.body.setAttribute('data-id', a)")
    #time.sleep(60)  #the prompt will last for 60 seconds for the user to input pass code
    #passcode = driver.find_element_by_tag_name('body').get_attribute('data-id')  # get the input

    #under if statement, grab the passcode xpath, use variable passcode and submit form to pass authentication





    #func body ends here

    #dump cookies for next process to use
    #pickle.dump( driver.get_cookies() , open("cookies.pkl","wb"))






    
    #return 0

def runScrap():
    return 0









if (__name__ == "__main__"):
    capture()