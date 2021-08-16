from selenium import webdriver
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import win32com.client as wincl
from threading import Timer
import pandas as pd
import numpy as np
import multiprocessing
from multiprocessing import SimpleQueue
import functools


courses_list = []

class Courses(object):
    def __init__(self,courseName,section,term,tutr,lab):
        self.courseName = courseName
        self.section = section
        self.term = term
        self.tutr = tutr
        self.lab = lab


def stage(course):
    #web drive initialization

    #Options
    chrome_option = Options()
    #chrome_option.add_argument('--headless')
    #chrome_option.add_argument('--disable-gpu')
    #chrome_option.add_argument('log-level = 3')
    chrome_option.add_argument('window-size=1920x1080')
    #chrome_option.add_experimental_option('excludeSwitches', ['enable-logging'])
    #chrome_option.add_argument("user-data-dir=selenium")
    chrome_option.add_argument("user-data-dir=C:\\Users\\yin00036\\AppData\\Local\\Google\\Chrome\\User Data\\Default") #using stored google profile as chrome cookies


    #url = df.url
    #num_course = df.num_course
    url = 'https://schedulebuilder.yorku.ca'
    driver = webdriver.Chrome(executable_path="E:\\Coursebro\\chromedriver.exe", options = chrome_option)
    driver.get(url)
    driver.maximize_window()

    #dataFrame initialization
    # df = pd.read_excel('E:\Coursebro\Coursebro_v2\CourseBro\courseList.xls',sheet_name='course_list')
    # df = df.dropna(axis = 0, how = 'all').reset_index(drop=True)
    # df = df.fillna(np.nan)#NaN character

    #courses = df.values.tolist()

    # for c in courses:
    #     courses_list.append(Courses(*c))

    

    
    logInSystem(driver)
    selectCourse(driver,course)

    #if driver.current_url == url:
        
    



def logInSystem(driver):
    PassportUsername = "zhx1997"
    PassportPassword = "Niyoyo1201!!"
    wait = WebDriverWait(driver,20)
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
    time.sleep(5)

def selectCourse(driver,course):
    #select fall/winter session
    wait = WebDriverWait(driver,20)
    sessionbox = wait.until(lambda driver: driver.find_element_by_id("term_2021102119"))
    sessionbox.click()

    i = 0
    #enters while loop
    # while i < len(courses_list):
    #     if i != 0 and courses_list[i].courseName == courses_list[i-1].courseName:
    #         i+=1
    #         continue 
        #enter course
    enterFiled =wait.until(lambda driver: driver.find_element_by_id("code_number"))
    enterFiled.send_keys(course.courseName)
    addCourseButton = wait.until(lambda driver: driver.find_element_by_id("addCourseButton"))
    addCourseButton.click()
        # try specific classes
        # selectSpecific = Select(driver.find_element_by_xpath("/html/body/div[1]/div[3]/div[4]/div/table/tbody/tr/td[1]/div[4]/div[10]/div[3]/div[2]/div[11]/select"))
        # selectSpecific.select_by_value("ss")
        # select none
        # selectNone = wait.until(lambda driver: driver.find_element_by_xpath("/html/body/div[1]/div[3]/div[4]/div/table/tbody/tr/td[1]/div[4]/div[10]/div[3]/div[2]/div[13]/span/div/a[2]"))
        # selectNone.click()
        #fill in current term,sec,tutr
    hideFullBox = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,"/html/body/div[1]/div[3]/div[4]/div/table/tbody/tr/td[2]/div[5]/div[3]/div[3]/div[4]/table/tbody/tr[1]/td/div[2]/label[1]/input")))
    if hideFullBox.is_selected() == True:
        
        action = ActionChains(driver)
        action.move_to_element(hideFullBox).click().perform()

    resultMax = driver.find_element_by_class_name("resultMax")
    numOfAval = int(resultMax.get_attribute('innerText'))
    result = 0
    avalSect = []
    while result < numOfAval:
        sectName = wait.until(lambda driver:driver.find_element_by_xpath("/html/body/div[1]/div[3]/div[4]/div/table/tbody/tr/td[2]/div[5]/div[3]/div[4]/div[2]/div[1]/div[1]/div/div/div/label/div/div[1]/table/tbody/tr[1]/td[2]/strong")).get_attribute('innerText')
        avalSect.append(sectName)
        nextButton = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,"/html/body/div[1]/div[3]/div[4]/div/table/tbody/tr/td[2]/div[5]/div[3]/div[3]/div[1]/table/tbody/tr[1]/td[4]/a")))
        nextButton.click()
        result += 1
        
    removeButton = wait.until(lambda driver: driver.find_element_by_xpath("/html/body/div[1]/div[3]/div[4]/div/table/tbody/tr/td[1]/div[4]/div[10]/div[3]/div[2]/div[2]/a/img"))
    removeButton.click()

    # for those in avalSect:
    #         print(those)
    return avalSect


        

    time.sleep(1)

        
        
        #while i+1.name equals to i name
            #i += 1
            #keep selecting
        #i += 1


def main():


    return 0









if (__name__ == "__main__"):
    multipro = multiprocessing.Pool(12)
    df = pd.read_excel('E:\Coursebro\courseList.xlsx',sheet_name='course_list')
    df = df.dropna(axis = 0, how = 'all').reset_index(drop=True)
    df = df.fillna(np.nan)#NaN character
    courses = df.values.tolist()

    for c in courses:
        courses_list.append(Courses(*c))
    results = multipro.map(stage,courses_list,12)
    multipro.close()
    multipro.join()

    print(results)
    