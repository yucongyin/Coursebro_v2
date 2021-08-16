from selenium import webdriver
import time
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import os
from threading import Timer
import pandas as pd
import sys
import numpy as np
from threading import Thread
import multiprocessing
from multiprocessing import SimpleQueue
import functools
import json
import win32com.client as wincl

#from joblib import Parallel, delayed
#system voice
echo = wincl.Dispatch("SAPI.SpVoice")
echo.rate = 4 #-1

def stage(prof_num):
    chrome_option = Options()
    chrome_option.add_argument('--headless')
    chrome_option.add_argument('--disable-gpu')
    chrome_option.add_argument('log-level = 3')
    chrome_option.add_argument('window-size=1920x1080')
    profile = "user-data-dir=C:\\Users\\czw\\AppData\\Local\\Google\\Chrome\\User Data\\Profile " + str(prof_num)
    chrome_option.add_argument(profile) #using stored google profile as chrome cookies
    browser = webdriver.Chrome(executable_path="D:\\project\\Coursebro_v2\\chromedriver.exe", options=chrome_option)
    return browser


def capture(df):
    url = df.url
    num_course = df.num_course
    browser = webdriver.Chrome(executable_path="D:\\project\\Coursebro_v2\\chromedriver.exe", options=chrome_option)
    browser.get(url)
    browser.maximize_window()
    current_url_str = 'https://passportyork.yorku.ca/ppylogin/ppylogin'
    current_url = browser.current_url

    # java_script = 'var items = {}; for (index = 0; index < arguments[0].attributes.length; ++index) { items[arguments[0].attributes[index].name] = arguments[0].attributes[index].value }; return items;'
    # when it's on login page
    if current_url == current_url_str:
        name = browser.find_element_by_id("mli")
        name.send_keys("zhx1997")
        pwd = browser.find_element_by_id("password")
        pwd.send_keys("Niyoyo1201!!")
        element = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.NAME, "dologin")))
        browser.execute_script("arguments[0].click();", element)  # click on arg[0] which is element with name "dologin"


    ava_list = []

    #     num_cores = 4
    #     if len(num_course) < num_cores:
    #         num_cores = len(num_course)

    #     ava_course_name = Parallel(n_jobs = num_cores)(delayed(traverse_courses)(idx_course, browser) for idx_course in range(num_course))

    java_script = 'var items = {}; for (index = 0; index < arguments[0].attributes.length; ++index) { items[arguments[0].attributes[index].name] = arguments[0].attributes[index].value }; return items;'

    for idx_course in range(num_course):
        course_xpath = "//div[@class = 'courseDiv bc{} bd{}']".format(idx_course + 1, idx_course + 1)
        requirement = WebDriverWait(browser, 60).until(
            EC.presence_of_element_located((By.XPATH, course_xpath)))
        myElem = browser.find_element_by_xpath(course_xpath)
        check_box = myElem.find_element_by_xpath("div[@class = 'dropDownText']").find_element_by_xpath(
            "span[@class = 'class_checkboxes']")
        labels = check_box.find_elements_by_css_selector("label")
        if myElem.text:
            try:
                x = myElem.text.split('\n')[2]
            except:
                continue

            if myElem.text.split('\n')[2] == 'All classes are full':
                continue
            else:
                course_name = myElem.text.split('\n')[1].split(' ')[0]
                for i in range(len(labels)):
                    label = labels[i]
                    idx = browser.execute_script(java_script, label)['for']
                    input_val = label.find_element_by_xpath("//input[@id = '{}']".format(idx))
                    attrs = browser.execute_script(java_script, input_val)

                    if 'checked' in attrs.keys():
                        course_details = label.text
                        if course_details:
                            if '(Full)' not in course_details:
                                # full_course_name = course_name[3:12] + ' ' + course_details[5:7] + ' ' + course_details[17:18] + course_details[-5:-2]
                                full_course_name = course_name[3:12] + ' ' + course_details
                                remind_name = course_name[3:12] + ' Section ' + course_details[16:17]
                                echo.Speak(remind_name)
                                print(full_course_name)
                                ava_list.append(full_course_name)

    return ava_list


# def find_ele_in_label(i, labels, browser, course_name):
#     java_script = 'var items = {}; for (index = 0; index < arguments[0].attributes.length; ++index) { items[arguments[0].attributes[index].name] = arguments[0].attributes[index].value }; return items;'
#     label = labels[i]
#     idx = browser.execute_script(java_script, label)['for']
#     input_val = label.find_element_by_xpath("//input[@id = '{}']".format(idx))
#     attrs = browser.execute_script(java_script, input_val)

#     if 'checked' in attrs.keys():
#         course_details = label.text
#         full_course_name = course_name[3:12] + ' ' + course_details[5:7] + ' ' + course_details[16:18] + course_details[-5:-2]
#         remind_name = course_name[3:12]
#         print(course_details)
#         if '(Full)' not in course_details:
#             print('Yes')
#             echo.Speak(remind_name)
#             print(full_course_name)
#             return full_course_name
#     else:
#         return None

def webscraping():
    t0 = time.time()
    course_file = 'D:/project/Coursebro_v2/CoursebroTest/test.csv'
    # dataframe build
    df = pd.read_csv(course_file, dtype={'num_course': pd.Int64Dtype()})
    df = df.dropna(axis=0, how='all').reset_index(drop=True)
    df = df.fillna(np.nan)  # NaN character
    df = df.drop_duplicates(subset='url', keep='first')
    out_file = 'D:/project/Coursebro_v2/CoursebroTest/courses.txt'

    len_df = df.shape[0]
    items = [df.iloc[i] for i in range(len_df)]

    thread_num = 12
    if len_df < thread_num:
        thread_num = len_df

    multi_pro = multiprocessing.Pool(thread_num)
    ava_list = multi_pro.map(capture, items)
    multi_pro.close()
    multi_pro.join()

    ava_list = [v for v in ava_list if v != []]
    # Check if the available list is empty
    if ava_list:
        # echo.Speak("Available now")
        with open(out_file, 'w') as filehandle:
            filehandle.writelines("%s\n" % c for c in ava_list)
    t1 = time.time()
    print('One process is done: {0}'.format(t1 - t0))


def start_p(que):
    while True:
        webscraping()
        t = Timer(0.1, functools.partial(que.put, (None,))).start()
        que.get()


if __name__ == '__main__':
    que = SimpleQueue()
    start_process = multiprocessing.Process(target=start_p, args=(que,))
    start_process.start()



