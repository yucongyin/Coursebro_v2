#!/usr/bin/env python3

## command: pyinstaller --onefile webscraping.py

from selenium import webdriver
import time
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import win32com.client as wincl
# import os
from threading import Timer
import pandas as pd
import sys
import numpy as np
# from threading import Thread
import multiprocessing
from multiprocessing import SimpleQueue
import functools

echo = wincl.Dispatch("SAPI.SpVoice")
echo.rate = -1

chrome_option = Options()
chrome_option.add_argument('--headless')
chrome_option.add_argument('--disable-gpu')
chrome_option.add_argument('log-level = 3')
chrome_option.add_argument('window-size=1920x1080')

#def capture(df, browser):
def capture(df):
    ## Get data from DataFrame
    course = df.course
    term = df.term
    section = df.section
    tutorial = df.tutorial
    lab = df.lab
    lecture = df.lecture
    seminar = df.seminar
    lan = df.language_classes
    online = df.fully_online
    studio = df.studio
    blend = df.blend
    course_nick = df.course_name

    ## Initialize browser
    url = 'https://schedulebuilder.yorku.ca/vsb/criteria.jsp?access=0&lang=en&tip=1&page=results&scratch=0&term=0&sort=none&filters=iiiiiiii&bbs=&ds=&cams=0_1_2_3_4_5_6_7_8_9&locs=any'
    browser = webdriver.Chrome(executable_path=r"C:/Users/Tom/Desktop/tcl/chromedriver", chrome_options = chrome_option)
    browser.get(url)
    browser.maximize_window()
    name = browser.find_element_by_id("mli")
    name.send_keys("memorays")
    pwd = browser.find_element_by_id("password")
    pwd.send_keys("asd65656")
    
    element=WebDriverWait(browser,10).until(EC.element_to_be_clickable((By.NAME,"dologin")))
    browser.execute_script("arguments[0].click();", element)

	#login_button = browser.find_element_by_name("dologin")
    #login_button.click()
    # time.sleep(3)

    ## Check data in the browser
    if not term or not section:
        sys.exit('Please fill in term or section for course ' + course)  
    # term_click = None   
    # if term.startswith('S'):
    # 	term_click = browser.find_element_by_id('term_2019115117')
    # else:
    # 	term_click = browser.find_element_by_id('term_2020102119')
    # term_click.click()

    term_id = None
    if term.startswith('S'):
    	term_id = 'term_2020115117'
    else:
    	term_id = 'term_2020102119'

    term_click = WebDriverWait(browser,10).until(EC.element_to_be_clickable((By.ID,term_id)))
    browser.execute_script("arguments[0].click();", term_click)

    code_number = browser.find_element_by_id("code_number")
    code_number.send_keys(course)
    code_button = browser.find_element_by_id("addCourseButton")
    code_button.click()
    # browser.implicitly_wait(5)
    myElem = WebDriverWait(browser, 60).until(EC.presence_of_element_located((By.XPATH, "//div[@class = 'courseDiv bc1 bd1']")))
    term_num = browser.find_element_by_xpath("//div[@class = 'courseDiv bc1 bd1']")
    time.sleep(3)
    term_num = Select(term_num.find_element_by_class_name('course_dropdown').find_elements_by_class_name('select_period')[0])
    term_num.select_by_visible_text('Term ' + term)
    section_num = Select(browser.find_elements_by_class_name('select_usn')[1])
    section_num.select_by_visible_text('Section '+ section)
    class_no = Select(browser.find_elements_by_class_name('dropdownSelect')[1])
    class_no.select_by_visible_text('Try specific classes...')
    time.sleep(2)
    #browser.switch_to_window(browser.window_handles[-1])
    class_no = browser.find_elements_by_class_name('class_checkboxes')
    class_text = class_no[1].text
    s = False
    #if np.isnan(tutorial) == False:
    if pd.isna(tutorial) == False:
        prefix = 'Tutr '
        if (int(tutorial/10) == 0) & (tutorial % 10 != 0):
            prefix = 'Tutr 0'
        tutorial_name = prefix + str(tutorial) + ' (Full)'

        if tutorial_name in class_text:
            pass
        else:
            #tutorial_name = 'Tutr 0' + tutorial
            # echo_str = "say [[rate 150]]"+ course_nick + ' is available now'
            # os.system(echo_str)
            # echo_str = course_nick + ' is available now'
            # echo.Speak(echo_str)
            echo.Speak(course_nick)
            print(course + ' ' + course_nick)
            s = True

    #elif np.isnan(lecture) == False:
    elif pd.isna(lecture) == False:
        lecture_name = 'Lect 0' + str(lecture) + ' (Full)'
        if lecture_name in class_text:
            pass
        else:
            # echo_str = course_nick + ' is available now'
            # # os.system(echo_str)
            # echo.Speak(echo_str)
            echo.Speak(course_nick)
            print(course + ' ' + course_nick)
            s = True
            
    #elif np.isnan(lab) == False:
    elif pd.isna(lab) == False:
        prefix = 'Lab '
        if (int(lab/10) == 0) & (lab % 10 != 0):
            prefix = 'Lab 0'
        lab_name = prefix + str(lab) + ' (Full)'
        if lab_name in class_text:
            pass
        else:
            # echo_str = course_nick + ' is available now'
            # #os.system(echo_str)
            # echo.Speak(echo_str)
            echo.Speak(course_nick)
            print(course + ' ' + course_nick)
            s = True
    elif seminar == True:
        seminar_name = 'Term ' + term + ': Section ' + section + ': Semr 01 (Full)'
        if seminar_name in class_text:
            pass
        else:
            # echo_str = course_nick + ' is available now'
            # #os.system(echo_str)
            # echo.Speak(echo_str)
            echo.Speak(course_nick)
            print(course + ' ' + course_nick)
            s = True
    elif lan == True:
        lan_name = 'Term ' + term + ': Section ' + section + ': Language Classes 01 (Full)'
        if lan_name in class_text:
            pass
        else:
            # echo_str = course_nick + ' is available now'
            # #os.system(echo_str)
            # echo.Speak(echo_str)
            echo.Speak(course_nick)
            print(course + ' ' + course_nick)
            s = True
    elif online == True:
        online_name = 'Term ' + term + ': Section ' + section + ': Fully Online 01 (Full)'
        if online_name in class_text:
            pass
        else:
            # echo_str = course_nick + ' is available now'
            # #os.system(echo_str)
            # echo.Speak(echo_str)
            echo.Speak(course_nick)
            print(course + ' ' + course_nick)
            s = True
    elif pd.isna(studio) == False:
        studio_name = 'Studio 0' + str(studio) + ' (Full)'
        if studio_name in class_text:
            pass
        else:
            echo.Speak(course_nick)
            print(course + ' ' + course_nick)
            s = True
    elif pd.isna(blend) == False:
        blend_name = 'Term ' + term + ': Section ' + section + ': Blended Online and Classroom 0' + str(blend) + ' (Full)'
        if blend_name in class_text:
            pass
        else:
            echo.Speak(course_nick)
            print(course + ' ' + course_nick)
            s = True
    else:
        s = False
    #browser.execute_script("window.history.go(-1)")
    browser.quit()
    return s 

    
def webscraping():

    t0 = time.time()
    course_file = 'C:/Users/Tom/Desktop/tcl/final.csv'
    # course_file = 'C:/Users/Tom/Desktop/tcl/test.csv'
    #courses = list(set(pd.read_csv(course_file)['Course']))
    df = pd.read_csv(course_file, dtype = {'lab': pd.Int64Dtype(), 'tutorial': pd.Int64Dtype(), 'lecture': pd.Int64Dtype(), 'studio': pd.Int64Dtype(), 'blend': pd.Int64Dtype()})
    df = df.dropna(axis = 0, how = 'all').reset_index(drop=True)
    df = df.fillna(np.nan)
    df = df.drop_duplicates(subset = 'course', keep = 'first')
    df['term'] = df['term'].str.upper()
    df['course'] = df['course'].str.upper()
    df['section'] = df['section'].str.upper()
    df = df[df.suspend != 'Y']
    out_file = 'C:/Users/Tom/Desktop/tcl/courses.txt'

    len_df = df.shape[0]
    items = [df.iloc[i] for i in range(len_df)]

    # thread_num = multiprocessing.cpu_count()
    thread_num = 12
    if len_df < thread_num:
    	thread_num = len_df

    multi_pro = multiprocessing.Pool(thread_num)
    batch = multi_pro.map(capture, items)
    multi_pro.close()
    multi_pro.join()

    ava_list = []
    for i in range(len_df):
    	if batch[i] == True:
    		ava_list.append(df.iloc[i].course)
    
    ## Check if the available list is empty
    if ava_list:
        #tkinter.messagebox.showinfo('Information','Some courses are available now')
        #os.system('say "Some courses are available now"')
        echo.Speak("Some courses are available now")
        with open(out_file, 'w') as filehandle:
            filehandle.writelines("%s\n" %c for c in ava_list)
    # Timer(5, webscraping).start()
    t1 = time.time()
    print('One process is done: {0}'.format(t1-t0))


def start_p(que):
	while True:
		webscraping()
		t = Timer(5, functools.partial(que.put, (None,))).start()
		que.get()

    
if __name__ == '__main__':
	que = SimpleQueue()
	start_process = multiprocessing.Process(target = start_p, args=(que,))
	start_process.start() 
