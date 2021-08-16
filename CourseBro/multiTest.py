import multiprocessing
from multiprocessing import Pool
import pandas as pd


courses_list = []


class Courses(object):
    def __init__(self,courseName,section,term,tutr,lab,no):
        self.courseName = courseName
        self.section = section
        self.term = term
        self.tutr = tutr
        self.lab = lab
        self.no = no

def process_line(line):
    return "The line is "+ str(line.no)




if __name__ == "__main__":
    multipro = multiprocessing.Pool(12)
    df = pd.read_excel('E:\Coursebro\courseList.xlsx',sheet_name='course_list')

    courses = df.values.tolist()
    for c in courses:
        courses_list.append(Courses(*c))
    results = multipro.map(process_line,courses_list,12)

    print(results)