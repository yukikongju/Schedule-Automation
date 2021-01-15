#!/usr/bin/python

import glob
import os
from course import Course

class SessionManager:
    def __init__(self, name, course_path):
        self.name = name
        self.course_path = course_path
        self.courses = self.fetch_courses()
        print(self.courses)
        #  self.generate_spreadsheets

    def fetch_courses(self):
        """ fetch all courses from .txt files in path 
        
            Return: list of courses object
        """
        def fetch_course_name(txt_file):
            """ Return course name """
            return txt_file.replace(".txt", "")

        def fetch_and_generate_course(txt_file):
            """ Return Course object """
            course_name = fetch_course_name(txt_file)
            lectures_list = []
            with open(txt_file) as f:
                lectures_list = [line.strip() for line in f]
            f.close()
            return Course(self.name, lectures_list)

        courses = []
        os.chdir(self.course_path) # change directory
        for txt_file in glob.glob("*.txt"):
            course = fetch_and_generate_course(txt_file)
            courses.append(course)
        return courses

    def generate_spreadsheets(self):
        # generate_active_recall_spreadsheet
        #  ar_spreadsheet = ActiveRecall(session_name, material_path,
        #          courses_lectures_list, courses_names)

        # generate schedule spreadsheet
        #  schedule_spreadsheet = Schedule(session_name, material_path,
        #          courses_lectures_list, courses_names)
        pass
