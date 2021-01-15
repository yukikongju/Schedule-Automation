#!/usr/bin/python

from openpyxl import Workbook

class Spreadsheet:
    #  def __init__(self, session_name, material_path, courses_lectures_list,
            #  courses_names):
    def __init__(self, session_name, courses):
        self.session_name = session_name
        self.courses = courses
        #  self.courses_lectures_list = courses_lectures_list # deprecated
        #  self.courses_names = courses_names # deprecated
        self.wb = Workbook()                 
        self.generate_spreadsheet()
        self.save_spreadsheet()

    def generate_spreadsheet(self):
        pass

    def save_spreadsheet(self):
        pass

