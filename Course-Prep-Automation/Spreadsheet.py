#!/usr/bin/python

from openpyxl import Workbook

class Spreadsheet:
    def __init__(self, session_name, material_path, courses_lectures_list, courses_names):
        self.session_name = session_name               # name of the session
        self.material_path = material_path     # dir with all classes for the session
        self.courses_lectures_list = courses_lectures_list
        self.courses_names = courses_names
        self.wb = Workbook()                   # initialize the spreadsheet
        self.generate_spreadsheet()
        self.save_spreadsheet()

    def generate_spreadsheet(self):
        pass

    def save_spreadsheet(self):
        pass

