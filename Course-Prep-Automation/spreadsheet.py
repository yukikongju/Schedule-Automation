#!/usr/bin/python

from openpyxl import Workbook

class Spreadsheet:
    def __init__(self, session_name, courses):
        self.session_name = session_name
        self.courses = courses
        self.wb = Workbook()                 
        self.generate_spreadsheet()
        self.save_spreadsheet()

    def generate_and_save_spreadsheet(self):
        self.generate_spreadsheet()
        self.save_spreadsheet()

    def generate_spreadsheet(self):
        pass

    def save_spreadsheet(self):
        pass

