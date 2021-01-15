#!/usr/bin/python

""" Generate all Management Documents for a session
    
    Documents to generate:
        - [x] Schedule - Session 2 (spreadsheet) : 
        - [x] Active Review - Session 2 (spreadsheet) : 
        - [ ] TPs and Exercices - Session 2 (spreadsheet)
        - [r] Note de Cours - Session 2 (OneNote)

"""

import glob
import openpyxl
import os

from recall import ActiveRecall
from schedule import Schedule

from manager import SessionManager

download_path = ""
course_path = "courses_material/session2"
starting_date = "" # TODO: add a date for each week based on starting date
session_name = "Session 2" # TODO: create Session class

def main():
    manager = SessionManager(session_name, course_path)
    manager.generate_spreadsheets()

if __name__ == "__main__":
    main()
