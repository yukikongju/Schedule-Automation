#!/usr/bin/python

""" Generate all Management Documents for a session
    
    Documents to generate:
        - [ ] Schedule - Session 2 (spreadsheet) : 
        - [ ] Active Review - Session 2 (spreadsheet) : 
        - [ ] TPs and Exercices - Session 2 (spreadsheet)
        - [r] Note de Cours - Session 2 (OneNote)

"""

import glob
import openpyxl
import os

from active_recall import generate_active_recall_spreadsheet

download_path = ""
#  material_path = "courses_material"
material_path = "courses_material/winter2020"
session_name = "Winter 2020"

def main():
    # get all lectures for each course
    courses_lectures_list, courses_names = get_courses_material()

    # generate_active_recall_spreadsheet
    generate_active_recall_spreadsheet(courses_lectures_list, courses_names,
            session_name)

    pass

def get_courses_material():
    """ Get a list for all courses

        Implementation:
            - Open all files in the material directory and save courses name as
              one list and the lectures for each courses as another one

        Parameters:
            - courses_names = ['course1', 'course2']
            - courses_lectures_list = [['course1_lecture1'],['course2_lecture1']]
        
        Return: void
        TODO: return courses list

    """
    courses_names = [] # a list of the course name
    courses_lectures_list = [] # a list of the lecture list for all courses

    os.chdir(material_path) # change directory
    for text_doc in glob.glob("*.txt"):
        # get course name
        course_name = text_doc.replace(".txt", "")
        # get lecture list
        lectures_list = []
        with open(text_doc) as f:
            lectures_list = [line.strip() for line in f]
        f.close()
        # put data into dataframe
        courses_names.append(course_name) 
        courses_lectures_list.append([lecture for lecture in lectures_list])
    return courses_lectures_list, courses_names

if __name__ == "__main__":
    main()
