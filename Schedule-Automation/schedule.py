#!/usr/bin/python

"""
    1. Create a tab file separating the courses material by week for all files
       specified
            a. Add pending for all columns
                * [slides] - green if lecture preparation has been done
                * [lecture] - green if lecture has been watched
                * [review] - green if lecture review has been done
            b. Separe material by weeks
                * set a number of slides to go through per week
                * set unavailable day
                * create a merged bar for row with week
                * add a start date and end date for the session
    2. Create a daily schedule from the weekly separation of all tabs
    3. Open Lecture Notes for the day automatically

"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.utils import get_column_letter
import glob
import os

material_path = "courses_material"
courses_files = ['courses.txt']

slides_per_day = 2 # number of slides to go through per day
lectures_per_day = 2 # number of lecture to attend per day
exercices_per_day = 1 # number of exercices to do per day

class_lectures_per_week = 2 # num of lecture for a class in a week
class_exercices_per_week = 1

title_row = 6
start_row = 8
start_col_title = 'A'
end_col_title = 'T'

GREY = openpyxl.styles.colors.Color(rgb = "808080")

def main():
    wb = Workbook() # init workbook
    #  ws = ws.active # get reference for first worksheet

    
    #  print(os.getcwd())
    os.chdir(material_path)
    for text_doc in glob.glob("*.txt"):
        """ Create the tab 
                1. Retrieve the text files with the lecture
                    a. Retrieve everything in the directory?
                2. Separate content by week
                    a. pop lecture from list when it is placed
            
        """
        # create the tab with course name
        course_name = text_doc
        ws = wb.create_sheet(course_name)

        # retrieve course material from txt file
        #  print(text_doc)
        lectures_list = []
        with open(text_doc) as f:
            #  lectures_list = [[line.strip()] for line in f]
            lectures_list = [line.strip() for line in f]
        f.close()

        week_index = 1 # keeping track of the column
        row_index = 0 # keeping track of the row
        lecture_index = 0
        lecture_col = 1 # column 'A'
        # separe lecture by weeks
        while len(lectures_list) != 0:
        #  while lecture_index < len(lectures_list):
            # create new week
            if row_index % class_lectures_per_week == 0:
                row_index += 1
                ws.cell(column = lecture_col, row = start_row + row_index).\
                        value = "NEW WEEK"
                week_index += 1
                row_index += 1

            ws.cell(column = lecture_col, row = start_row + row_index).\
                    value = lectures_list.pop()
            row_index += 1
            lecture_index += 1

            # put the exercices?

            # 
    #  wb.save(os.path.join(path, 'Schedule - Test.xlsx'))
    wb.save('Schedule - Test.xlsx')



        
        




        #  while len(lectures_list) is not 0:
        #      ws.merge_cells(start_row = start_row + i, start_column = 1,
        #              end_column = 22, end_row = start_row + 1)

        #      background_color = openpyxl.styles.colors.Color(rgb = GREY)
        #      fill_color = openpyxl.styles.fills.PatternFill(patternType='solid',\
        #              fgColor=background_color)
        #      #  column_letter = get_column_letter()
        #      ws[f'K{6+i}'].fill = fill_color





    






if __name__ == "__main__":
    main()
