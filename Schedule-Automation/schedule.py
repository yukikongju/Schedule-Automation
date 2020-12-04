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
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.formatting.rule import Rule
from openpyxl.utils import get_column_letter
import glob
import os

material_path = "courses_material"
#  courses_files = ['courses.txt']


slides_per_day = 2 # number of slides to go through per day
lectures_per_day = 2 # number of lecture to attend per day
exercices_per_day = 1 # number of exercices to do per day

class_lectures_per_week = 2 # num of lecture for a class in a week
class_exercices_per_week = 1

title_row = 6
start_row = 8
start_col_week = 'A'
end_col_week = 'T'

lecture_col_width = 25
col_title_width = 15

start_index_title = 3 # titles column starts at C

GREY = "696969"
PASTEL = "5f9ea0"
YELLOW = "fdfd96"

col_titles = ['Slides',             # if lecture has been prepared
              'Lecture',            # if lecture has been attended
              'Review'              # if lecture review has been done
              ]

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
            lectures_list = [line.strip() for line in f]
        f.close()

        # adding columns titles
        for i, title in enumerate(col_titles):
            column_index = start_index_title + i
            ws.cell(column = column_index, row = title_row).\
                    value = title   # add title
            column_letter = get_column_letter(column_index)
            ws.column_dimensions[f'{column_letter}'].width =\
                   col_title_width

        # change lecture column size
        ws.column_dimensions['A'].width = lecture_col_width # adjust legend col width

        week_index = 1 # keeping track of the column
        #  row_index = 0 # keeping track of the row
        row_index = start_row # start at gap
        lecture_index = 0
        lecture_col = 1 # column 'A'
        # separe lecture by weeks
        while len(lectures_list) != 0:
            # create new week
            if lecture_index % (class_lectures_per_week) == 0:
                row_index += 1
                # name week
                ws.cell(column = lecture_col, row = row_index).\
                        value = f"Week {week_index}"
                # set background color
                column_letter = get_column_letter(lecture_col)
                ws[f'{column_letter}{row_index}'].fill = PatternFill(
                        fgColor= GREY, fill_type="solid")
                # merge row
                ws.merge_cells(f'{column_letter}{row_index}:T{row_index}')

                # increment
                week_index += 1
                row_index += 1 # skip a line

            # add lecture
            ws.cell(column = lecture_col, row = row_index).\
                    value = lectures_list.pop(0)

            # TODO: add pending value to column
            yellow_fill = PatternFill(bgColor = YELLOW)
            pending_rule_style = DifferentialStyle(fill = yellow_fill)

            for i, _ in enumerate(col_titles):
                column_index = start_index_title + i
                ws.cell(column = column_index, row = row_index).\
                        value = "Pending"  
                column_letter = get_column_letter(column_index)
                pending_rule = Rule(type = "containsText",
                        operator = "containsText", text = "Pending",
                        dxf = pending_rule_style)
                pending_rule.formula =\
                    [f'NOT(ISERROR(SEARCH("Pending",{column_letter}{row_index})))']
                ws.conditional_formatting.add(f'{column_letter}{row_index}',
                        pending_rule)



            # TODO: add conditional formating to make date green

            # increment index
            lecture_index += 1
            row_index += 1

            # put the exercices?

            # 
    #  wb.save(os.path.join(path, 'Schedule - Test.xlsx'))
    wb.save('Schedule - Test.xlsx')




if __name__ == "__main__":
    main()
