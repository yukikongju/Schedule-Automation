#!/usr/bin/python

""" Generate Active Recall Spreadsheet for each Lectures

    1. Get list of lectures for each courses
    2. Generate header for each lectures and give some space in between

"""
import glob
import openpyxl
import os

from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.formatting.rule import Rule
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter

GREY = "c1c1c1"
GREEN = ""

def generate_active_recall_spreadsheet(courses_lectures_list, courses_names,
        session_name):
    """ Generate Spreadsheet for active recall

    """
    wb = Workbook()

    # generate big picture tab from first tab
    starting_col = 1
    row_line = 3
    col_question_width = 60
    ws = wb.active
    ws.title = "Lectures Recall"
    #  ws.column_dimensions['A'].width = col_question_width
    ws.column_dimensions[get_column_letter(starting_col)].width = col_question_width
    for j, course in enumerate(courses_lectures_list):
        # get course_name
        course_name = courses_names[j]

        # TODO: make course header
        row_line += 1 # offset
        ws.cell(column = starting_col, row = row_line).\
                value = f"{course_name}"
        column_letter = get_column_letter(starting_col)
        ws[f'{column_letter}{row_line}'].fill = PatternFill(
                fgColor = GREY, fill_type = "solid")
        ws.merge_cells(f'{column_letter}{row_line}:T{row_line}')
        row_line += 2 # offset of 1

        # TODO: add all lecture for course
        for i, lecture in enumerate(course):
            ws.cell(column = starting_col, row = row_line).\
                    value = f"{lecture}"
            row_line += 1

    # generate lectures tab for each course
    for j, course in enumerate(courses_lectures_list):
        # create the worksheet
        course_name = courses_names[j]
        ws = wb.create_sheet(course_name)

        # init the variables
        starting_col = 1
        row_line = 4 # starting at line 4
        ws.column_dimensions[get_column_letter(starting_col)].width =\
                col_question_width
        #  print(course_name)
        for i, lecture in enumerate(course):
            #  print(lecture)
            ws.cell(column = starting_col, row = row_line).\
                    value = f"{lecture}"
            column_letter = get_column_letter(starting_col)
            ws[f'{column_letter}{row_line}'].fill = PatternFill(
                    fgColor = GREY, fill_type = "solid")
            ws.merge_cells(f'{column_letter}{row_line}:T{row_line}')
            row_line += 12 # offset of 10

    # saving spreadsheet
    spreadsheet_name = f"Active Review - {session_name}.xlsx"
    wb.save(spreadsheet_name)

