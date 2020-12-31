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

def generate_active_recall_spreadsheet(courses_lectures_list, courses_names,
        session_name):
    """ Generate Spreadsheet for active recall

    """
    wb = Workbook()

    for j, course in enumerate(courses_lectures_list):
        # create the worksheet
        course_name = courses_names[j]
        ws = wb.create_sheet(course_name)

        # init the variables
        starting_col = 1
        row_line = 4 # starting at line 4
        print(course_name)
        for i, lecture in enumerate(course):
            print(lecture)
            ws.cell(column = starting_col, row = row_line).\
                    value = f"{lecture}"
            column_letter = get_column_letter(starting_col)
            ws[f'{column_letter}{row_line}'].fill = PatternFill(
                    fgColor = GREY, fill_type = "solid")
            ws.merge_cells(f'{column_letter}{row_line}:T{row_line}')
            row_line += 10 # offset of 8
    # saving spreadsheet
    spreadsheet_name = f"Active Review - {session_name}.xlsx"
    wb.save(spreadsheet_name)

