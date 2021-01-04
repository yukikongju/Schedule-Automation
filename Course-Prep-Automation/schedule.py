#!/usr/bin/python

from spreadsheet import Spreadsheet

from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.formatting.rule import Rule
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter

num_slides_per_day = 2 # number of slides to go through per day
num_lectures_per_day = 2 # number of lecture to attend per day
num_exercices_per_day = 1 # number of exercices to do per day

class_lectures_per_week = 2 # num of lecture for a class in a week
class_exercices_per_week = 1

DAYS_OF_WEEK_FRENCH = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi',
        'Dimanche']

LINKS_CELLS = {'Slides Directory':'B1',
        'Notes de Cours One Note': 'B2',
        'Exercices Spreadsheet': 'B3',
        'Table des Matieres': 'B4'
         }

title_row = len(LINKS_CELLS) + 2
start_row = title_row + 2

lecture_col_width = 25
col_title_width = 15

start_index_title = 3 # titles column starts at C

GREY = "c1c1c1"
YELLOW = "ffff33"
GREEN = "adff2f"

col_titles = ['Slides',             # if lecture has been prepared
              'TP Prep',            # if questions for the TPs has been added to TPs Spreasheet
              'Lecture',            # if lecture has been attended
              'Review',             # if lecture review has been done
              'Recall'              # if concept has been added to recall sheet
              ]

class Schedule(Spreadsheet):
    def __init__(self,  *args, **kwargs):
        super(Schedule, self).__init__(*args, **kwargs)

    def save_spreadsheet(self):
        spreadsheet_name = f"Schedule - {self.session_name}.xlsx"
        self.wb.save(spreadsheet_name)

    def generate_spreadsheet(self):
        self.generate_courses_tab()
        self.generate_daily_schedule()
        self.save_spreadsheet()

    def generate_courses_tab(self):
        """ Creating weekly tabs for all courses 
            
            Implementation:
                - From the courses_names and the courses_lectures_list, generate
                  a weekly schedule for all courses by
                        1. Add columns titles
                        2. Schedule lectures by week according to need
                        3. Add conditional formatting for all lectures row

        """
        for j, course_lectures in enumerate(self.courses_lectures_list):
            # create the tab with course name
            course_name = self.courses_names[j] # course name is saved at index 0
            ws = self.wb.create_sheet(course_name)

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
            row_index = start_row # start at gap
            lecture_index = 0
            lecture_col = 1 # column 'A'
            num_tp_per_week = 1

            # separe lecture by weeks
            for lecture in course_lectures:
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
                        value = lecture

                # add conditional formatting
                yellow_fill = PatternFill(bgColor = YELLOW)
                pending_rule_style = DifferentialStyle(fill = yellow_fill)
                green_fill = PatternFill(bgColor = GREEN)
                completed_rule_style = DifferentialStyle(fill = green_fill)

                for i, _ in enumerate(col_titles):

                    # add pending value to column
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

                    # add conditional formating to make date green
                    completed_rule = Rule(type = "containsText",
                            operator = "containsText", text = "*",
                            dxf = completed_rule_style)
                    completed_rule.formula =\
                        [f'NOT(ISERROR(SEARCH("*",{column_letter}{row_index})))']
                    ws.conditional_formatting.add(f'{column_letter}{row_index}',
                            completed_rule)

                # increment index
                lecture_index += 1
                row_index += 1 # skip a line

                # put the exercices?
                if lecture_index % (class_lectures_per_week) == 0:
                    row_index += 1
                    ws.cell(column = lecture_col, row = row_index).\
                            value = f"TP {week_index - 1}"
                    row_index += 1

    def generate_daily_schedule(self):
        """ Generating Daily Schedule from all the course
            
            Notes:
                a. Each lecture has 1 slide preparation and 1 review
                b. Each chapter has one TP associated with it
                c. Some days may be unavailable for learning

            Implementation:
                - Traverse first lecture for all courses before scheduling another
                  one
                - create two temporary list: one for the lectures, one for the
                  slides, and fill the list by popping the lecture form the list

            Parameters:
                - is_weekend_available (bool): True if user can study on weekends

        """
        # setting up schedule parameters
        is_weekend_available = False
        num_tp_per_lecture = 1
        num_slides_preparation_per_lecture = 1
        num_lecture_review_per_lecture = 1

        # create lectures list to watch in order
        ordered_lectures_list = self.get_ordered_lecture_list()

        # create slides for lecture list
        slides_list = ordered_lectures_list.copy()

        # get reference for first worksheet and rename it
        ws = self.wb.worksheets[0]
        ws.title = "Daily Schedule"


        # Create week
        week_count = 1
        row_index = 6 # begin to draw at row 6 
        start_day_column_index = 3 # days of week start at B
        #  while len(ordered_lectures_list) != 0 or len(slides_list) !=0:
        while len(ordered_lectures_list) != 0:
            # create week row
            _week = ws.cell(column = 1, row = row_index, value = f'Week {week_count}')
            ws.merge_cells(start_row = row_index, start_column = 1, end_column
                    = 15, end_row = row_index)
            _week.fill = PatternFill(fgColor = GREY, fill_type= "solid")
            # TODO: ADD date to week
            row_index += 2 # skip a line

            # create week days column
            for i, day in enumerate(DAYS_OF_WEEK_FRENCH):
                column = start_day_column_index + i
                _cell = ws.cell(row = row_index, column = column, value = day)
                ws.column_dimensions[get_column_letter(column)].width =\
                        lecture_col_width
            row_index += 1
            
            # saving row index for backtracking
            start_lecture_index = row_index # to reset row index after each iter
            start_slides_index = start_lecture_index + num_lectures_per_day

            # First Column TITLE name
            # Lecture 1 to n
            for i in range(num_lectures_per_day):
                ws.cell(column = 1, row = row_index, value = f'Lecture {i+1}')
                row_index += 1
            # Slides 1 to n
            for i in range(num_slides_per_day):
                ws.cell(column = 1, row = row_index, value = f'Slide {i+1}')
                row_index += 1


            # Schedule lectures for the week
            lecture_days_indexes = [1,2,3,4,5] # lectures from monday to friday only
            for col, day in enumerate(lecture_days_indexes):
                for j in range(num_lectures_per_day):
                    if len(ordered_lectures_list) != 0 :
                        ws.cell(column = start_day_column_index + col, 
                                row = start_lecture_index + j, 
                                value = ordered_lectures_list.pop(0))
                    else:
                        break 

            # schedule slides for the week
            #  slide_days_indexes = [1,2,3,4,5,7] # samedi is break from slides prep
            slide_days_indexes = [1,2,3,4,5] # slides only on week days
            for col, day in enumerate(slide_days_indexes):
                for j in range(num_slides_per_day):
                    if len(slides_list) != 0 :
                        ws.cell(column = start_day_column_index + col, 
                                row = start_slides_index + j, 
                                value = slides_list.pop(0))
                    else:
                        break 

            # increment variables
            week_count += 1
            row_index += 1

    def get_ordered_lecture_list(self):
        """ Get a list of the order to watch the lecture for all classes
            
            Implementation:
                - Traverse first lecture for all courses before adding the second
                  one
        """
        # initalize the list
        ordered_list = []

        # retrieve the lectures to watch in order
        lecture_index = 0
        while len(self.courses_lectures_list) != 0:
            # add lectures to list until there are none left
            for i, course in enumerate(self.courses_lectures_list):
                if len(course) != 0:
                    ordered_list.append(self.courses_lectures_list[i].pop(0))
                else: # we pop the empty list
                    self.courses_lectures_list.pop(i)
            lecture_index += 1
        return ordered_list

