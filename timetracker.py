#!/usr/bin/python

"""
    0. Create the document
        a. Generate its name generation: Time Tracker - [Month Year]
        b. Generate Sheet for all weeks in the month
        c. Automate the file creation with windows tasks automation

    1. Create a legend 
        a. Properties
            - name (string): name of the tag
            - color (string): color code for the tag
            - id (string): ??? an id if we can't recognize the tag by its color
            - description (string): description of the tag
    2. Generate the template with date and time
    3. Create weekly stats with streamlit

"""
import calendar
import datetime
import openpyxl
import os

from openpyxl import Workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
#  from openpyxl.cell import Cell
from openpyxl.utils import get_column_letter

legend = {'Active Recall': '#ff0000', 
          'TPs and Exercices': '#ff0000',
          #  'Slides':'#151',
          #  'Lecture Review':'#134',
          #  'Unanswered Questions':'#346',
          #  'Coding':'#236',
          #  'Sleep':'#345',
          #  'Management': '#456',
          #  'Wasted Time': '#ff0000',
          #  'Chores + Toilettes': '#876543',
          #  'Eating': '#324',
          'Training':'#0999ff'}

days_of_week = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi',
        'Dimanche']

start_day = '' # time at which we start our day 
end_day = '' # time at which we end our time increment

row_title = 4 # gap for the row title
column_width = 25

# path where the file will be saved
path = "C:/Users/emuli/OneDrive - Universite de Montreal/Bac-Maths-Info/Organization Documents/Time Tracker"

#  template_file = "Time Tracker - Template.xlsx"


def main():
    wb = Workbook() # create workbook
    ws = wb.active

    # generate week tab
    for i in range(4):
        ws = wb.create_sheet(f"Week {i+1}")
    
        # Add legend to tab
        ws['K4'] = "Légende"
        for i, tag in enumerate(legend):
            line = 6+i #6 is the cell where the legend begins
            ws[f'K{6+i}'] = f"{tag}"  
            background_color = openpyxl.styles.colors.Color(rgb = legend[tag][1:])
            fill_color = openpyxl.styles.fills.PatternFill(patternType='solid',\
                    fgColor=background_color)
            ws[f'K{6+i}'].fill = fill_color
            
        #  Adjust columm width
        ws.column_dimensions['K'].width = column_width # adjust legend col width
        #  for i in enumerate(days_of_week)
        #  ws.column_dimensions['B'].width = column_width # adjust days column width

        #  Add days of week
        for i, day in enumerate(days_of_week):
            column_index = i+2
            _ = ws.cell(column = column_index, row = row_title, value = day)
            column_letter = get_column_letter(column_index)
            # adjust days column width
            ws.column_dimensions[f'{column_letter}'].width = column_width 

        # Add dates above days of week
        
        # Add timestamp



        
    # rename workbook - Month Year

    # save file
    wb.save(os.path.join(path, 'Time Tracker - [Month Year].xlsx'))









if __name__ == "__main__":
    main()




