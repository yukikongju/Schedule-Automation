#!/usr/bin/python

class DaysOfWeek(object):
    DAYS_OF_WEEK_FRENCH = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi',
        'Dimanche']

class Color(object):
    GREY = "c1c1c1"
    YELLOW = "ffff33"
    GREEN = "adff2f"

class SpreadsheetParameter(object):
    TITLE_ROW = 4
    STARTING_ROW = TITLE_ROW + 2
    LECTURE_COL_WIDTH = 25
    TITLE_COL_WIDTH = 15
    QUESTION_COL_WIDTH = 60
    STARTING_COL_INDEX_TITLE = 3 # TITLES COLUMN STARTS AT c

class WeekParameter(object):
    NUM_SLIDES_PER_DAY = 2 
    NUM_LECTURES_PER_DAY = 2
    NUM_EXERCICES_PER_DAY = 1 
    NUM_LECTURES_PER_CLASS_PER_WEEK = 2
    NUM_EXERCICES_PER_CLASS_PER_WEEK = 1

class ScheduleParameter(object):
    COL_TITLES = ['Slides',             # if lecture has been prepared
                  'TP Prep',            # if questions for the TPs has been added to TPs Spreasheet
                  'Lecture',            # if lecture has been attended
                  'Review',             # if lecture review has been done
                  'Recall'              # if concept has been added to recall sheet
                  ]
