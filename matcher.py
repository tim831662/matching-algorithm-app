import pandas as pd
import random
import os

# Configuration
INPUT_FILENAME = 'students_survey.csv'
OUTPUT_SCHEDULE = 'Final_Presentation_Schedule.xlsx'
OUTPUT_PRESS_POOL = 'Final_Press_Pool.xlsx'
STUDENTS_PER_GROUP = 2
GROUPS_PER_DATE = 2
MAX_STUDENTS_PER_DATE = STUDENTS_PER_GROUP * GROUPS_PER_DATE
REVIEWS_PER_STUDENT = 2

available_dates = [
    "9/8/2025", "9/10/2025", "9/15/2025", "9/17/2025", "9/22/2025",
    "9/24/2025", "9/29/2025", "10/1/2025", "10/6/2025", "10/8/2025",
    "10/13/2025", "10/15/2025", "10/20/2025", "10/22/2025", "10/27/2025",
    "10/29/2025", "11/3/2025", "11/5/2025", "11/10/2025", "11/12/2025",
    "11/17/2025", "11/19/2025", "12/1/2025", "12/3/2025", "12/8/2025"
]