#Imports
import openpyxl
import smtplib
from datetime import date
from datetime import timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

#Constants
SCHEDULE_FILE_NAME = "timesheet.xlsx"
SCHEDULE_SHEET_NAME = 'Sheet1'
SCHEDULE_SHEET_WEEK_DATES_BEGINNING_CELL = 'B19'
SCHEDULE_SHEET_WEEK_DATES_ENDING_CELL = 'B25'
SCHEDULE_SHEET_WEEK_STARTING_CELL = 'E8'
SCHEDULE_SHEET_SIGNATURE_DATE_CELL = 'E34'
NUMBER_OF_DAYS_IN_WEEK = 7
TODAY = date.today()

################### Fill Timesheet Data ###################

scheduleWorkBook = openpyxl.load_workbook(SCHEDULE_FILE_NAME)

scheduleSheet = scheduleWorkBook.get_sheet_by_name(SCHEDULE_SHEET_NAME)

#Replaces the values from cells B19 to B25 while counting down from today's date.
daysToSubstract = NUMBER_OF_DAYS_IN_WEEK

for rows in scheduleSheet[SCHEDULE_SHEET_WEEK_DATES_BEGINNING_CELL:SCHEDULE_SHEET_WEEK_DATES_ENDING_CELL]:
    for cell in rows:
        day = (TODAY - timedelta(days=daysToSubstract)).strftime("%m/%d/%Y")
        daysToSubstract -= 1
        scheduleSheet[cell.coordinate] = day

#Replaces the 'Week Starting' cell by removing 7 days from today's date.
weekStarting = (TODAY - timedelta(days=NUMBER_OF_DAYS_IN_WEEK)).strftime("%m/%d/%Y")
scheduleSheet[SCHEDULE_SHEET_WEEK_STARTING_CELL] = weekStarting

#Replaces the date field of the employee signature area by today's date.
scheduleSheet[SCHEDULE_SHEET_SIGNATURE_DATE_CELL] = 'Date: ' + TODAY.strftime("%m/%d/%Y")

#Save all changes to the workbook.
scheduleWorkBook.save(SCHEDULE_FILE_NAME)