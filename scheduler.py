#Author: Etienne Audet-Cobello

#Imports
import openpyxl
from datetime import date
from datetime import timedelta
import win32com.client as win32

#User settings constants
EMAIL_TIMESHEET_RECIPIENT = "example@example.com"
SCHEDULE_FILE_PATH = "C:\\Path\\To\\Timesheet\\timesheet.xlsx"
SCHEDULE_SHEET_NAME = 'Sheet1'
SCHEDULE_SHEET_WEEK_DATES_BEGINNING_CELL = 'B19'
SCHEDULE_SHEET_WEEK_STARTING_CELL = 'E8'
SCHEDULE_SHEET_WEEK_DATES_ENDING_CELL = 'B25'
SCHEDULE_SHEET_SIGNATURE_DATE_CELL = 'E34'

#Other constants
NUMBER_OF_DAYS_IN_WEEK = 7
TODAY = date.today()

#Formats the date.
def formatDate(date):
    return date.strftime("%m/%d/%Y")

scheduleWorkBook = openpyxl.load_workbook(SCHEDULE_FILE_PATH)

scheduleSheet = scheduleWorkBook.get_sheet_by_name(SCHEDULE_SHEET_NAME)

#Replaces the values from cells B19 to B25 while counting down from today's date.
daysToSubstract = NUMBER_OF_DAYS_IN_WEEK

for rows in scheduleSheet[SCHEDULE_SHEET_WEEK_DATES_BEGINNING_CELL:SCHEDULE_SHEET_WEEK_DATES_ENDING_CELL]:
    for cell in rows:
        day = formatDate((TODAY - timedelta(days=daysToSubstract)))
        daysToSubstract -= 1
        scheduleSheet[cell.coordinate] = day

#Replaces the 'Week Starting' cell by removing 7 days from today's date.
weekStarting = formatDate((TODAY - timedelta(days=NUMBER_OF_DAYS_IN_WEEK)))
scheduleSheet[SCHEDULE_SHEET_WEEK_STARTING_CELL] = weekStarting

#Replaces the date field of the employee signature area by today's date.
scheduleSheet[SCHEDULE_SHEET_SIGNATURE_DATE_CELL] = 'Date: ' + formatDate(TODAY)

#Save all changes to the workbook.
scheduleWorkBook.save(SCHEDULE_FILE_PATH)

#Send email
outlook = win32.Dispatch("outlook.application")
mail = outlook.CreateItem(0)
mail.To = EMAIL_TIMESHEET_RECIPIENT
mail.Subject = "Timesheet - Week of : " + weekStarting

attachment = SCHEDULE_FILE_PATH
mail.Attachments.Add(attachment)

mail.Send()
