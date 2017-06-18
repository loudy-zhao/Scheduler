#Author: Etienne Audet-Cobello

#Imports
import openpyxl
import smtplib
from datetime import date
from datetime import timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

#User settings constants
MY_EMAIL = ""
MY_EMAIL_PASSWORD = ""
EMAIL_TIMESHEET_RECIPIENT = ""
SCHEDULE_FILE_NAME = "timesheet.xlsx"
SCHEDULE_SHEET_NAME = 'Sheet1'
SCHEDULE_SHEET_WEEK_DATES_BEGINNING_CELL = 'B19'
SCHEDULE_SHEET_WEEK_DATES_ENDING_CELL = 'B25'
SCHEDULE_SHEET_WEEK_STARTING_CELL = 'E8'
SCHEDULE_SHEET_SIGNATURE_DATE_CELL = 'E34'
EMAIL_PROVIDER = "smtp.gmail.com" # No need to change this.
EMAIL_PROVIDER_PORT = 587 # No need to change this.

#Other constants
NUMBER_OF_DAYS_IN_WEEK = 7
TODAY = date.today()

#Formats the date.
def formatDate(date):
    return date.strftime("%m/%d/%Y")

scheduleWorkBook = openpyxl.load_workbook(SCHEDULE_FILE_NAME)

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
scheduleWorkBook.save(SCHEDULE_FILE_NAME)

#Create email.
msg = MIMEMultipart()
msg['From'] = MY_EMAIL
msg['To'] = EMAIL_TIMESHEET_RECIPIENT
msg['Date'] = formatDate(TODAY)
msg['Subject'] = "Timesheet - Week of : " + weekStarting

part = MIMEBase('application', "octet-stream")
part.set_payload(open(SCHEDULE_FILE_NAME, "rb").read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment; filename=' + SCHEDULE_FILE_NAME)
msg.attach(part)

smtp = smtplib.SMTP(EMAIL_PROVIDER, EMAIL_PROVIDER_PORT)
smtp.starttls()
smtp.login(MY_EMAIL, MY_EMAIL_PASSWORD)
smtp.sendmail(MY_EMAIL, EMAIL_TIMESHEET_RECIPIENT, msg.as_string())
smtp.quit()