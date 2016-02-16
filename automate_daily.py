import win32com.client
import datetime
import bcrypt
import getpass
from sqlalchemy import create_engine

SENDER_ID = 'email of sender goes here'
SUBJECT_ID = 'subject of email goes here'
UPLOAD_DESTINATION = 'file path to where you\'re uploading goes here. in my case sharepoint'

def connect():
  _pass = getpass.getpass('Enter password to connect to database')
  engine = create_engine('postgres://postgres:' + _pass + '@localhost:5432/scottsdb')
  return engine

def authenticate(): #next steps - break up this function more, handle errors, make more user friendly prompts & put in while loop
  engine = connect()
  user = raw_input('Enter your username: ')
  password = getpass.getpass('Enter your WATIAM password: ')
  db_user = engine.execute('SELECT username FROM users WHERE username = %s', (user,)).fetchone()[0]
  db_pass = engine.execute('SELECT password FROM users WHERE username = %s', (user,)).fetchone()[0]
  
  if bcrypt.hashpw(password, str(db_pass)) == str(db_pass):
    print "Login credentials verified. Searching for daily report"
    find_daily_report()
  else:
    print "Password was incorrect. Run the program again with proper login credentials\n\n"

def get_filename():
  '''returns the appropriate file name for the 
    daily report based on the current date as a string'''

  date = str(datetime.date.today())
  return date + ' ' + '-' + ' ' + 'Jobs to be Approved.xls'
  
def get_inbox_mail():
  '''returns an Outlook 2013 inbox object'''

  outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
  inbox_mail =  outlook.GetDefaultFolder(6).Items
  return inbox_mail

def get_correct_date():
  '''returns the correct date with proper formatting to match against
   the email time of creation as a string'''

  d =  str(datetime.date.today())
  formatted_date = d[5:7] + '/' + d[8:] + '/' + d[2:4]
  return formatted_date

def find_daily_report(): #you hardcoded the date in the logic for testing purposes, switch back to correct_date on Mon
  inbox = get_inbox_mail()
  correct_date = get_correct_date()
  found_daily = False

  for msg in inbox:
    
    if msg.Attachments.Count == 1 and str(msg.CreationTime)[:8] == correct_date \
    and str(msg.subject) == SUBJECT_ID and str(msg.Sender) == SENDER_ID:
      print "Found the daily report. Uploading to Sharepoint."
      upload_to_sharepoint(msg)
      found_daily = True
      break
  
  if not found_daily:
    print 'Unable to find the daily report.'
  
def upload_to_sharepoint(msg):
  '''uploads first attachment of an email to Sharepoint folder'''
  
  daily_report = msg.Attachments[0]
  file_name = get_filename()
  daily_report.saveAsFile(UPLOAD_DESTINATION + '\\' + file_name) #save to synceed sharepoint folder

if __name__ == '__main__':
  authenticate()