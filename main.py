#Libraries
from datetime import date
from datetime import time
from datetime import datetime
import os
import openpyxl
from openpyxl import Workbook
import openpyxl.workbook

# --Functions--
# Excel File Creation
def file_creation(path):
  crntfile = Workbook()
  crntfile.save(path)
  print('A new excel sheet has been created.')
  return crntfile

# Excel File Appending
def file_append(path):
  crntfile = openpyxl.load_workbook(path)
  print('The existing excel sheet will now be modified.')
  file_exist = True
  return crntfile

# Add a record
def add_record():
  input_list = []
  time_input_dec = input('Would you like to enter time of income/expense (T) or use Current Time (C): ')
  if time_input_dec.upper() == 'T': #Inputting time
    time_input = input('Enter the time of income/expense (HH:MM): ')
  elif time_input_dec.upper() == 'C':
    current_datetime = datetime.now()
    time_input = current_datetime.strftime("%H:%M")
  input_list.append(time_input)
  for i in range(len(headertitles)-2):
    user_input = input(f'Enter the {headertitles[i+1]} if applicable: ')
    input_list.append(user_input)
  input_tuple = (input_list)
  ws.append(input_tuple)
  currentfile.save(filepath)

#Title
Titlestr = 'DAILY EXPENSE TRACKER'
trgtdate = date.today()
print(f"\n{Titlestr:-^70} \nToday's Date: {trgtdate}")

#Optional Date Input
diffdate = input("Would you like to track expenses for a different date (Y/N)? ")
if diffdate.upper() == 'Y':
  trgtdate = input('Enter target date in YYYY-MM-DD format (Include \'-\'): ')

# File Existence Check
filename = f'{trgtdate}-expenses.xlsx'
foldername = 'Daily-Income-Expense-Sheets'
filepath = os.path.join(foldername, filename)
file_exist = False
if os.path.exists(filepath) and os.path.isfile(filepath):
  currentfile = file_append(filepath) # Append file since file exists
else:
  currentfile = file_creation(filepath) # Create file since file does not exist

# Default Header Writing
ws = currentfile.active
headertitles = ['Time', 'Description', 'Category', 'Income', 'Expense', 'Balance']
if file_exist is False:
  for i in range(6):
    ws.cell(row=1, column = i+1, value = headertitles[i])
  currentfile.save(filepath)
  
# Record
print(ws.max_row)
recordask = input('Would you like to enter a new record (Y/N)? ')
while recordask.upper() == 'Y':
  add_record()
  if ws.max_row==2:
    ws.cell(row=ws.max_row, column=6, value=(int(ws.cell(row=ws.max_row, column=4).value) - int(ws.cell(row=ws.max_row, column=5).value)))
  else:
    ws.cell(row=ws.max_row, column=6, value=(int(ws.cell(row=ws.max_row, column=4).value) - int(ws.cell(row=ws.max_row, column=5).value) + int(ws.cell(row=ws.max_row - 1, column=6).value)))
  currentfile.save(filepath)
  recordask = input('Would you like to enter a new record (Y/N)? ')


