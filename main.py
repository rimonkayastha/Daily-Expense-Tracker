#Libraries
from datetime import date
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
  crntfile = openpyxl.workbook.load(path)
  print('The existing excel sheet will now be modified.')
  file_exist = True
  return crntfile

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
headertitles = ['Time', 'Description', 'Category', 'Income', 'Expense']
if file_exist is False:
  for i in range(5):
    ws.cell(row=1, column = i+1, value = headertitles[i])
  currentfile.save(filepath)