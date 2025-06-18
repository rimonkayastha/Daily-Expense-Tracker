#Libraries
from datetime import date
import os
import openpyxl
from openpyxl import Workbook
import openpyxl.workbook

# --Functions--
# Excel File Creation
def file_creation(folder, file):
  crntfile = Workbook()
  crntfile.save(os.path.join(folder, file))
  print('A new excel sheet has been created.')
  return crntfile

# Excel File Appending
def file_append(folder, file):
  crntfile = openpyxl.workbook.load(os.path.join(folder, file))
  print('The existing excel sheet will now be modified.')
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
filepath = f'./{foldername}/{filename}'
if os.path.exists(filepath) and os.path.isfile(filepath):
  currentfile = file_append(foldername, filename) # Append file since file exists
else:
  currentfile = file_creation(foldername, filename) # Create file since file does not exist

# Records