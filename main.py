#Libraries
from datetime import date
import os

# --Functions--
# Excel File Creation
def file_creation(folder, file):
  crntfile = open(os.path.join(folder, file), 'x')
  print('A new excel sheet has been created.')
  return crntfile

# Excel File Appending
def file_append(folder, file):
  filecreate = input('Would you like to overwrite (O) or add to the existing file (A)? ')
  if filecreate.upper() == 'O':
    crntfile = open(os.path.join(folder, file), 'w')
    print('The existing excel sheet will now be overwritten.')
  elif filecreate.upper() == 'A':
    crntfile = open(os.path.join(folder, file), 'a')
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
  currentfile = file_append(foldername, filename)
else:
  currentfile = file_creation(foldername, filename)
