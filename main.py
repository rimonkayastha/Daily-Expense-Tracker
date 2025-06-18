#Libraries
from datetime import date
import os

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
filepath = f'./Daily-Income-Expense-Sheets/{filename}'
if os.path.exists(filepath) and os.path.isfile(filepath):
  print('Already Exists!')
else:
  print('Does not Exist')

#Functions