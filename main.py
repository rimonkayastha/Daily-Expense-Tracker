#Libraries
from datetime import date

#Title
Titlestr = 'DAILY EXPENSE TRACKER'
Datetdystr = date.today()
print(f"\n{Titlestr:-^70} \nToday's Date: {Datetdystr}")

#Ask for date
diffdate = input("Would you like to track expenses of a different date (Y/N)? ")
if diffdate.upper() == 'Y':
  trgtdate = input('Enter target date in YYYY-MM-DD format (Include \'-\'): ')

#Functions