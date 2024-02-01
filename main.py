import pandas as pd
import numpy as np
import datetime
import re
from icalendar import Calendar, Event

YEAR = 2024
MONTH = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
EmployeeName = 'Mahshid'#.lower()


input_excel_file = "shifts.xlsx"
df = pd.read_excel(input_excel_file, engine='openpyxl', header=None, sheet_name="Enter Projections")

rows, cols = np.where(df == 'Month:') 

month = df.iloc[rows,cols+1].values

print('Valid Month') if month in MONTH else print('ERROR, month name({month}) is not Valid')



df = pd.read_excel(input_excel_file, engine='openpyxl', header=None, sheet_name="PROSPECT")




m_rows, m_cols = np.where(df == EmployeeName)  
d_rows, d_cols = np.where(df == 'Date')  
n_rows, n_cols = np.where(df == 'Name')  


shift_hours = list()
shift_date = list()
shift_weekday = list()

colleages_name = list()

for i in range(len(m_rows)):

    shift_hours.append(df.iloc[m_rows[i],m_cols[i]+1:m_cols[i]+8].values.tolist())
    shift_date.append(df.iloc[d_rows[i],d_cols[i]+1:d_cols[i]+8].values.tolist())
    shift_weekday.append(df.iloc[n_rows[i],n_cols[i]+1:n_cols[i]+8].values.tolist())


    A = df.iloc[n_rows[i]+1:n_rows[i]+20, n_cols[i]].values.tolist()
    for j in range(7):
        B = df.iloc[n_rows[i]+1:n_rows[i]+20, n_cols[i]+j+1].values.tolist()
        colleages_name.append([str(A[i])+': '+str(B[i]) for i in range(len(A))])



shift_hours = np.array(shift_hours).flatten()
shift_date = np.array(shift_date).flatten()
shift_weekday = np.array(shift_weekday).flatten()

for i in range(len(shift_date)):
    if shift_date[i] < 32:
        shift_date[i] = shift_date[i-1] +1



shift_date = [datetime.datetime.fromordinal(datetime.datetime(1900, 1, 1).toordinal() + excel_date - 2)  for excel_date in shift_date] 



cal = Calendar()

for index, day in enumerate(shift_date):
        
    employee_name = EmployeeName
    # if   day>20 and index < 10:
    #     mon = MONTH.index(month)
    # elif day<10 and index > 20:
    #     mon = MONTH.index(month) + 2
    # else:
    #     mon = MONTH.index(month) + 1
    
    
    # if mon > 12:
    #     mon = mon - 12
    #     YEAR += 1
    # elif mon == 0:
    #     mon = 12
    #     YEAR -= 1

    if shift_hours[index] in ['X', 'XX', 'XXX', 'nan']:
        continue
    
    # breakpoint()
    # Define the regex pattern
    pattern = r'\b(\d+)-(\d+)(?:\s*([a-zA-Z]+)?)'

    # Use re.findall to extract all matching numbers
    start, end, txt = re.findall(pattern, shift_hours[index])[0]

    start, end = int(start), int(end)
    if (end - start) > 0:
        start += 12
        
    delta_t = datetime.timedelta(hours=start-1, minutes=30)
    # date = datetime.datetime(YEAR, mon , day, start-1, 30, 0).strftime("%c")
    date = day + delta_t
    date = date.strftime("%c")
    
    print('ERROR') if date[0:3]!=shift_weekday[index].capitalize() else print

    event = Event()
    event.add('summary', f'{employee_name} - {shift_hours[index]}')
    event.add('dtstart', pd.to_datetime(date))
    event.add('dtend', pd.to_datetime(date) + pd.DateOffset(hours=8.5))
    event.add('description',  "-------".join(colleages_name[index]))
    cal.add_component(event)

    # Write the iCalendar to a file
    with open('output.ics', 'wb') as f:
        f.write(cal.to_ical())

    print(event)

        