import numpy as np
import pandas as pd
import os
from tika import parser
from datetime import datetime 
from datetime import date

current_year = date.today().year
directory = os.fsencode('dateien')
    
# Iterates through every list in a given directory    
for file in os.listdir(directory):
     filename = os.fsdecode(file)
     if filename.endswith('.pdf'): 
        # Extract content from a -pdf file
        raw_text = parser.from_file(f'dateien\\{filename}')
        raw_list = raw_text['content'].splitlines()

        # Append contents from .pdf file to temporary .txt file
        with open('temp\\kontoauszug.txt', 'a') as f:
            f.write(raw_text['content'])
            f.close()
     else:
         continue

# Types of actions we want to capture from the banking data
typen = ['Abbuchung', 'Dauerauftrag/Terminueberw.', 'Gutschrift', 'Gutschrift/Dauerauftrag', 'Lastschrift', 'Ueberweisung', 'Wertpapiergutschrift', 'Zins/Dividende WP']

# Lists for later use
dates = []
values = []
types = []
description = []

# Iterate through lines of .txt file to parse the lines
with open(f'temp\\kontoauszug.txt') as kontoauszug:
    for line in kontoauszug:
        if any(typ in line for typ in typen):
            try:
                line_contents = line.split()

                date = line_contents[0]
                value = line_contents[-1]
                type_ = line_contents[1]

                descr_raw = line_contents[2:-1]
                descr = " ".join(descr_raw)

                dates.append(datetime.strptime(date, "%d.%m.%Y").date())
                values.append(float(value.replace('.', '').replace(',', '.')))
                types.append(type_)
                description.append(descr)
            except:
                pass

# Delete temporary conents of the .txt file
open(f'temp\\kontoauszug.txt', 'w').close()

# Create pandas dataframe, change date columns from string to datetime
data = pd.DataFrame({'Datum': dates, 'Typ': types, 'Beschreibung': description, 'Betrag': values})
data['Datum'] = pd.to_datetime(data['Datum'], errors='coerce')

# Dict with months in german, used for naming scheme later
months_ger = {
    1: 'Januar', 2: 'Februar', 3: 'März',
    4: 'Arpil', 5: 'Mai', 6: 'Juni',
    7: 'Juli', 8: 'August', 9: 'September',
    10: 'Oktober', 11: 'November', 12: 'Dezember'
}

# Iterate over months and create Excel file for each month
for i in months_ger.keys():
    for j in data['Datum'].dt.month:
        if i == j:
            data_per_month = data[data['Datum'].dt.month == i]
            data_per_month.to_excel(f'excel_output\\{i}_Kontoauszug_{months_ger[i]}_{current_year}.xlsx')
