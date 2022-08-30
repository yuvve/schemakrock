"""
Skapa en ICS-fil från Fredriks färdiga schema.
"""
import os
import datetime
import csv
import pandas as pd
import numpy as np
from ics import Calendar, Event

SCHEMA_FILE_NAME = "finished_schema.xlsx"
YOUR_NAME = "FIRST_NAME LAST_NAME"

c = Calendar()
out_file = open(os.path.join('files','out.txt'),'w',encoding='utf8')

df = pd.read_excel(os.path.join("files",SCHEMA_FILE_NAME))
np.savetxt(os.path.join('files','finished_schema.csv'), df, delimiter=';',fmt='%s',encoding='utf8')

with open(os.path.join('files','finished_schema.csv'), 'r',newline='',encoding='utf8') as csv_file:
    reader = csv.reader(csv_file, delimiter=';')
    # throw away rows before table content
    for row in reader:
        if 'Tid' in row and 'Kurs' in row and 'Sal' in row:
            for i,name in enumerate(row):
                if name.lower()==YOUR_NAME.lower():
                    name_column = i
                    break
            else:
                print(f"Couldn't find {YOUR_NAME}!")
            break

    for row in reader:
        tid,kurs,sal,entry = row[0],row[1],row[2],row[name_column]
        start_time = (
            datetime.datetime.strptime(f"{tid[:10]}T{tid[12:17]}",'%Y-%m-%dT%H:%M')
            - datetime.timedelta(hours = 2)
        )

        end_time = (
            datetime.datetime.strptime(f"{tid[:10]}T{tid[18:]}",'%Y-%m-%dT%H:%M')
            - datetime.timedelta(hours = 2)
        )

        if entry != 'nan':
            e = Event()
            e.name = f"Handledning {kurs}"
            e.location = sal
            e.begin = start_time
            e.end = end_time
            c.events.add(e)

with open(os.path.join('files','handledningsschema.ics'), 'w',encoding='utf8') as my_file:
    my_file.writelines(c.serialize_iter())
