"""
Automatisk ifyllning av Fredriks spärrschema
"""
import os
import csv
import datetime
import regex as re
import pandas as pd

OFFSET_MINUTES = 15 # typ tiden det tar att sig till och från skolan från en tidigare aktivitet
CALENDAR_FILE_NAME = 'calendar.ics'
SCHEMA_FILE_NAME = 'schema.xlsx'

with open(os.path.join('files',CALENDAR_FILE_NAME),'r',encoding='utf8') as calendar_file:
    calendar_contents = calendar_file.read()

out_file = open(os.path.join('files','out.txt'),'w',encoding='utf8')

xlsx_file = pd.read_excel(os.path.join("files",SCHEMA_FILE_NAME))
xlsx_file.to_csv(os.path.join("files","schema.csv"),
                  index = None,
                  header= False,
                  encoding='utf8')

with open(os.path.join('files','schema.csv'), 'r',newline='',encoding='utf8') as csv_file:
    reader = csv.reader(csv_file, delimiter=',')
    # throw away rows before table content
    for row in reader:
        if 'Kan INTE handleda' in row and 'Veckodag' in row:
            break

    out_file.write('Kan INTE handleda\n')
    for row in reader:
        date,start,stop = row[3],row[4],row[6]
        shift_start = (datetime.datetime.strptime(f"{date}T{start}",'%Y-%m-%dT%H:%M')
            - datetime.timedelta(minutes=OFFSET_MINUTES))

        shift_end = (datetime.datetime.strptime(f"{date}T{stop}",'%Y-%m-%dT%H:%M')
            + datetime.timedelta(minutes=OFFSET_MINUTES))

        year,month,day = date.split('-')
        re_start = r"(?<=" + re.escape(f"DTSTART:{year}{month}{day}T") +r")[\S]*.(?=Z)"
        re_end = r"(?<=" + re.escape(f"DTEND:{year}{month}{day}T") +r")[\S]*.(?=Z)"
        found_start = re.findall(re_start,calendar_contents)
        found_end = re.findall(re_end,calendar_contents)

        UTC_OFFSET = datetime.timedelta(hours=2) # ICS timestamps in UTC
        all_times = zip(found_start,found_end)

        for time in all_times:
            existing_event_start = datetime.datetime.strptime(f"{year}{month}{day}T{time[0]}",'%Y%m%dT%H%M%S') + UTC_OFFSET
            existing_event_finish = datetime.datetime.strptime(f"{year}{month}{day}T{time[0]}",'%Y%m%dT%H%M%S') + UTC_OFFSET

            if (existing_event_finish > shift_start) and (existing_event_start < shift_end):
                print(f"Hittat schemakrock på {year}{month}{day}{time}!")
                out_file.write("X\n")
                break
        else:
            out_file.write('\n')

out_file.close()
print(f"Saved results into {os.path.join('files','out.txt')}")
