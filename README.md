# Schemakrock
## Automatisk ifyllning av Fredriks spärrschema:
1. `pip install -r requirements.txt`.
2. Skapa en mapp 'files' under roten till detta projekt.
3. Spara spärrschemat som xlsx in files-mappen.
4. Spara din kalender som ICS och lägg in i files-mappen.
5. Fyll i konstanterna (`OFFSET_MINUTES`, `CALENDAR_FILE_NAME` och `SCHEMA_FILE_NAME`).
6. Kör programmet.
7. Resultatet finns in i 'files/out.txt' - klistra innehållet i tabellens första kolumn.

## Skapa en ICS från färdig schema
1. `pip install -r requirements.txt`.
2. Skapa en mapp 'files' under roten till detta projekt.
3. Spara det färdiga schemat som '.xlsx' in i 'files'-mappen.
4. Fyll i konstanterna (SCHEMA_FILE_NAME, YOUR_NAME)
5. Kör programmet.
6. Importera handledningsschema.ics till din kalender.
