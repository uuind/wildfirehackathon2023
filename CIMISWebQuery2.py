import webbrowser
from openpyxl.reader.excel import load_workbook
import openpyxl
import os
import time

file = "negative_query.xlsx"
appKey = "59ed8951-4948-4662-9265-728137d72e4f"

wb = load_workbook(filename=file, data_only=True)
ws = wb.active

stations = tuple(ws.rows)

for row_index in range(1, len(stations)):
    row = stations[row_index]
    station_name = row[0].value
    earliest = row[1].value
    latest = row[2].value
    ID = str(row[3].value)
    if not isinstance(earliest, type(None)) and not isinstance(latest, type(None)):
        for year_index in range(earliest, latest+1):
            earliest_date = "01-01-" + str(year_index)
            if year_index == 2023:
                latest_date = "5-29-" + str(year_index)
            else:
                latest_date = "12-31-" + str(year_index)
            webbrowser.open(
                "http://et.water.ca.gov/api/data?appKey=" + appKey + "&targets=" + ID + "&startDate=" + earliest_date + "&endDate=" + latest_date + "&dataItems=day-precip&unitOfMeasure=E")
            while not os.path.exists("C:/Users/purpl/Downloads/data"):
                time.sleep(1)
            src = "C:/Users/purpl/Downloads/data"
            dst = "C:/Users/purpl/Downloads/" + station_name + str(year_index) + ".xml"
            os.rename(src, dst)
    # http://et.water.ca.gov/api/data?appKey=59ed8951-4948-4662-9265-728137d72e4f&targets=&startDate=&endDate=&dataItems=day-precip&unitOfMeasure=E
