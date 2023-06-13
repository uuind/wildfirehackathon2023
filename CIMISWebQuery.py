import webbrowser
from openpyxl.reader.excel import load_workbook
import openpyxl
import os
import time

file = "rainfall_negative.xlsx"
appKey = "59ed8951-4948-4662-9265-728137d72e4f"

wb = load_workbook(filename=file)
ws = wb.active

stations = tuple(ws.rows)

for row_index in range(1, len(stations)):
    row = stations[row_index]
    event_id = str(row[0].value)
    ID = str(row[2].value)
    request_earliest = str(row[7].value)
    request_latest = str(row[8].value)
    #http://et.water.ca.gov/api/data?appKey=59ed8951-4948-4662-9265-728137d72e4f&targets=&startDate=&endDate=&dataItems=day-precip&unitOfMeasure=E
    webbrowser.open("http://et.water.ca.gov/api/data?appKey="+appKey+"&targets="+ID+"&startDate="+request_earliest+"&endDate="+request_latest+"&dataItems=day-precip&unitOfMeasure=E")
    time.sleep(3)
    src = "C:/Users/purpl/Downloads/data"
    dst = "C:/Users/purpl/Downloads/" + event_id + ".xml"
    os.rename(src, dst)
    








