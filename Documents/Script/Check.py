from openpyxl import load_workbook
import time

file = 'ZEK - USER - 2019 - Vorlage.xltx.xlsx'
workbook = load_workbook(file)

getTime = time.ctime()
clk = getTime[11:16]
day = getTime[8:10]
yAxis = str(int(day) + 14)

if "Jan" in getTime:
    worksheet = workbook["Jan"]
elif "Feb" in getTime:
    worksheet = workbook["Feb"]
elif "Mar" in getTime:
    worksheet = workbook["Mär"]
elif "Apr" in getTime:
    worksheet = workbook["Apr"]
elif "May" in getTime:
    worksheet = workbook["Mai"]
elif "Jun" in getTime:
    worksheet = workbook["Jun"]
elif "Jul" in getTime:
    worksheet = workbook["Jul"]
else:
    print("Error: no sheet index")

if worksheet['D'+ yAxis].value == None:
    worksheet['D'+ yAxis].value = clk
else:
    worksheet['E' + yAxis].value = clk

workbook.save(file)
